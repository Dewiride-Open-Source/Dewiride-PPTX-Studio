/**
 * The gate.
 *
 * The plan sets sub-phase 0.5's exit criterion as *parse -> serialize ->
 * byte-identical for 100% of parts*, plus an assertion that the serialized
 * document's prefix map equals the input's. Both are here, and both are weaker
 * than they look, for the same reason: **in a round trip nothing is dirty**, so
 * every node re-emits as a slice and the byte comparison is testing the
 * tokenizer's spans rather than the serializer. `return document.source` passes
 * it.
 *
 * So this module checks the property the gate is a proxy for:
 *
 * > Serialize the document, parse the result, and the two trees must agree on
 * > everything that carries meaning - node order and kind, qualified names,
 * > attribute names, values, order and quoting, text, and the namespace
 * > bindings each prefix has.
 *
 * That holds whether or not anything was edited, which is what makes it useful:
 * mark every node in a real part dirty, force the rebuild path over all of it,
 * and this says whether the rebuild is faithful. The byte comparison cannot,
 * because a faithful rebuild is *not* byte-identical in general - `&#62;` and
 * `>` are the same character data spelled two ways, and we re-emit whichever
 * spelling is minimal rather than whichever one the original producer chose.
 *
 * Comparing qualified names at every use site also subsumes the plan's
 * prefix-map assertion rather than merely restating it. The declaration map
 * alone cannot see a rewrite that renames a prefix consistently at both its
 * binding and its uses - and that is precisely the rewrite `XMLSerializer` is
 * permitted to perform, and precisely the one that silently turns
 * `mc:Choice/@Requires="a14"` into a hard error in PowerPoint. The map is
 * compared too, because it is cheap and it names the failure clearly when it is
 * the binding that moved.
 */

import { serializeXmlString } from './serialize.js';
import { isXmlError } from '../errors.js';
import {
  parseXmlString,
  prefixMap,
  type XDocument,
  type XNode,
  type XmlParseLimits,
} from '../parse/xnode.js';

/** One way in which a document and its own serialization disagree. */
export interface RoundTripDifference {
  /**
   * `parse` - the output is not well-formed, which is the worst case and the
   * only one that makes the rest of the comparison impossible.
   * `structure` - the trees differ at some node.
   * `prefix` - a prefix is bound to a different set of URIs.
   * `bom` - the byte order mark was gained or lost.
   */
  readonly kind: 'parse' | 'structure' | 'prefix' | 'bom';
  /** Offset in the **input** document, where one is attributable. */
  readonly at: number;
  readonly detail: string;
}

interface Entry {
  readonly at: number;
  readonly key: string;
}

/**
 * A flat, order-preserving description of everything a round trip must not
 * change.
 *
 * Iterative for the same reason everything else in this package is: a part
 * nested thousands deep is well-formed XML that PowerPoint opens.
 *
 * **Adjacent text nodes are one entry.** Two text children in a row serialize to
 * one run and reparse as one node, so comparing them node by node would report a
 * difference where the document has none - XML 1.0 has character data, not a
 * list of text nodes, and the split between two of them is an artefact of how
 * the tree was built rather than anything the file says. A parse never produces
 * that shape; an edit that removes an element from between two whitespace runs
 * does, and that edit is correct. Coalescing here is what lets sub-phase 0.6's
 * `removeChild` be honest about it. Empty runs fall out of the same rule.
 *
 * CDATA is *not* coalesced into a neighbouring run. It is a lexical distinction
 * this package preserves deliberately, and merging it away would hide a CDATA
 * section that came back as ordinary text.
 */
function signature(document: XDocument): Entry[] {
  const entries: Entry[] = [];
  const stack: (XNode | Entry)[] = [];
  for (let i = document.children.length - 1; i >= 0; i--) stack.push(document.children[i]!);

  let pending: { at: number; value: string } | undefined;
  const flush = (): void => {
    if (pending === undefined) return;
    if (pending.value !== '') {
      entries.push({ at: pending.at, key: 'text ' + JSON.stringify(pending.value) });
    }
    pending = undefined;
  };
  const push = (entry: Entry): void => {
    flush();
    entries.push(entry);
  };

  while (stack.length > 0) {
    const work = stack.pop()!;
    if (!('type' in work)) {
      push(work);
      continue;
    }
    switch (work.type) {
      case 'element': {
        let key =
          'element <' + work.qname + (work.selfClosing && work.children.length === 0 ? '/>' : '>');
        for (const attribute of work.attributes) {
          key +=
            ' ' +
            attribute.qname +
            '=' +
            attribute.quote +
            JSON.stringify(attribute.value) +
            attribute.quote;
        }
        push({ at: work.start, key });
        stack.push({ at: work.end, key: 'close </' + work.qname + '>' });
        for (let i = work.children.length - 1; i >= 0; i--) stack.push(work.children[i]!);
        break;
      }
      case 'text':
        if (pending === undefined) pending = { at: work.start, value: work.value };
        else pending.value += work.value;
        break;
      case 'cdata':
        push({ at: work.start, key: 'cdata ' + JSON.stringify(work.value) });
        break;
      case 'comment':
        push({ at: work.start, key: 'comment ' + JSON.stringify(work.value) });
        break;
      case 'processingInstruction':
        push({
          at: work.start,
          key: 'pi ' + work.target + ' ' + JSON.stringify(work.data),
        });
        break;
      case 'declaration':
        push({
          at: work.start,
          key:
            'declaration ' +
            work.version +
            ' ' +
            String(work.encoding) +
            ' ' +
            String(work.standalone),
        });
        break;
    }
  }
  flush();
  return entries;
}

function prefixSignature(document: XDocument): string {
  return [...prefixMap(document)]
    .map(
      ([prefix, uris]) => (prefix === '' ? '(default)' : prefix) + '=' + [...uris].sort().join('|'),
    )
    .sort()
    .join(' ');
}

/** How many differences to report before giving up. The first is usually enough. */
const MAX_REPORTED = 10;

/**
 * Serialize `document`, parse the result, and report every way the two differ.
 *
 * An empty array is the pass. Deliberately returns findings rather than
 * throwing, matching `checkTreeCoverage` and `checkSpanCoverage`: this is a
 * measuring instrument, and the corpus gate wants to count failures across
 * thousands of parts rather than stop at the first.
 */
export function checkRoundTrip(
  document: XDocument,
  limits?: XmlParseLimits,
): RoundTripDifference[] {
  const differences: RoundTripDifference[] = [];
  const serialized = serializeXmlString(document);

  let reparsed: XDocument;
  try {
    reparsed = parseXmlString(serialized, limits);
  } catch (error) {
    return [
      {
        kind: 'parse',
        at: 0,
        detail:
          'the serialized document does not parse: ' +
          (isXmlError(error)
            ? error.code + ' at ' + String(error.detail.offset) + ' - ' + error.message
            : String(error)),
      },
    ];
  }

  if (document.bom !== reparsed.bom) {
    differences.push({
      kind: 'bom',
      at: 0,
      detail: document.bom ? 'the byte order mark was lost' : 'a byte order mark appeared',
    });
  }

  const before = signature(document);
  const after = signature(reparsed);
  for (let i = 0; i < Math.max(before.length, after.length); i++) {
    if (differences.length >= MAX_REPORTED) break;
    const a = before[i];
    const b = after[i];
    if (a !== undefined && b !== undefined && a.key === b.key) continue;
    differences.push({
      kind: 'structure',
      at: a?.at ?? document.source.length,
      detail:
        'node ' +
        i +
        ': wrote ' +
        (a === undefined ? '(nothing)' : a.key) +
        ', read back ' +
        (b === undefined ? '(nothing)' : b.key),
    });
  }

  const prefixesBefore = prefixSignature(document);
  const prefixesAfter = prefixSignature(reparsed);
  if (prefixesBefore !== prefixesAfter) {
    differences.push({
      kind: 'prefix',
      at: document.root.start,
      detail: 'prefix bindings were ' + prefixesBefore + ', are now ' + prefixesAfter,
    });
  }

  return differences;
}

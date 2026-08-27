import { describe, expect, it } from 'vitest';
import { isXmlError } from './errors.js';
import { checkRoundTrip } from './roundtrip.js';
import { serializeXmlString } from './serialize.js';
import { applyEdits, insertInOrder, newElement, type XmlEdit } from './edit.js';
import {
  checkTreeCoverage,
  descendantElements,
  parseXmlString,
  type XDocument,
  type XNode,
} from './xnode.js';

/**
 * Structured mutation of realistic markup.
 *
 * The plan commits Phase 12 to an invariant this package can already be held
 * to: *the parser either returns a document or throws a typed error - never
 * hangs, never over-allocates, never throws a raw `RangeError`.* Waiting for
 * the fuzzing harness to assert it means shipping five sub-phases on top of an
 * unverified promise.
 *
 * Mutation rather than random bytes, deliberately. Random bytes are rejected in
 * the first few characters and never reach the interesting paths; a real slide
 * part with one character changed reaches all of them.
 *
 * The second assertion is the one that could not be reached any other way: a
 * mutation that yields a *valid* document whose spans no longer tile the source
 * would be the worst outcome of all, because no error code would ever report
 * it.
 *
 * Sub-phase 0.5 added two more assertions to the same loop: every mutant that
 * parses must serialize back byte-identically, and must still read back as the
 * same document once every node is forced onto the rebuild path.
 *
 * Sub-phase 0.6 added a third: every mutant that parses is then *edited* -
 * attributes set, added and removed, children inserted in schema order and
 * taken out, text rewritten - and undone, and must come back byte-identical
 * with no dirty flag left behind. The corpus gate runs the same battery over
 * 2834 real parts; this runs it over documents no producer would ever write,
 * and it earned its place immediately: the sweep found that the mixed-content
 * guard could refuse an *inverse*, so a document mangled into holding text
 * beside elements could be edited and then not put back. The guard now applies
 * to forward edits only, and `applyEdits` rolls a refused batch all the way
 * back.
 *
 * This run is a deterministic 4000-case slice of a 40 000-case sweep that
 * produced 2855 parses and 37 145 typed errors - disjoint, so nothing threw
 * after a successful parse - with zero untyped throws, zero coverage violations
 * and zero fidelity failures. 2829 of the parsed documents were genuinely
 * changed by the edit battery, 48 322 edits in all, every one of them undone
 * exactly; 26 batches were refused at apply time and rolled back to the
 * original bytes. Slowest single case 197 ms, which is the 2000-fold repeated
 * open tag being edited rather than merely parsed.
 */

const SEED = 0x2f6e1a3b;

/** Force every node onto the serializer's rebuild path. See `serialize.test.ts`. */
function forceRebuild(document: XDocument): XDocument {
  const stack: XNode[] = [...document.children];
  while (stack.length > 0) {
    const node = stack.pop()!;
    node.dirty = true;
    if (node.type !== 'element') continue;
    for (const attribute of node.attributes) attribute.dirty = true;
    for (const child of node.children) stack.push(child);
  }
  return document;
}

/**
 * A battery of edits spread across one document, the same shape the corpus gate
 * applies to real parts.
 *
 * Applied together rather than one at a time, because the interesting failures
 * are in the interaction: one edit's `markDirty` stopping at an ancestor that
 * another edit already dirtied, and the undo having to unpick exactly its own
 * share of that.
 */
function planEdits(document: XDocument): XmlEdit[] {
  const edits: XmlEdit[] = [];
  for (const element of descendantElements(document.root)) {
    if (element.attributes.length > 0) {
      edits.push({
        kind: 'setAttribute',
        element,
        qname: element.attributes[0]!.qname,
        value: 'edited',
      });
    }
    if (element.attributes.length > 1) {
      edits.push({
        kind: 'removeAttribute',
        element,
        qname: element.attributes[element.attributes.length - 1]!.qname,
      });
    }
    edits.push({ kind: 'setAttribute', element, qname: 'zzFuzz', value: 'x' });

    const firstElementChild = element.children.find((child) => child.type === 'element');
    if (firstElementChild !== undefined) {
      try {
        edits.push(insertInOrder(element, newElement(firstElementChild.qname)));
      } catch (error) {
        // Most of these documents are not OOXML at all, so most parents have no
        // content model. Refusing is the correct answer and not a fuzz finding.
        if (!isXmlError(error) || error.code !== 'ERR_SCHEMA_ORDER') throw error;
      }
    }
    const last = element.children[element.children.length - 1];
    if (last !== undefined) edits.push({ kind: 'removeChild', parent: element, node: last });
    const text = element.children.find((child) => child.type === 'text');
    if (text !== undefined) edits.push({ kind: 'setValue', node: text, value: 'edited text' });
  }
  return edits;
}

function anyDirty(document: XDocument): boolean {
  const stack: XNode[] = [...document.children];
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (node.dirty) return true;
    if (node.type !== 'element') continue;
    for (const attribute of node.attributes) if (attribute.dirty) return true;
    for (const child of node.children) stack.push(child);
  }
  return false;
}

function makeRandom(seed: number): () => number {
  let state = seed;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) / 4294967296;
  };
}

const SLIDE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Rectangle 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>' +
  '<a:lstStyle><a:extLst/></a:lstStyle><a:p><a:r><a:rPr lang="en-US" dirty="0"/>' +
  '<a:t>Hello &amp; welcome — café 日本語</a:t></a:r></a:p></p:txBody></p:sp>' +
  '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';

const SOURCES = [
  SLIDE,
  '<a><!-- c --><?pi d?><![CDATA[x]]>t</a>',
  '<r xmlns="urn:d"><c xmlns=""><g a=\'1\' b="2" /></c></r>',
  '﻿<?xml version="1.0"?>\r\n<a>\r\n  <b/>\r\n</a>\r\n',
  // Character references, and only character references, can put a tab, a line
  // feed or a carriage return into a *value* - written literally, the parser
  // has already normalized them away by the time the serializer sees them. So
  // without this source the escaping rules are unreachable from here: three
  // deliberate breaks to the escapers survived the whole 4000-case run before
  // it existed. `&#62;` and `]]&gt;` are here for the opposite reason, to hold
  // the decision *not* to escape `>` except where §2.4 forces it.
  '<a n="x&#9;y" m="p&#xA;q" o="r&#xD;s"><b>t&#xD;u</b><c>1 &#62; 0 ]]&gt; z</c></a>',
];

const INJECTIONS = [
  '<',
  '>',
  '&',
  '"',
  "'",
  '/',
  '?',
  '!',
  '-',
  ']',
  '[',
  '=',
  ' ',
  '\0',
  '',
  '￿',
  '<!--',
  '-->',
  '<![CDATA[',
  ']]>',
  '<?',
  '?>',
  '&#',
  ';',
  '&#xD800;',
  '&nope;',
  '&#xD;',
  '&#9;',
  '&#62;',
  ']]&gt;',
  '<!DOCTYPE x>',
  'xmlns=',
  '\ud800',
  '\udfff',
];

type Mutator = (source: string, random: () => number) => string;

const MUTATORS: Record<string, Mutator> = {
  truncate: (s, r) => s.slice(0, Math.floor(r() * s.length)),
  chop: (s, r) => s.slice(Math.floor(r() * s.length)),
  inject: (s, r) => {
    const at = Math.floor(r() * s.length);
    return s.slice(0, at) + INJECTIONS[Math.floor(r() * INJECTIONS.length)]! + s.slice(at);
  },
  deleteRun: (s, r) => {
    const at = Math.floor(r() * s.length);
    return s.slice(0, at) + s.slice(at + 1 + Math.floor(r() * 40));
  },
  swapChar: (s, r) => {
    const at = Math.floor(r() * s.length);
    return s.slice(0, at) + INJECTIONS[Math.floor(r() * INJECTIONS.length)]! + s.slice(at + 1);
  },
  duplicateAttribute: (s) => s.replace(/(\s[\w:]+="[^"]*")/, '$1$1'),
  unbalance: (s) => s.replace('</', '<'),
  repeatOpenTag: (s) => {
    const match = /<([\w:]+)([^>/]*)>/.exec(s);
    return match ? s.slice(0, match.index) + match[0].repeat(2000) + s.slice(match.index) : s;
  },
  hugeAttributeCount: (s) =>
    s.replace(
      /<([\w:]+)/,
      (_, name: string) =>
        '<' + name + Array.from({ length: 2000 }, (__, i) => ` a${i}="1"`).join(''),
    ),
  ampersandStorm: (s) => s.replace(/>/g, '>&'),
};

describe('hostile and merely broken input', () => {
  it('never throws anything but an XmlError, and never loses coverage or fidelity', () => {
    const random = makeRandom(SEED);
    const kinds = Object.keys(MUTATORS);
    const untyped: string[] = [];
    const uncovered: string[] = [];
    const unfaithful: string[] = [];
    let parsed = 0;
    let edited = 0;
    let refused = 0;
    let thrown = 0;
    let slowest = 0;

    for (let i = 0; i < 4000; i++) {
      const source = SOURCES[Math.floor(random() * SOURCES.length)]!;
      const kind = kinds[Math.floor(random() * kinds.length)]!;
      const mutated = MUTATORS[kind]!(source, random);

      const started = performance.now();
      try {
        const document = parseXmlString(mutated);
        parsed++;
        const gaps = checkTreeCoverage(document);
        if (gaps.length > 0 && uncovered.length < 3) {
          uncovered.push(kind + ': ' + JSON.stringify(gaps[0]));
        }
        // Sub-phase 0.5's gate, over 4000 documents no producer would write.
        // Nothing is dirty, so every node takes the slice path.
        if (serializeXmlString(document) !== mutated && unfaithful.length < 3) {
          unfaithful.push(kind + ': a clean serialize is not byte-identical');
        }
        // And the half the byte gate cannot reach: mark everything dirty and
        // the rebuild must still read back as the same document. What this
        // reaches is escaping and quoting, on input far stranger than a slide
        // part. What it cannot reach is the empty-element form, because a
        // self-closing element with children only comes from an edit - that one
        // is pinned in serialize.test.ts instead.
        const differences = checkRoundTrip(forceRebuild(document));
        if (differences.length > 0 && unfaithful.length < 3) {
          unfaithful.push(kind + ': ' + JSON.stringify(differences[0]));
        }

        // Sub-phase 0.6's gate. A fresh parse, because the tree above has been
        // forced dirty and could not come back to a slice.
        const editable = parseXmlString(mutated);
        const edits = planEdits(editable);
        if (edits.length > 0) {
          let inverses: XmlEdit[] | undefined;
          try {
            inverses = applyEdits(edits);
          } catch (error) {
            // A batch can be refused at apply time - a mutation that puts text
            // beside elements makes an insertion into mixed content, and this
            // battery cannot know that in advance. The batch must then have
            // rolled all the way back, which is a stronger assertion than the
            // successful path and is why this is not simply skipped.
            if (!isXmlError(error) || error.code !== 'ERR_INVALID_EDIT') throw error;
            refused++;
          }
          if (inverses === undefined) {
            if (serializeXmlString(editable) !== mutated && unfaithful.length < 3) {
              unfaithful.push(kind + ': a refused batch did not roll back');
            }
          } else {
            if (serializeXmlString(editable) !== mutated) edited++;
            applyEdits(inverses);
          }
          if (serializeXmlString(editable) !== mutated && unfaithful.length < 3) {
            unfaithful.push(kind + ': undo did not restore the source');
          }
          if (anyDirty(editable) && unfaithful.length < 3) {
            unfaithful.push(kind + ': undo left a dirty flag behind');
          }
          const afterUndo = checkRoundTrip(editable);
          if (afterUndo.length > 0 && unfaithful.length < 3) {
            unfaithful.push(kind + ': after undo, ' + JSON.stringify(afterUndo[0]));
          }
        }
      } catch (error) {
        thrown++;
        if (!isXmlError(error) && untyped.length < 3) {
          untyped.push(
            kind + ': ' + (error instanceof Error ? error.constructor.name : typeof error),
          );
        }
      }
      slowest = Math.max(slowest, performance.now() - started);
    }

    expect(untyped).toEqual([]);
    expect(uncovered).toEqual([]);
    expect(unfaithful).toEqual([]);
    // Both halves must actually happen, or the run proves nothing: all-throw
    // would mean the mutations are too destructive to exercise the parser, and
    // all-parse would mean they are not destructive at all.
    expect(parsed).toBeGreaterThan(100);
    // And the edit battery must actually have changed something: an addition
    // that silently no-ops is how 0.5's first fuzz extension turned out to be
    // vacuous, and it took mutation testing to notice.
    expect(edited).toBeGreaterThan(100);
    // Some batches are refused, and their rollback is the sharper assertion.
    expect(refused).toBeGreaterThan(0);
    expect(thrown).toBeGreaterThan(100);
    // A 2000-attribute element or a 2000-fold repeated tag must stay linear.
    expect(slowest).toBeLessThan(2000);
  });

  it('refuses the pathological shapes without scanning to the end of the world', () => {
    const started = performance.now();
    for (const nasty of [
      '<a ' + 'b="1" '.repeat(5000) + '/>', // attribute flood
      '<a>' + '&'.repeat(20_000) + '</a>', // bare ampersands
      '<a>' + '<b>'.repeat(20_000), // unclosed nesting
      '<!--' + 'x'.repeat(200_000), // unterminated comment
      '<a n="' + 'x'.repeat(200_000), // unterminated attribute value
      '<![CDATA[' + 'x'.repeat(200_000), // unterminated CDATA
    ]) {
      expect(() => parseXmlString(nasty)).toThrowError();
      try {
        parseXmlString(nasty);
      } catch (error) {
        expect(isXmlError(error)).toBe(true);
      }
    }
    expect(performance.now() - started).toBeLessThan(5000);
  });
});

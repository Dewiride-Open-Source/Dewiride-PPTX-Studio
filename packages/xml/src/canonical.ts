import { XmlError } from './errors.js';
import { NS } from './namespaces.js';
import { escapeAttributeValue, escapeText } from './references.js';
import type { XAttribute } from './tokenizer.js';
import type { XDocument, XElement, XNode } from './xnode.js';

/**
 * One document, one string, and two documents that mean the same thing produce
 * the same string.
 *
 * ## What this is for
 *
 * Sub-phase 1.4 asks whether a deck we wrote back is the same deck we read.
 * Comparing the archives byte for byte answers a different and much weaker
 * question - the plan rules it out and the reason is worth repeating, because
 * a byte-equality round-trip test is red on day one and disabled on day two:
 * entry order, deflate level, DOS timestamps and attribute order all differ
 * between two archives that hold the identical document.
 *
 * So the comparison is made on a normal form. Everything a producer is free to
 * choose is spelled one way here; everything that carries meaning is left
 * exactly as it was found.
 *
 * ## Not W3C Canonical XML, and the three places it departs
 *
 * The name is borrowed and the escaping rules are the same, but C14N was
 * designed so that a signature survives a document being moved between
 * processors. That is not this problem, and two of its rules are actively
 * wrong for OOXML.
 *
 * **Namespace declarations are kept, including the ones nothing uses.** C14N
 * prunes a declaration no name in scope refers to. In a PresentationML part
 * that prune is data loss: PowerPoint writes `xmlns:a14="..."` on the root
 * alongside `mc:Ignorable="a14"`, and there is frequently no `a14:` element
 * anywhere in the part - the declaration exists so that the `mc:Ignorable`
 * token resolves. Delete it and a well-formed part becomes one naming an
 * undeclared prefix.
 *
 * **Prefixes are kept as written, never renamed.** For the same reason, one
 * layer up. `mc:Ignorable` and `mc:Choice/@Requires` hold *prefixes*, not URIs;
 * a canonical form free to rename `a14` to `ns3` would report two documents as
 * identical when one of them has quietly become invalid. This is the whole
 * reason `@pptx-studio/xml` exists rather than a call to `XMLSerializer`, and
 * a canonical form that gave it away here would be undoing that decision at
 * the far end of the pipeline.
 *
 * **The XML declaration is kept.** C14N drops it. Losing one is a real change
 * to a `.pptx` - Office writes it on every part - so it is normalised rather
 * than discarded: the version and `standalone` survive, the encoding does not,
 * because this is a comparison of characters and a part that arrived UTF-16
 * and left UTF-8 is the same part.
 *
 * ## What is normalised
 *
 * - Attributes are sorted: namespace declarations first, then by namespace URI
 *   and local name. Attribute order is not significant in XML and two writers
 *   will not agree on it.
 * - Attribute values are the parser's, which are already normalised per XML 1.0
 *   3.3.3, re-escaped one way and quoted with `"`.
 * - `<a/>` and `<a></a>` both become `<a/>`.
 * - CDATA becomes ordinary escaped text. The two are the same character data;
 *   the lexical distinction is preserved by the writer and checked by `V027`,
 *   and it is not a difference between two *documents*.
 * - A byte order mark is dropped. It is a property of the encoding.
 *
 * ## What is not normalised, deliberately
 *
 * Whitespace, anywhere. It is tempting to collapse the whitespace between
 * elements, since no OOXML consumer reads it - and it would hide the one
 * failure this comparison exists to catch, which is a writer that reformats a
 * part on the way out. Text inside `a:t` is content and collapsing that would
 * be worse still.
 *
 * Comments and processing instructions are kept for the same reason: dropping
 * them would make losing them invisible.
 */

/** How many nested elements the walk will carry. Matches the parser's own ceiling. */
const MAX_DEPTH = 5000;

export interface CanonicalXmlOptions {
  /**
   * Rewrite an attribute value before it is written.
   *
   * The one caller is the round-trip comparator, which relabels relationship
   * ids: `r:embed="rId3"` and `r:embed="rId7"` are the same reference if the
   * two `.rels` agree about where rId3 and rId7 point. Passing the mapping in
   * rather than building it here keeps this module ignorant of relationships,
   * which are a package concept and not an XML one.
   *
   * `namespaceUri` is the attribute's namespace - the empty string for an
   * unprefixed attribute, which is in no namespace rather than in the element's
   * default one.
   */
  readonly rewriteAttributeValue?: (
    value: string,
    namespaceUri: string,
    localName: string,
  ) => string;
}

/** The namespace of `xmlns:*` declarations, which are attributes in the eyes of the tokenizer. */
const XMLNS = NS.xmlns;

interface Frame {
  readonly element: XElement;
  /** Prefix to URI, nearest binding winning. Shared with the parent when nothing is declared. */
  readonly scope: ReadonlyMap<string, string>;
  /** Index of the next child to visit. */
  at: number;
}

/** The implicit bindings every document has, per *Namespaces in XML* 3. */
const IMPLICIT: ReadonlyMap<string, string> = new Map([
  ['xml', NS.xml],
  ['xmlns', XMLNS],
]);

function scopeFor(
  element: XElement,
  parent: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> {
  if (!element.hasNamespaceDeclarations) return parent;
  const scope = new Map(parent);
  for (const attr of element.attributes) {
    if (attr.qname === 'xmlns') scope.set('', attr.value);
    else if (attr.prefix === 'xmlns') scope.set(attr.local, attr.value);
  }
  return scope;
}

/**
 * The sort key of one attribute, and the fact that a declaration is not sorted
 * with the rest.
 *
 * C14N puts namespace declarations first because they establish the meaning of
 * everything after them, and that reads better in a diff too: the first line of
 * a slide part is its namespace list rather than its `type` attribute.
 */
function sortKey(attr: XAttribute, scope: ReadonlyMap<string, string>): string {
  if (attr.qname === 'xmlns') return '0';
  if (attr.prefix === 'xmlns') return '1' + attr.local;
  // An unprefixed attribute is in no namespace - it does not pick up the
  // element's default. Its empty URI sorts first inside this group, which is
  // what C14N does with the same set.
  if (attr.prefix === '') return '2 ' + attr.local;
  const uri = scope.get(attr.prefix);
  // A prefix nothing binds has no URI to sort by, so it sorts by prefix in a
  // group of its own after the rest. Unreachable through `parseXml`, which
  // refuses the document first; here so that the order stays total either way.
  if (uri === undefined) return '3' + attr.prefix + ' ' + attr.local;
  return '2' + uri + ' ' + attr.local;
}

function namespaceOfAttribute(attr: XAttribute, scope: ReadonlyMap<string, string>): string {
  if (attr.qname === 'xmlns' || attr.prefix === 'xmlns') return XMLNS;
  if (attr.prefix === '') return '';
  return scope.get(attr.prefix) ?? '';
}

function startTag(
  element: XElement,
  scope: ReadonlyMap<string, string>,
  options: CanonicalXmlOptions,
): string {
  let out = '<' + element.qname;
  if (element.attributes.length > 0) {
    const ordered = element.attributes
      .map((attr) => ({ attr, key: sortKey(attr, scope) }))
      // The qname breaks a tie, so the order is total. Two attributes with the
      // same namespace and local name are not well-formed and the parser has
      // already refused them; this is here so that the sort is deterministic
      // rather than merely usually deterministic.
      .sort((a, b) =>
        a.key === b.key ? a.attr.qname.localeCompare(b.attr.qname) : a.key < b.key ? -1 : 1,
      );
    for (const { attr } of ordered) {
      const value =
        options.rewriteAttributeValue === undefined
          ? attr.value
          : options.rewriteAttributeValue(
              attr.value,
              namespaceOfAttribute(attr, scope),
              attr.local,
            );
      out += ' ' + attr.qname + '="' + escapeAttributeValue(value, '"') + '"';
    }
  }
  return out;
}

function canonicalDeclaration(document: XDocument): string {
  const declaration = document.declaration;
  if (declaration === undefined) return '';
  const standalone =
    declaration.standalone === undefined ? '' : ' standalone="' + declaration.standalone + '"';
  return '<?xml version="' + declaration.version + '"' + standalone + '?>';
}

/** A node outside any element - the prolog and the epilog, where only some kinds are legal. */
function canonicalLeaf(node: XNode, document: XDocument): string {
  switch (node.type) {
    case 'text':
      return escapeText(node.value);
    case 'cdata':
      return escapeText(node.value);
    case 'comment':
      return '<!--' + node.value + '-->';
    case 'processingInstruction': {
      // The tokenizer keeps the whitespace that separates the target from the
      // instruction, because the serializer has to put it back. Here it is one
      // of the things two producers will not agree on, so it becomes a single
      // space - which is what C14N does with the same construct.
      const data = node.data.replace(/^\s+/, '');
      return data === '' ? '<?' + node.target + '?>' : '<?' + node.target + ' ' + data + '?>';
    }
    case 'declaration':
      return canonicalDeclaration(document);
    default:
      return '';
  }
}

/**
 * The canonical form of `document`.
 *
 * Iterative, like everything else that walks a tree in this package: a part
 * nested a thousand deep is well-formed XML that PowerPoint opens, and a
 * recursive canonicaliser would be the one place in the pipeline that fell over
 * on it.
 */
export function canonicalXml(document: XDocument, options: CanonicalXmlOptions = {}): string {
  const out: string[] = [];
  const stack: Frame[] = [];

  const openElement = (element: XElement, parentScope: ReadonlyMap<string, string>): void => {
    const scope = scopeFor(element, parentScope);
    if (element.children.length === 0) {
      out.push(startTag(element, scope, options) + '/>');
      return;
    }
    out.push(startTag(element, scope, options) + '>');
    stack.push({ element, scope, at: 0 });
  };

  for (const child of document.children) {
    if (child.type !== 'element') {
      out.push(canonicalLeaf(child, document));
      continue;
    }
    openElement(child, IMPLICIT);

    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!;
      if (frame.at >= frame.element.children.length) {
        out.push('</' + frame.element.qname + '>');
        stack.pop();
        continue;
      }
      const node = frame.element.children[frame.at]!;
      frame.at++;
      if (node.type === 'element') {
        if (stack.length >= MAX_DEPTH) {
          // Unreachable through `parseXml`, whose own depth limit is lower.
          // Here so that a tree assembled by hand cannot turn this into a hang,
          // and typed rather than a bare `RangeError` because every failure in
          // this package carries a code a caller can branch on.
          throw new XmlError(
            'ERR_LIMIT_EXCEEDED',
            'canonicalXml: nesting deeper than ' + String(MAX_DEPTH),
            { offset: node.start },
          );
        }
        openElement(node, frame.scope);
      } else {
        out.push(canonicalLeaf(node, document));
      }
    }
  }

  return out.join('');
}

/**
 * The first index at which two canonical forms differ, or -1.
 *
 * Reported rather than the whole string, because the useful part of "these two
 * slides are not the same" is the forty characters around the disagreement and
 * not the two hundred kilobytes on either side of it.
 */
export function firstDifference(a: string, b: string): number {
  const shared = Math.min(a.length, b.length);
  for (let i = 0; i < shared; i++) if (a.charCodeAt(i) !== b.charCodeAt(i)) return i;
  return a.length === b.length ? -1 : shared;
}

/** A window of `text` around `at`, for a message a person has to read. */
export function excerpt(text: string, at: number, radius = 60): string {
  const from = Math.max(0, at - radius);
  const to = Math.min(text.length, at + radius);
  return (from > 0 ? '...' : '') + text.slice(from, to) + (to < text.length ? '...' : '');
}

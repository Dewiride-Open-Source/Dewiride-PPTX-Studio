import {
  attribute,
  type XAttribute,
  type XDocument,
  type XElement,
  type XNode,
} from '@pptx-studio/xml';

/**
 * Where a finding is, in a form a person can act on.
 *
 * ## Why an XPath and not a byte offset
 *
 * Both, actually - the offset is on the finding too, because it is what
 * `cli bisect` needs in sub-phase 1.5 and what a diff viewer can highlight. But
 * the offset is useless to a human: it names a position in a file that is
 * usually one line long and often two megabytes wide, and it stops being true
 * the moment anything above it changes.
 *
 * An XPath survives editing, survives reformatting, and can be pasted into
 * every XML tool there is. More to the point, it is the only form in which
 * a report about eleven parts is *readable* - `ppt/slides/slide3.xml` at
 * `/p:sld/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvPr/@id` says what is wrong in
 * one line, and byte 41 207 does not.
 *
 * ## Prefixes are the document's own, always
 *
 * The path is built from `qname` as written in the file - not from a canonical
 * prefix table, and not from the namespace URI. If a part spells the
 * PresentationML namespace `pp:` instead of `p:`, the path says `pp:`, because
 * the point of the path is that a reader can find the element in the file in
 * front of them. A canonicalised path would be *more correct* as an XPath
 * expression and less useful as a location, and this is a diagnostic rather
 * than a query.
 *
 * The same choice is made everywhere in this project for the same reason, and
 * it is not stylistic: `mc:Ignorable` and `mc:Choice/@Requires` hold prefixes
 * rather than URIs, so a prefix is load-bearing data in OOXML and inventing one
 * is never a neutral act.
 *
 * ## The positional predicate
 *
 * `[n]` is emitted only where it disambiguates - that is, when the element has
 * a sibling of the same `qname`. `/p:sld/p:cSld/p:spTree/p:sp[2]` is exact;
 * `/p:sld[1]/p:cSld[1]/p:spTree[1]` is noise. Positions are 1-based, counting
 * only elements of that name, which is what XPath means by `position()` inside
 * a name test and is what every XML tool will agree with.
 */

/** A location inside a part, or the part itself. */
export interface Location {
  /**
   * The part name, with its leading slash: `/ppt/slides/slide1.xml`.
   *
   * `/` means the finding is about the package rather than about any one part -
   * the archive shape, the content-type map, the relationship graph as a whole.
   */
  readonly part: string;
  /** The path to the element or attribute, or `null` for a whole-part finding. */
  readonly xpath: string | null;
  /** Byte offset of the node in the part's source, or `null`. For `cli bisect`. */
  readonly offset: number | null;
}

/** A location that names a part and nothing inside it. */
export function partLocation(part: string): Location {
  return { part, xpath: null, offset: null };
}

/** The location of the package itself: the archive, the content types, the graph. */
export const PACKAGE_LOCATION: Location = { part: '/', xpath: null, offset: null };

/** 1-based position among same-named element siblings, or `null` if it is alone. */
function positionAmongSiblings(element: XElement): number | null {
  const parent = element.parent;
  if (parent === undefined) return null;
  let position = 0;
  let seen = 0;
  for (const child of parent.children) {
    if (child.type !== 'element' || child.qname !== element.qname) continue;
    seen++;
    if (child === element) position = seen;
  }
  return seen > 1 ? position : null;
}

/**
 * The XPath of an element, from the document root.
 *
 * Never throws and never returns an empty string: a detached element - one
 * whose `parent` chain does not reach a root - still gets its own name, so a
 * finding built from a node the caller synthesised is degraded rather than
 * lost.
 */
export function xpathOf(element: XElement): string {
  const steps: string[] = [];
  for (let node: XElement | undefined = element; node !== undefined; node = node.parent) {
    const position = positionAmongSiblings(node);
    steps.push(position === null ? node.qname : node.qname + '[' + String(position) + ']');
  }
  steps.reverse();
  return '/' + steps.join('/');
}

/** The XPath of an attribute: its owner's path, then `/@qname`. */
export function xpathOfAttribute(owner: XElement, attr: XAttribute | string): string {
  const qname = typeof attr === 'string' ? attr : attr.qname;
  return xpathOf(owner) + '/@' + qname;
}

/** A location for an element inside a named part. */
export function elementLocation(part: string, element: XElement): Location {
  return { part, xpath: xpathOf(element), offset: element.start };
}

/**
 * A location for one attribute of an element.
 *
 * The offset is the attribute's own start when the element carries it, and the
 * element's start when it does not - which is the case that matters, because a
 * *missing* required attribute is a finding and it has to point somewhere.
 */
export function attributeLocation(part: string, element: XElement, qname: string): Location {
  const attr = attribute(element, qname);
  return {
    part,
    xpath: xpathOfAttribute(element, qname),
    offset: attr === undefined ? element.start : attr.start,
  };
}

/**
 * `line:column` for an offset, 1-based, or `null`.
 *
 * Not part of `Location`, because computing it means scanning the source from
 * the beginning and a report with three hundred findings would scan it three
 * hundred times. Callers that render a report for a human ask for it once, at
 * the point of rendering, for the findings they are about to show.
 */
export function lineColumn(source: string, offset: number): { line: number; column: number } {
  let line = 1;
  let lineStart = 0;
  const limit = Math.min(offset, source.length);
  for (let i = 0; i < limit; i++) {
    if (source.charCodeAt(i) === 0x0a) {
      line++;
      lineStart = i + 1;
    }
  }
  return { line, column: limit - lineStart + 1 };
}

/** Document order, for sorting findings within a part. */
export function inDocumentOrder(a: XNode, b: XNode): number {
  return a.start - b.start;
}

/** The root element's qname, for the handful of rules that dispatch on it. */
export function rootName(document: XDocument): string {
  return document.root.qname;
}

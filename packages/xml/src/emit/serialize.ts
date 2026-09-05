/**
 * Turning a tree back into the bytes it came from.
 *
 * The contract, and the reason this package exists:
 *
 * > **A node whose `dirty` flag is false re-emits as `source.slice(start,
 * > end)`.** Byte for byte, whatever it contains, whether or not we understand
 * > it.
 *
 * So a document nobody edited serializes to its own source, an edited one
 * differs only where it was edited, and a `p:timing` tree or an
 * `mc:AlternateContent` branch we have never looked at survives a save because
 * nothing ever asked us to reproduce it - only to copy it.
 *
 * ## Why this is more than twenty lines
 *
 * Taken literally, sub-phase 0.5's gate is passable by `return document.source`.
 * Nothing is ever dirty in a round trip, so the slice path is the only one it
 * exercises. The gate proves the *tokenizer's* spans tile the source; it does
 * not prove the serializer can write anything at all.
 *
 * What the gate cannot see is the rebuild path, which is the half sub-phase 0.6
 * runs on. So the work is in the rebuild, and it is held to a stronger property
 * than the plan asks for:
 *
 * > **Rebuilding a node from its value and reparsing it yields the same value.**
 *
 * That is a right-inverse requirement, and it is where the two-tier
 * normalization trap bites in the opposite direction. Four characters are lost
 * outright by a serializer that writes values out as it finds them - see
 * `escapeText` and `escapeAttributeValue` in `references.ts` for the table.
 *
 * ## The stack is explicit, for the same reason the parser's is
 *
 * PowerPoint opens a part nested 5000 elements deep. A recursive serializer
 * answers that with a `RangeError`, which is the one thing this package has
 * promised never to throw. So the walk carries its own stack, and the deferred
 * end tags ride on it as plain strings.
 */

import { isNameChar, isNameStartChar } from '../parse/chars.js';
import { XmlError } from '../errors.js';
import { escapeAttributeValue, escapeText } from '../mce/references.js';
import { encodeXmlSource } from '../parse/source.js';
import type { XAttribute } from '../parse/tokenizer.js';
import type { XDocument, XElement, XNode } from '../parse/xnode.js';

function fail(message: string, offset: number, name?: string): never {
  throw new XmlError('ERR_MALFORMED_XML', message, { offset, name });
}

/**
 * A name we are about to write must still be a name.
 *
 * Only reached on the rebuild path, so it costs a clean document nothing. It is
 * here because sub-phase 0.6 hands this function strings, and the failure it
 * prevents - a part that leaves us looking well-formed and is refused by
 * PowerPoint with no diagnostic - is the most expensive one in the project to
 * debug.
 */
function assertName(qname: string, at: number): void {
  const bad = (): never =>
    fail('"' + qname + '" is not a name and cannot be written as one', at, qname);
  if (qname.length === 0 || !isNameStartChar(qname.charCodeAt(0))) bad();
  for (let i = 1; i < qname.length; i++) {
    if (!isNameChar(qname.charCodeAt(i))) bad();
  }
  const colon = qname.indexOf(':');
  if (colon === 0 || colon === qname.length - 1 || qname.indexOf(':', colon + 1) > 0) {
    fail('"' + qname + '" is not a usable qualified name', at, qname);
  }
}

/** True if the span genuinely indexes `source`. */
function indexes(source: string, start: number, end: number): boolean {
  return start >= 0 && end >= start && end <= source.length;
}

const NON_SPACE = /[^ \t\r\n]/;

function sliceOf(source: string, node: XNode): string {
  if (!indexes(source, node.start, node.end)) {
    fail(
      'a clean node claims the span [' +
        node.start +
        ', ' +
        node.end +
        ') of a ' +
        source.length +
        '-character source; a synthesized node must be marked dirty',
      node.start,
    );
  }
  return source.slice(node.start, node.end);
}

/**
 * The whitespace an attribute was written with, when there is any to recover.
 *
 * `attribute.start` is the start of its *leading whitespace*, not of its name,
 * so this is a slice rather than a guess. It matters more than it looks: 70 822
 * of the 98 777 self-closing tags in the corpus come from a producer that puts a
 * space before the slash, and an edit that rewrites one attribute of such a tag
 * should leave the shape of the rest of it alone.
 */
function leadingSpace(source: string, attribute: XAttribute): string {
  const { start, nameStart } = attribute;
  if (nameStart > start && indexes(source, start, nameStart)) {
    const run = source.slice(start, nameStart);
    if (!NON_SPACE.test(run)) return run;
  }
  return ' ';
}

/** The whitespace between the last attribute and `>` or `/>`, likewise. */
function trailingSpace(source: string, element: XElement): string {
  const to = element.selfClosing ? element.end - 2 : element.openTagEnd - 1;
  if (to > element.trailingSpaceStart && indexes(source, element.trailingSpaceStart, to)) {
    const run = source.slice(element.trailingSpaceStart, to);
    if (!NON_SPACE.test(run)) return run;
  }
  return '';
}

function attributeText(source: string, attribute: XAttribute): string {
  if (!attribute.dirty && indexes(source, attribute.start, attribute.end)) {
    return source.slice(attribute.start, attribute.end);
  }
  assertName(attribute.qname, attribute.nameStart);
  return (
    leadingSpace(source, attribute) +
    attribute.qname +
    '=' +
    attribute.quote +
    escapeAttributeValue(attribute.value, attribute.quote, attribute.valueStart) +
    attribute.quote
  );
}

/**
 * Whether to write `<x/>` or `<x>`.
 *
 * Not simply `element.selfClosing`: sub-phase 0.6 adds children to elements that
 * were written empty, and `<a:ext cx="0" cy="0"/>` with a child is not a tag at
 * all. Deriving it here means no edit operation has to remember to flip the
 * flag. The reverse - an element written `<x></x>` whose children were all
 * removed - keeps its long form, because that is what it was.
 */
function isEmptyForm(element: XElement): boolean {
  return element.selfClosing && element.children.length === 0;
}

function startTag(source: string, element: XElement): string {
  assertName(element.qname, element.start);
  let out = '<' + element.qname;
  for (const attribute of element.attributes) out += attributeText(source, attribute);
  return out + trailingSpace(source, element) + (isEmptyForm(element) ? '/>' : '>');
}

/**
 * A construct with no escaping mechanism cannot carry `forbidden` at all.
 *
 * Comments, CDATA sections and processing instructions each end at a fixed
 * character sequence and have no way to quote it. There is nothing clever to do
 * about a value containing one; refusing is the honest answer, and the caller
 * that produced it has a bug.
 */
function assertAbsent(value: string, forbidden: string, node: XNode, what: string): void {
  const at = value.indexOf(forbidden);
  if (at >= 0) {
    fail(
      'a ' + what + ' cannot contain "' + forbidden + '"; there is no way to escape it',
      node.start + at,
    );
  }
}

/**
 * A carriage return in a comment or a CDATA section is unrepresentable.
 *
 * XML 1.0 §2.11 normalizes line endings everywhere in the document, and neither
 * construct admits a character reference to undo it, so a literal `\r` written
 * into one reparses as `\n`. The parser therefore never produces such a value -
 * only a caller that synthesized one can, and it is better told than silently
 * altered. Text nodes have `&#xD;` and are not affected.
 */
function assertNoCarriageReturn(value: string, node: XNode, what: string): void {
  const at = value.indexOf('\r');
  if (at >= 0) {
    fail(
      'a ' +
        what +
        ' cannot carry a carriage return: XML normalizes line endings and offers no escape here, so it would come back as a line feed',
      node.start + at,
    );
  }
}

function rebuiltLeaf(node: Exclude<XNode, XElement>): string {
  switch (node.type) {
    case 'text':
      return escapeText(node.value, node.start);
    case 'cdata':
      assertAbsent(node.value, ']]>', node, 'CDATA section');
      assertNoCarriageReturn(node.value, node, 'CDATA section');
      return '<![CDATA[' + node.value + ']]>';
    case 'comment':
      // §2.5. "--" is forbidden outright, which also forbids a comment ending in
      // "-", because that would make the terminator "--->".
      assertAbsent(node.value, '--', node, 'comment');
      if (node.value.endsWith('-')) fail('a comment may not end with "-"', node.end);
      assertNoCarriageReturn(node.value, node, 'comment');
      return '<!--' + node.value + '-->';
    case 'processingInstruction': {
      assertName(node.target, node.start);
      if (node.target.toLowerCase() === 'xml') {
        fail('"' + node.target + '" is a reserved processing-instruction target', node.start);
      }
      assertAbsent(node.data, '?>', node, 'processing instruction');
      if (node.data.length > 0 && NON_SPACE.test(node.data.charAt(0))) {
        fail('a processing-instruction target must be followed by whitespace', node.start);
      }
      return '<?' + node.target + node.data + '?>';
    }
    case 'declaration': {
      // Rebuilt with double quotes: `XDeclaration` records the three values and
      // not the quoting, because nothing edits a declaration. A clean one is a
      // slice and keeps whatever it had.
      let out = '<?xml version="' + node.version + '"';
      if (node.encoding !== undefined) out += ' encoding="' + node.encoding + '"';
      if (node.standalone !== undefined) out += ' standalone="' + node.standalone + '"';
      return out + '?>';
    }
  }
}

/** Append the serialization of `nodes`, and everything beneath them, to `out`. */
function emit(out: string[], source: string, nodes: readonly XNode[]): void {
  const stack: (XNode | string)[] = [];
  for (let i = nodes.length - 1; i >= 0; i--) stack.push(nodes[i]!);

  while (stack.length > 0) {
    const work = stack.pop()!;
    if (typeof work === 'string') {
      out.push(work);
      continue;
    }
    if (!work.dirty) {
      out.push(sliceOf(source, work));
      continue;
    }
    if (work.type !== 'element') {
      out.push(rebuiltLeaf(work));
      continue;
    }
    out.push(startTag(source, work));
    if (isEmptyForm(work)) continue;
    // Pushed first so it is popped last, after every child.
    stack.push('</' + work.qname + '>');
    for (let i = work.children.length - 1; i >= 0; i--) stack.push(work.children[i]!);
  }
}

/**
 * Serialize a document.
 *
 * For a document nobody edited this returns `document.source`, reached one slice
 * per top-level node - the byte-identical round trip, and it falls out of the
 * span invariant rather than being a special case in the code.
 */
export function serializeXmlString(document: XDocument): string {
  const out: string[] = [];
  // The byte order mark is not a node - it is markup-free and may precede the
  // declaration - so it lives on the document and is re-emitted here.
  if (document.bom) out.push('\ufeff');
  emit(out, document.source, document.children);
  return out.join('');
}

/** Serialize a document to the bytes that go back into the package. */
export function serializeXml(document: XDocument): Uint8Array {
  return encodeXmlSource(serializeXmlString(document));
}

/**
 * Serialize one node and its subtree.
 *
 * The counterpart to `sourceOf`, which is only correct while the node is clean.
 */
export function serializeNode(document: XDocument, node: XNode): string {
  const out: string[] = [];
  emit(out, document.source, [node]);
  return out.join('');
}

import { OpcError } from './errors.js';

/**
 * A reader for flat, attribute-only XML.
 *
 * This is **not** the XML layer. `@pptx-studio/xml` (sub-phase 0.4) builds the
 * real thing: a byte-offset tokenizer over `Uint8Array` producing an `XNode`
 * tree that re-emits clean subtrees by slicing the original buffer. That
 * machinery exists so a slide part can be edited in one place without
 * disturbing the animation tree next to it.
 *
 * The two package-level parts need none of it, and it is worth saying why
 * rather than treating this file as a stopgap:
 *
 * - `[Content_Types].xml` and every `.rels` part are **flat**: a root element
 *   and one level of empty children. No text content, no mixed content, no
 *   `mc:AlternateContent`, no `extLst`.
 * - Their grammar is fixed by ECMA-376 **Part 2**, not by the Part 1 schemas
 *   that 0.6 generates element-order tables from. It cannot drift.
 * - They are never edited node by node. When the content-type map or a
 *   relationship collection changes we regenerate the whole part, so there is
 *   nothing for byte-level preservation to preserve.
 *
 * So they get a reader sized to the job. What it does have is the same posture
 * as everything else here: it reads untrusted bytes, so it rejects rather than
 * interprets. **A DOCTYPE is refused outright** - that single rule removes XXE
 * and billion-laughs from this package's attack surface, because there is no
 * DTD subset to expand. An undeclared entity reference is likewise a hard
 * error, not a passthrough.
 */

/** One start tag, with its attributes as written. */
export interface FlatElement {
  /** Local name with any prefix stripped: `Override`, not `ct:Override`. */
  readonly name: string;
  /** The name exactly as written, prefix included. */
  readonly qname: string;
  /** 0 for the root element. */
  readonly depth: number;
  /** Attributes keyed by their written name, values entity-decoded. */
  readonly attrs: ReadonlyMap<string, string>;
}

export interface FlatXmlLimits {
  /** Ceiling on elements in one part. */
  readonly maxElements: number;
  /** Ceiling on attributes on one element. */
  readonly maxAttributes: number;
  /** Ceiling on nesting. These parts are two deep; anything more is not one of them. */
  readonly maxDepth: number;
}

export const DEFAULT_FLAT_XML_LIMITS: FlatXmlLimits = {
  maxElements: 100_000,
  maxAttributes: 64,
  maxDepth: 8,
};

const NAME_END = new Set([' ', '\t', '\r', '\n', '/', '>', '=']);
const WHITESPACE = new Set([' ', '\t', '\r', '\n']);

function fail(message: string, context: string, offset?: number): never {
  throw new OpcError('ERR_MALFORMED_XML', context + ': ' + message, {
    entry: context,
    ...(offset === undefined ? {} : { offset }),
  });
}

/** Strip a prefix: `ct:Override` -> `Override`. */
export function localName(qname: string): string {
  const colon = qname.indexOf(':');
  return colon < 0 ? qname : qname.slice(colon + 1);
}

/**
 * Decode `bytes` as XML text.
 *
 * OPC permits UTF-8 and UTF-16, with or without a byte-order mark, and the mark
 * is not hypothetical: Office writes these parts as UTF-8 with no BOM, while
 * the other producer in our corpus prefixes every one of them with `ef bb bf`.
 * We missed that at first because the census decoded through `TextDecoder`,
 * which had already eaten it - so this handles the mark on the bytes, before
 * anything can quietly swallow it. An encoding declaration naming anything
 * else shows up as a decode failure rather than as mojibake, which is the
 * outcome we want.
 */
function decodeXml(bytes: Uint8Array, context: string): string {
  let label = 'utf-8';
  let start = 0;
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    label = 'utf-16le';
    start = 2;
  } else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    label = 'utf-16be';
    start = 2;
  } else if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    start = 3;
  }
  try {
    return new TextDecoder(label, { fatal: true }).decode(bytes.subarray(start));
  } catch (cause) {
    throw new OpcError('ERR_MALFORMED_XML', context + ': not valid ' + label + ' text', {
      entry: context,
      cause,
    });
  }
}

const NAMED_ENTITIES = new Map([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
]);

/**
 * Resolve the five predefined entities and numeric character references.
 *
 * Anything else throws. In a document with no DTD every other entity reference
 * is undeclared and therefore a well-formedness error, so refusing them is
 * both correct and the reason no entity-expansion attack reaches this code.
 * Hyperlink targets carry `&amp;` constantly, so this path is not exotic.
 */
function decodeEntities(raw: string, context: string): string {
  if (!raw.includes('&')) return raw;
  let out = '';
  let i = 0;
  for (;;) {
    const amp = raw.indexOf('&', i);
    if (amp < 0) return out + raw.slice(i);
    out += raw.slice(i, amp);
    const end = raw.indexOf(';', amp);
    if (end < 0) fail('unterminated entity reference in an attribute value', context);
    const body = raw.slice(amp + 1, end);
    const named = NAMED_ENTITIES.get(body);
    if (named !== undefined) {
      out += named;
    } else if (/^#\d+$/.test(body) || /^#x[0-9A-Fa-f]+$/.test(body)) {
      const code = body[1] === 'x' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      if (
        !Number.isFinite(code) ||
        code < 0 ||
        code > 0x10ffff ||
        (code >= 0xd800 && code <= 0xdfff)
      ) {
        fail('character reference &' + body + '; is not a valid code point', context);
      }
      out += String.fromCodePoint(code);
    } else {
      fail(
        'undeclared entity reference &' +
          body +
          ';. This document has no DTD, so every entity beyond the five predefined ones is ' +
          'undeclared - and a DTD is exactly what this reader refuses to have.',
        context,
      );
    }
    i = end + 1;
  }
}

/**
 * Read every start tag in document order.
 *
 * The caller takes element 0 as the root and filters the rest by name. Text
 * between tags is ignored: neither of these two grammars has any.
 */
export function readFlatXml(
  bytes: Uint8Array,
  context: string,
  limits: FlatXmlLimits = DEFAULT_FLAT_XML_LIMITS,
): FlatElement[] {
  const s = decodeXml(bytes, context);
  const out: FlatElement[] = [];
  const open: string[] = [];
  let rootsSeen = 0;
  let i = 0;

  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt < 0) break;
    i = lt;

    if (s.startsWith('<!--', i)) {
      const end = s.indexOf('-->', i + 4);
      if (end < 0) fail('unterminated comment', context, i);
      i = end + 3;
      continue;
    }
    if (s.startsWith('<![CDATA[', i)) {
      const end = s.indexOf(']]>', i + 9);
      if (end < 0) fail('unterminated CDATA section', context, i);
      i = end + 3;
      continue;
    }
    if (s.startsWith('<!', i)) {
      // `<!DOCTYPE`, `<!ENTITY`, `<!ELEMENT` - the whole markup-declaration
      // family, refused as one. See the file comment.
      throw new OpcError(
        'ERR_DOCTYPE_FORBIDDEN',
        context +
          ': this part contains a document type declaration. Package parts have no DTD, and a ' +
          'DTD is where entity expansion and external entity resolution live, so the construct ' +
          'is refused rather than implemented safely.',
        { entry: context, offset: i },
      );
    }
    if (s.startsWith('<?', i)) {
      const end = s.indexOf('?>', i + 2);
      if (end < 0) fail('unterminated processing instruction', context, i);
      i = end + 2;
      continue;
    }
    if (s.startsWith('</', i)) {
      const end = s.indexOf('>', i + 2);
      if (end < 0) fail('unterminated end tag', context, i);
      const name = s.slice(i + 2, end).trim();
      const expected = open.pop();
      if (expected === undefined) fail('end tag </' + name + '> with nothing open', context, i);
      if (expected !== name) {
        fail('end tag </' + name + '> closes <' + expected + '>', context, i);
      }
      i = end + 1;
      continue;
    }

    // --- a start tag ---
    if (out.length >= limits.maxElements) {
      fail('more than ' + String(limits.maxElements) + ' elements', context, i);
    }
    let p = i + 1;
    const nameStart = p;
    while (p < s.length && !NAME_END.has(s[p]!)) p++;
    const qname = s.slice(nameStart, p);
    if (qname === '') fail('empty element name', context, i);

    const attrs = new Map<string, string>();
    let selfClosing = false;
    for (;;) {
      while (p < s.length && WHITESPACE.has(s[p]!)) p++;
      if (p >= s.length) fail('unterminated start tag <' + qname + '>', context, i);
      if (s[p] === '>') {
        p++;
        break;
      }
      if (s.startsWith('/>', p)) {
        selfClosing = true;
        p += 2;
        break;
      }
      const attrStart = p;
      while (p < s.length && !NAME_END.has(s[p]!)) p++;
      const attrName = s.slice(attrStart, p);
      if (attrName === '') fail('malformed attribute in <' + qname + '>', context, p);
      while (p < s.length && WHITESPACE.has(s[p]!)) p++;
      if (s[p] !== '=') fail('attribute ' + attrName + ' has no value', context, p);
      p++;
      while (p < s.length && WHITESPACE.has(s[p]!)) p++;
      const quote = s[p];
      if (quote !== '"' && quote !== "'") {
        fail('attribute ' + attrName + ' value is not quoted', context, p);
      }
      const valueStart = p + 1;
      const valueEnd = s.indexOf(quote, valueStart);
      if (valueEnd < 0) fail('unterminated value for attribute ' + attrName, context, p);
      if (attrs.has(attrName)) {
        fail('attribute ' + attrName + ' appears twice on <' + qname + '>', context, p);
      }
      if (attrs.size >= limits.maxAttributes) {
        fail(
          'more than ' + String(limits.maxAttributes) + ' attributes on <' + qname + '>',
          context,
          p,
        );
      }
      attrs.set(attrName, decodeEntities(s.slice(valueStart, valueEnd), context));
      p = valueEnd + 1;
    }

    const depth = open.length;
    if (depth === 0) {
      rootsSeen++;
      if (rootsSeen > 1) fail('a second root element <' + qname + '>', context, i);
    }
    if (depth >= limits.maxDepth) {
      fail('nested deeper than ' + String(limits.maxDepth) + ' elements', context, i);
    }
    out.push({ name: localName(qname), qname, depth, attrs });
    if (!selfClosing) open.push(qname);
    i = p;
  }

  if (open.length > 0) fail('<' + open[open.length - 1]! + '> is never closed', context);
  if (out.length === 0) fail('no elements at all', context);
  return out;
}

// --- writing ---------------------------------------------------------------

/**
 * Escape a string for use in an attribute value.
 *
 * `<` is escaped although a bare `<` is only forbidden in *content*; it is
 * cheap and it means an attribute value can never terminate the tag around it.
 * `>` is escaped for symmetry with what Office writes.
 */
export function escapeAttribute(value: string): string {
  let out = '';
  for (const ch of value) {
    switch (ch) {
      case '&':
        out += '&amp;';
        break;
      case '<':
        out += '&lt;';
        break;
      case '>':
        out += '&gt;';
        break;
      case '"':
        out += '&quot;';
        break;
      default:
        out += ch;
    }
  }
  return out;
}

/**
 * The XML declaration Office writes, verbatim - **including the CRLF**.
 *
 * Measured from PowerPoint 365 and PowerPoint 14 output rather than chosen:
 * uppercase `UTF-8`, `standalone="yes"`, no BOM, then `0d 0a`, then the whole
 * rest of the document on one line with no trailing newline. That line break is
 * the only whitespace in either package grammar and it is easy to miss, since
 * it is invisible in every tool that shows you the file as text.
 *
 * The other producer in our corpus writes something different at every one of
 * those decisions - a UTF-8 BOM, lowercase `utf-8`, no `standalone`, no line
 * break, and ` />` with a space. Both open. We match Office because when a
 * regenerated part is diffed against an Office-authored one, the useful diff is
 * the one that is empty.
 */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

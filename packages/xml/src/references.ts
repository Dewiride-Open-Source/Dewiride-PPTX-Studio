/**
 * Resolving entity and character references, and normalizing attribute values.
 *
 * With no DTD - and there is never a DTD, because we reject DOCTYPE outright -
 * the entity vocabulary is closed: the five predefined entities of XML 1.0
 * §4.6 and nothing else. PowerPoint agrees; `&nbsp;` inside an `<a:t>` is
 * refused with `0x80070570`.
 */

import { isXmlCodePoint } from './chars.js';
import { XmlError } from './errors.js';

/** XML 1.0 §4.6. The complete list, because there is no DTD to extend it. */
const PREDEFINED = new Map<string, string>([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
]);

/**
 * The longest thing that can legally sit between `&` and `;`.
 *
 * `#x` plus six hex digits is eight; the longest predefined name is four. The
 * cap exists so that a stray `&` in a megabyte of text costs a bounded scan
 * rather than a search to the end of the part - the difference between linear
 * and quadratic behaviour on hostile input.
 */
const MAX_REFERENCE_BODY = 12;

function invalidCodePoint(code: number, offset: number, raw: string): never {
  throw new XmlError(
    'ERR_INVALID_CHARACTER',
    'character reference "' +
      raw +
      '" names U+' +
      code.toString(16).toUpperCase().padStart(4, '0') +
      ', which is not an XML character' +
      (code >= 0xd800 && code <= 0xdfff ? ' (it is half of a surrogate pair)' : ''),
    { offset, name: raw },
  );
}

/**
 * No `;` was found before the run ended, or before the cap.
 *
 * The two are different failures and deserve different codes. "There is no `;`
 * here at all" is a bare ampersand, which is what an author who forgot to
 * escape one has written. "The body ran past the cap" is a real reference that
 * is too long to be one, and calling that a bare ampersand sends the reader
 * hunting for an `&` that is correctly written.
 */
function unterminated(offset: number, capped: boolean): never {
  throw new XmlError(
    capped ? 'ERR_LIMIT_EXCEEDED' : 'ERR_MALFORMED_XML',
    capped
      ? 'no reference may be longer than ' +
          MAX_REFERENCE_BODY +
          ' characters; the cap is what keeps a stray "&" a bounded scan rather than a search to the end of the part'
      : 'a bare "&" is not allowed in XML; write "&amp;" for a literal ampersand',
    { offset, limit: MAX_REFERENCE_BODY },
  );
}

/**
 * `CharRef ::= '&#' [0-9]+ ';' | '&#x' [0-9a-fA-F]+ ';'`.
 *
 * The `x` is a terminal in a case-sensitive grammar - only the *digits* are
 * case-insensitive - so `&#X41;` is not a character reference at all.
 */
function characterReference(body: string, raw: string, offset: number): string {
  const hex = body.charCodeAt(1) === 0x78;
  const digits = hex ? body.slice(2) : body.slice(1);
  if (digits.length === 0 || !(hex ? /^[0-9a-fA-F]+$/ : /^[0-9]+$/).test(digits)) {
    throw new XmlError('ERR_MALFORMED_XML', '"' + raw + '" is not a character reference', {
      offset,
      name: raw,
    });
  }
  const code = parseInt(digits, hex ? 16 : 10);
  if (!isXmlCodePoint(code)) invalidCodePoint(code, offset, raw);
  return String.fromCodePoint(code);
}

/**
 * Decode the reference starting at `source[start]`, which must be `&`.
 *
 * `end` bounds the run the reference must live inside - the end of a text node
 * or of an attribute value. Passing it is not optional: without it the scan
 * runs past the closing quote or the `<` that terminates the run, and a
 * reference that is merely unterminated gets diagnosed as an unknown entity
 * whose name contains the rest of the document. `<r a="&amp" b="x;"/>` reported
 * `no entity named "amp" b="x"` before this argument existed.
 *
 * Returns the replacement text and the index one past the `;`.
 */
export function decodeReference(
  source: string,
  start: number,
  end: number = source.length,
): { readonly text: string; readonly end: number } {
  const stop = Math.min(source.length, end);
  const limit = Math.min(stop, start + 1 + MAX_REFERENCE_BODY + 1);
  let semi = -1;
  for (let i = start + 1; i < limit; i++) {
    if (source.charCodeAt(i) === 0x3b) {
      semi = i;
      break;
    }
  }
  if (semi < 0) unterminated(start, limit < stop);

  const body = source.slice(start + 1, semi);
  const raw = source.slice(start, semi + 1);

  if (body.charCodeAt(0) === 0x23) {
    return { text: characterReference(body, raw, start), end: semi + 1 };
  }

  const replacement = PREDEFINED.get(body);
  if (replacement === undefined) {
    throw new XmlError(
      'ERR_UNKNOWN_ENTITY',
      'no entity named "' +
        body +
        '"; without a DTD the only entities are amp, lt, gt, quot and apos',
      { offset: start, name: body },
    );
  }
  return { text: replacement, end: semi + 1 };
}

/**
 * Expand every reference in a run of character data, and apply XML 1.0 §2.11
 * end-of-line normalization to the literal text around them.
 *
 * The two are interleaved rather than applied one after the other, and that is
 * load-bearing: `&#xD;` is a *reference* to a carriage return and survives
 * §2.11 untouched, whereas a literal `\r` does not. Running
 * `normalizeLineEndings` over the already-expanded string would silently turn
 * the first into a line feed and lose the distinction.
 */
export function decodeCharacterData(source: string, start: number, end: number): string {
  let out = '';
  let plain = start;
  let i = start;
  while (i < end) {
    const code = source.charCodeAt(i);
    if (code === 0x26) {
      out += source.slice(plain, i);
      const { text, end: next } = decodeReference(source, i, end);
      out += text;
      i = next;
      plain = i;
      continue;
    }
    if (code === 0x0d) {
      out += source.slice(plain, i) + '\n';
      i += source.charCodeAt(i + 1) === 0x0a ? 2 : 1;
      plain = i;
      continue;
    }
    i++;
  }
  return plain === start ? source.slice(start, end) : out + source.slice(plain, end);
}

/**
 * XML 1.0 §3.3.3 attribute-value normalization, for the CDATA case - which is
 * every attribute we will ever see, because declaring an attribute as anything
 * else requires a DTD.
 *
 * The rule that catches people: a **literal** tab, line feed or carriage
 * return in the value becomes a space, but a **character reference** to the
 * same character does not. So `name="a&#9;b"` and `name="a<TAB>b"` are two
 * different values, and a normalizer that expands references first and then
 * squashes whitespace collapses them into one.
 */
export function normalizeAttributeValue(source: string, start: number, end: number): string {
  let out = '';
  let plain = start;
  let i = start;
  while (i < end) {
    const code = source.charCodeAt(i);
    if (code === 0x26) {
      out += source.slice(plain, i);
      const { text, end: next } = decodeReference(source, i, end);
      out += text; // NOT whitespace-normalized: it came from a reference
      i = next;
      plain = i;
      continue;
    }
    if (code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d) {
      out += source.slice(plain, i) + ' ';
      // \r\n is one line break and so becomes one space, per §2.11 running first.
      i += code === 0x0d && source.charCodeAt(i + 1) === 0x0a ? 2 : 1;
      plain = i;
      continue;
    }
    i++;
  }
  return plain === start ? source.slice(start, end) : out + source.slice(plain, end);
}

/**
 * True if the run needs no transformation at all, so callers can hand back a
 * plain slice.
 *
 * Worth the extra pass: across the corpus there are 39 entity references in
 * 194 148 elements, so the overwhelmingly common case is a value that is
 * already its own decoding.
 */
export function isLiteralRun(
  source: string,
  start: number,
  end: number,
  attribute: boolean,
): boolean {
  for (let i = start; i < end; i++) {
    const code = source.charCodeAt(i);
    if (code === 0x26 || code === 0x0d) return false;
    if (attribute && (code === 0x09 || code === 0x0a)) return false;
  }
  return true;
}

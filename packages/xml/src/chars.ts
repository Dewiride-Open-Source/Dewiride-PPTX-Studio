/**
 * The XML 1.0 character productions, as predicates over UTF-16 code units.
 *
 * We implement the **Fifth Edition** `Name` production. The change it made is
 * not cosmetic: the Fourth Edition enumerated permitted name characters
 * against Unicode 2.0 tables, so a name containing a character added to
 * Unicode later was invalid; the Fifth Edition inverted this into a small
 * exclusion list, which is both simpler and more permissive. Implementing the
 * Fourth Edition would reject names that current tools legitimately emit.
 *
 * These operate on code *units*, not code points, and that is deliberate. Our
 * source is a string decoded from valid UTF-8, so every surrogate in it is
 * already half of a well-formed pair, and a pair always denotes a code point in
 * [U+10000, U+10FFFF] - which the `Char` production permits in full. Checking
 * units therefore needs no pair reassembly on the hot path. The one place that
 * reasoning does not hold is a numeric character reference, which can name a
 * lone surrogate directly, so {@link isXmlCodePoint} exists for that case.
 */

/** XML 1.0 §2.3 `S`: space, tab, carriage return, line feed. Nothing else. */
export function isXmlWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x0a || code === 0x09 || code === 0x0d;
}

/**
 * XML 1.0 Fifth Edition §2.3 `NameStartChar`.
 *
 * Note that `:` is included. The colon is a legal name character in XML
 * itself; it is *Namespaces in XML* that gives it meaning, and splitting a
 * qualified name is a separate concern from validating one.
 */
export function isNameStartChar(code: number): boolean {
  return (
    (code >= 0x61 && code <= 0x7a) || // a-z
    (code >= 0x41 && code <= 0x5a) || // A-Z
    code === 0x3a || // :
    code === 0x5f || // _
    (code >= 0xc0 && code <= 0xd6) ||
    (code >= 0xd8 && code <= 0xf6) ||
    (code >= 0xf8 && code <= 0x2ff) ||
    (code >= 0x370 && code <= 0x37d) ||
    (code >= 0x37f && code <= 0x1fff) ||
    (code >= 0x200c && code <= 0x200d) ||
    (code >= 0x2070 && code <= 0x218f) ||
    (code >= 0x2c00 && code <= 0x2fef) ||
    (code >= 0x3001 && code <= 0xd7ff) ||
    (code >= 0xf900 && code <= 0xfdcf) ||
    (code >= 0xfdf0 && code <= 0xfffd) ||
    // [#x10000-#xEFFFF] as a surrogate pair: any high surrogate below the
    // range that would encode a code point past U+EFFFF.
    (code >= 0xd800 && code <= 0xdb7f)
  );
}

/** XML 1.0 Fifth Edition §2.3 `NameChar` - `NameStartChar` plus these. */
export function isNameChar(code: number): boolean {
  return (
    isNameStartChar(code) ||
    (code >= 0x30 && code <= 0x39) || // 0-9
    code === 0x2d || // -
    code === 0x2e || // .
    code === 0xb7 ||
    (code >= 0x300 && code <= 0x36f) ||
    (code >= 0x203f && code <= 0x2040) ||
    // low surrogates, completing a NameStartChar pair
    (code >= 0xdc00 && code <= 0xdfff)
  );
}

/**
 * XML 1.0 §2.2 `Char`, as a predicate over a code unit of already-decoded
 * text.
 *
 * The excluded set is small and worth knowing by heart, because it is what
 * PowerPoint enforces: the C0 controls other than tab, line feed and carriage
 * return; and the two permanently-unassigned noncharacters U+FFFE and U+FFFF.
 * A literal U+000B in an `<a:t>` is refused by PowerPoint with `0x80070570`.
 */
export function isXmlChar(code: number): boolean {
  if (code < 0x20) return code === 0x09 || code === 0x0a || code === 0x0d;
  return code !== 0xfffe && code !== 0xffff;
}

/**
 * True if the code unit at `index` is half of a surrogate pair with nothing on
 * the other side of it.
 *
 * Cheap because it only does work inside the surrogate range, and it closes the
 * one hole in this package's central guarantee. Decoding valid UTF-8 can never
 * produce a lone surrogate, so `parseXml(bytes)` is safe by construction - but
 * `parseXmlString` takes a string a caller assembled, and a lone surrogate is
 * the single value for which decode and encode are not inverses:
 * `encodeXmlSource('\uD800')` emits `EF BF BD`, the replacement character.
 * Left unchecked, a document would parse, pass coverage, and re-serialize to
 * different bytes.
 */
export function isUnpairedSurrogate(source: string, index: number): boolean {
  const code = source.charCodeAt(index);
  if (code < 0xd800 || code > 0xdfff) return false;
  if (code <= 0xdbff) {
    const next = source.charCodeAt(index + 1);
    return !(next >= 0xdc00 && next <= 0xdfff);
  }
  const previous = source.charCodeAt(index - 1);
  return !(index > 0 && previous >= 0xd800 && previous <= 0xdbff);
}

/**
 * XML 1.0 §2.2 `Char`, as a predicate over a full code point.
 *
 * Only a numeric character reference can reach the cases this catches and
 * {@link isXmlChar} cannot: `&#xD800;` names a lone surrogate, which is not a
 * character at all, and letting one into a JavaScript string produces a value
 * that cannot be encoded back to UTF-8. `&#x110000;` is past the end of
 * Unicode.
 */
export function isXmlCodePoint(code: number): boolean {
  if (code < 0x20) return code === 0x09 || code === 0x0a || code === 0x0d;
  if (code <= 0xd7ff) return true;
  if (code < 0xe000) return false; // surrogate range
  if (code <= 0xfffd) return true;
  return code >= 0x10000 && code <= 0x10ffff;
}

/**
 * XML 1.0 §2.11 end-of-line handling: a processor must behave as though every
 * `\r\n` and every lone `\r` in the input were a single `\n`.
 *
 * This is one of the two places the specification mandates a transformation of
 * the input, and it is why a node's *value* and a node's *source* are two
 * different things in this package. Writing `line1\r\nline2` into an `<a:t>`
 * and asking PowerPoint for the text back yields `line1\nline2` - PowerPoint
 * performs this normalization, and so must we. But a tokenizer that stores only
 * the normalized form can never reproduce the original bytes, so we store the
 * span and normalize on demand.
 */
export function normalizeLineEndings(text: string): string {
  if (text.indexOf('\r') < 0) return text;
  return text.replace(/\r\n?/g, '\n');
}

/** True if the run contains only `S` characters. */
export function isAllWhitespace(source: string, start: number, end: number): boolean {
  for (let i = start; i < end; i++) {
    if (!isXmlWhitespace(source.charCodeAt(i))) return false;
  }
  return true;
}

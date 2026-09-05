import { describe, expect, it } from 'vitest';
import {
  isAllWhitespace,
  isNameChar,
  isNameStartChar,
  isXmlChar,
  isXmlCodePoint,
  isXmlWhitespace,
  normalizeLineEndings,
} from './chars.js';

const codeOf = (ch: string): number => ch.charCodeAt(0);

describe('the S production', () => {
  it('is four characters and not one more', () => {
    for (const ch of [' ', '\t', '\r', '\n']) expect(isXmlWhitespace(codeOf(ch))).toBe(true);
    // The ones people assume are whitespace and XML does not.
    for (const ch of ['\f', '\v', ' ', ' ', '　']) expect(isXmlWhitespace(codeOf(ch))).toBe(false);
  });
});

describe('the Name productions, Fifth Edition', () => {
  it('accepts the ASCII names OOXML actually uses', () => {
    for (const ch of 'abzABZ_') expect(isNameStartChar(codeOf(ch))).toBe(true);
    for (const ch of '0.-·') expect(isNameStartChar(codeOf(ch))).toBe(false);
    for (const ch of '0.-·') expect(isNameChar(codeOf(ch))).toBe(true);
  });

  it('treats ":" as an ordinary name character', () => {
    // XML itself has no opinion about the colon - it is *Namespaces in XML*
    // that gives it meaning. Splitting a qualified name is a separate concern
    // from deciding whether it is a name at all.
    expect(isNameStartChar(codeOf(':'))).toBe(true);
    expect(isNameChar(codeOf(':'))).toBe(true);
  });

  it('accepts characters the Fourth Edition would have rejected', () => {
    // The Fourth Edition enumerated permitted name characters against Unicode
    // 2.0, so anything added to Unicode later was invalid in a name. The Fifth
    // inverted that into a small exclusion list. U+0870 (Arabic Extended-B,
    // added in Unicode 14) is inside the Fifth Edition's [#x370-#x1FFF] range
    // and was not assigned when the Fourth was written.
    expect(isNameStartChar(0x0870)).toBe(true);
  });

  it('rejects the characters that separate one name from the next', () => {
    for (const ch of ' \t\r\n<>/=\'"&') expect(isNameChar(codeOf(ch))).toBe(false);
  });
});

describe('the Char production', () => {
  it('rejects the C0 controls other than tab, line feed and carriage return', () => {
    expect(isXmlChar(0x09)).toBe(true);
    expect(isXmlChar(0x0a)).toBe(true);
    expect(isXmlChar(0x0d)).toBe(true);
    expect(isXmlChar(0x00)).toBe(false);
    expect(isXmlChar(0x1f)).toBe(false);
    // Verified against PowerPoint: a literal U+000B inside an <a:t> makes it
    // refuse the whole package with 0x80070570. This is not us being strict.
    expect(isXmlChar(0x0b)).toBe(false);
  });

  it('rejects the two noncharacters at the end of the BMP', () => {
    expect(isXmlChar(0xfffd)).toBe(true);
    expect(isXmlChar(0xfffe)).toBe(false);
    expect(isXmlChar(0xffff)).toBe(false);
  });

  it('accepts a surrogate code unit, because a decoded string only holds pairs', () => {
    // Scanning code units is what makes the hot loop cheap. It is sound only
    // because the source was decoded from valid UTF-8, where every surrogate is
    // already half of a well-formed pair denoting a code point the Char
    // production permits in full.
    expect(isXmlChar(0xd83d)).toBe(true);
    expect(isXmlChar(0xde00)).toBe(true);
  });
});

describe('the Char production over full code points', () => {
  it('rejects a lone surrogate, which only a character reference can name', () => {
    // `&#xD800;` is the one way to get a value into a JavaScript string that
    // cannot be encoded back to UTF-8, which would break the round trip this
    // whole package exists to guarantee.
    expect(isXmlCodePoint(0xd7ff)).toBe(true);
    expect(isXmlCodePoint(0xd800)).toBe(false);
    expect(isXmlCodePoint(0xdfff)).toBe(false);
    expect(isXmlCodePoint(0xe000)).toBe(true);
  });

  it('accepts the astral planes and stops at the end of Unicode', () => {
    expect(isXmlCodePoint(0x10000)).toBe(true);
    expect(isXmlCodePoint(0x1f600)).toBe(true);
    expect(isXmlCodePoint(0x10ffff)).toBe(true);
    expect(isXmlCodePoint(0x110000)).toBe(false);
  });
});

describe('end-of-line normalization', () => {
  it('turns CRLF and a lone CR into LF, per §2.11', () => {
    expect(normalizeLineEndings('a\r\nb')).toBe('a\nb');
    expect(normalizeLineEndings('a\rb')).toBe('a\nb');
    expect(normalizeLineEndings('a\nb')).toBe('a\nb');
    expect(normalizeLineEndings('a\r\r\nb')).toBe('a\n\nb');
  });

  it('returns the same string when there is nothing to do', () => {
    const untouched = 'no carriage returns here';
    expect(normalizeLineEndings(untouched)).toBe(untouched);
  });

  it('is a transformation the source must survive', () => {
    // Writing `line1\r\nline2` into an <a:t> and asking PowerPoint for the text
    // back yields `line1\nline2`, so PowerPoint performs this too. The point of
    // keeping spans is that we can perform it *and* still emit the CRLF.
    const source = '<a:t>line1\r\nline2</a:t>';
    expect(normalizeLineEndings(source)).not.toBe(source);
    expect(source.slice(5, source.indexOf('</'))).toBe('line1\r\nline2');
  });
});

describe('whitespace runs', () => {
  it('spots an inter-element formatting run', () => {
    //             0123 4 5 6 7
    const source = '<a>\r\n  <b/></a>';
    expect(isAllWhitespace(source, 3, 7)).toBe(true);
    expect(isAllWhitespace(source, 3, 8)).toBe(false); // reaches the "<"
    expect(isAllWhitespace(source, 4, 4)).toBe(true); // an empty run, vacuously
  });
});

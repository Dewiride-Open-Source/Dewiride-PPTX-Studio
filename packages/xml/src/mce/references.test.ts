import { describe, expect, it } from 'vitest';
import { isXmlError } from '../errors.js';
import {
  decodeCharacterData,
  decodeReference,
  isLiteralRun,
  normalizeAttributeValue,
} from './references.js';

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isXmlError(error)) return error.code;
    return 'NOT_AN_XML_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

const text = (s: string): string => decodeCharacterData(s, 0, s.length);
const attr = (s: string): string => normalizeAttributeValue(s, 0, s.length);

describe('the five entities that exist', () => {
  it('decodes each one', () => {
    expect(text('&amp;&lt;&gt;&quot;&apos;')).toBe('&<>"\'');
  });

  it('refuses anything else, as PowerPoint does', () => {
    // Verified: `&nbsp;` inside an <a:t> makes PowerPoint refuse the package
    // with 0x80070570. Without a DTD there is nowhere for another entity to
    // have been declared, and we reject a DOCTYPE outright.
    expect(codeOf(() => text('a&nbsp;b'))).toBe('ERR_UNKNOWN_ENTITY');
    expect(codeOf(() => text('&copy;'))).toBe('ERR_UNKNOWN_ENTITY');
  });

  it('refuses a bare ampersand', () => {
    expect(codeOf(() => text('Smith & Sons'))).toBe('ERR_MALFORMED_XML');
  });

  it('bounds the scan for the closing semicolon, and says so', () => {
    // Without a cap, one stray "&" in a megabyte of text costs a search to the
    // end of the part; with many, the behaviour turns quadratic. But "there is
    // no semicolon within the cap" is a different failure from "there is no
    // semicolon at all", and reporting the first as a bare ampersand sends the
    // reader hunting for an "&" that is correctly written.
    expect(codeOf(() => text('&' + 'x'.repeat(500_000) + ';'))).toBe('ERR_LIMIT_EXCEEDED');
    expect(codeOf(() => text('Smith & Sons'))).toBe('ERR_MALFORMED_XML');
  });

  it('will not read a reference past the end of the run it was given', () => {
    // decodeReference used to scan to the cap regardless of the caller's `end`,
    // so it read straight through a closing quote or the "<" that ends a text
    // node. `<r a="&amp" b="x;"/>` reported: no entity named `amp" b="x`.
    expect(codeOf(() => decodeCharacterData('a&amp;b', 0, 2))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => normalizeAttributeValue('&amp;xx', 0, 2))).toBe('ERR_MALFORMED_XML');
  });
});

describe('character references', () => {
  it('decodes decimal and hexadecimal, with the digits in either case', () => {
    expect(text('&#72;&#105;')).toBe('Hi');
    expect(text('&#x48;&#x69;')).toBe('Hi');
    expect(text('&#xAb;')).toBe('«');
  });

  it('refuses "&#X" - the x is a terminal, not a digit', () => {
    // `CharRef ::= '&#x' [0-9a-fA-F]+ ';'` in a case-sensitive grammar. Only
    // the digits are case-insensitive.
    expect(codeOf(() => text('&#X48;'))).toBe('ERR_MALFORMED_XML');
  });

  it('decodes an astral reference into a surrogate pair', () => {
    // Verified against PowerPoint: `&#x1F600;` in an <a:t> renders as 😀.
    expect(text('&#x1F600;ok')).toBe('\u{1F600}ok');
    expect(text('&#x1F600;').length).toBe(2);
  });

  it('refuses a reference to a lone surrogate', () => {
    // The one construct that can put a value into a JavaScript string that
    // cannot be encoded back to UTF-8.
    expect(codeOf(() => text('&#xD800;'))).toBe('ERR_INVALID_CHARACTER');
    expect(codeOf(() => text('&#55296;'))).toBe('ERR_INVALID_CHARACTER');
  });

  it('refuses a reference to a character XML does not permit', () => {
    expect(codeOf(() => text('&#0;'))).toBe('ERR_INVALID_CHARACTER');
    expect(codeOf(() => text('&#xB;'))).toBe('ERR_INVALID_CHARACTER');
    expect(codeOf(() => text('&#xFFFE;'))).toBe('ERR_INVALID_CHARACTER');
    expect(codeOf(() => text('&#x110000;'))).toBe('ERR_INVALID_CHARACTER');
  });

  it('refuses a malformed reference', () => {
    expect(codeOf(() => text('&#;'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => text('&#xZZ;'))).toBe('ERR_MALFORMED_XML');
    expect(codeOf(() => text('&#12a;'))).toBe('ERR_MALFORMED_XML');
  });

  it('reports where the reference ended', () => {
    expect(decodeReference('x&amp;y', 1)).toEqual({ text: '&', end: 6 });
  });
});

describe('the two tiers of normalization, which is where this gets subtle', () => {
  it('turns a literal tab, newline or carriage return in an attribute into a space', () => {
    // XML 1.0 §3.3.3, the CDATA case - which is every attribute we will ever
    // see, because declaring one as anything else requires a DTD.
    expect(attr('a\tb')).toBe('a b');
    expect(attr('a\nb')).toBe('a b');
    expect(attr('a\rb')).toBe('a b');
  });

  it('leaves a character reference to the same character alone', () => {
    // This is the trap. `name="a&#9;b"` and `name="a<TAB>b"` are two different
    // values, and an implementation that expands references first and then
    // squashes whitespace collapses them into one.
    expect(attr('a&#9;b')).toBe('a\tb');
    expect(attr('a&#xA;b')).toBe('a\nb');
    expect(attr('a&#xD;b')).toBe('a\rb');
    expect(attr('a\tb')).not.toBe(attr('a&#9;b'));
  });

  it('treats CRLF in an attribute as one line break and so one space', () => {
    expect(attr('a\r\nb')).toBe('a b');
    expect(attr('a\r\n\r\nb')).toBe('a  b');
  });

  it('does not touch an ordinary space', () => {
    expect(attr('Rectangle 3')).toBe('Rectangle 3');
    expect(attr('  padded  ')).toBe('  padded  ');
  });
});

describe('character data', () => {
  it('normalizes line endings but keeps a referenced carriage return', () => {
    // Same asymmetry as attributes, one tier up: §2.11 rewrites literal line
    // breaks, §4.6 expands references, and a reference is not a line break.
    expect(text('line1\r\nline2')).toBe('line1\nline2');
    expect(text('line1\rline2')).toBe('line1\nline2');
    expect(text('line1&#xD;line2')).toBe('line1\rline2');
  });

  it('keeps a literal tab, unlike an attribute value', () => {
    // Verified against PowerPoint: a literal tab inside an <a:t> comes back as
    // a tab, and survives a resave.
    expect(text('a\tb')).toBe('a\tb');
    expect(attr('a\tb')).toBe('a b');
  });

  it('mixes literal text and references in one run', () => {
    expect(text('a &amp; b &lt;c&gt; d')).toBe('a & b <c> d');
  });
});

describe('the fast path', () => {
  it('recognises a run that is already its own decoding', () => {
    expect(isLiteralRun('plain text', 0, 10, false)).toBe(true);
    expect(isLiteralRun('a&amp;b', 0, 7, false)).toBe(false);
    expect(isLiteralRun('a\rb', 0, 3, false)).toBe(false);
  });

  it('is stricter for attributes, because §3.3.3 rewrites more', () => {
    expect(isLiteralRun('a\tb', 0, 3, false)).toBe(true);
    expect(isLiteralRun('a\tb', 0, 3, true)).toBe(false);
    expect(isLiteralRun('a b', 0, 3, true)).toBe(true);
  });

  it('agrees with the slow path wherever it claims the fast one', () => {
    for (const value of ['plain', 'a b', '', '0', 'Rectangle 3', 'café']) {
      if (isLiteralRun(value, 0, value.length, true)) expect(attr(value)).toBe(value);
      if (isLiteralRun(value, 0, value.length, false)) expect(text(value)).toBe(value);
    }
  });
});

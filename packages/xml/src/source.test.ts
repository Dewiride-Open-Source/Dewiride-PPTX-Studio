import { describe, expect, it } from 'vitest';
import { isXmlError } from './errors.js';
import { decodeXmlSource, encodeXmlSource } from './source.js';

const encoder = new TextEncoder();

function codeOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    if (isXmlError(error)) return error.code;
    return 'NOT_AN_XML_ERROR: ' + (error instanceof Error ? error.constructor.name : typeof error);
  }
  return 'DID_NOT_THROW';
}

function messageOf(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  return '(did not throw)';
}

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

describe('decoding a part', () => {
  it('round-trips to the same bytes, which is what the whole package rests on', () => {
    // Not a happy accident: valid UTF-8 and sequences of Unicode scalar values
    // are in bijection, so decode-then-encode is lossless by construction.
    // Measured across all 2834 XML parts of the 37-package corpus as well.
    for (const text of [
      DECLARATION + '<p:sld/>',
      '<a:t>café — “£” • 日本語</a:t>',
      '<a:t>😀 astral</a:t>',
      '<a xmlns="urn:x"/>',
    ]) {
      const bytes = encoder.encode(text);
      expect(encodeXmlSource(decodeXmlSource(bytes).text)).toEqual(bytes);
    }
  });

  it('keeps a byte order mark instead of silently eating it', () => {
    // `TextDecoder`'s `ignoreBOM` option is named backwards: the default,
    // `false`, means "consume the BOM and do not emit it", i.e. delete the
    // first three bytes. 1037 of the 2834 XML parts in our corpus carry one, so
    // the default would corrupt 37% of them - invisibly, until an export is
    // compared byte-for-byte against its input.
    const bytes = encoder.encode('﻿' + DECLARATION + '<p:sld/>');
    const source = decodeXmlSource(bytes);
    expect(source.bom).toBe(true);
    expect(source.text.charCodeAt(0)).toBe(0xfeff);
    expect(encodeXmlSource(source.text)).toEqual(bytes);

    // The trap itself, so the assertion above is not mistaken for a tautology.
    expect(new TextDecoder('utf-8').decode(bytes).charCodeAt(0)).not.toBe(0xfeff);
  });

  it('reports no BOM when there is none', () => {
    expect(decodeXmlSource(encoder.encode('<a/>')).bom).toBe(false);
  });

  it('handles an empty part without throwing something untyped', () => {
    expect(decodeXmlSource(new Uint8Array(0)).text).toBe('');
  });
});

describe('encodings we do not implement', () => {
  it('refuses a UTF-16 part and names the encoding', () => {
    // PowerPoint opens this: a slide part written UTF-16LE with a matching
    // declaration reads back correctly. We do not, and say so.
    const utf16 = new Uint8Array([0xff, 0xfe, 0x3c, 0x00, 0x61, 0x00, 0x2f, 0x00, 0x3e, 0x00]);
    expect(codeOf(() => decodeXmlSource(utf16))).toBe('ERR_UNSUPPORTED_ENCODING');
    expect(messageOf(() => decodeXmlSource(utf16))).toContain('UTF-16LE');
  });

  it('spots UTF-16 with no byte order mark', () => {
    // XML 1.0 Appendix F: a document begins with `<?xml` or an element, so a
    // NUL in the first two bytes can only mean a 16-bit encoding.
    const be = new Uint8Array([0x00, 0x3c, 0x00, 0x61]);
    expect(messageOf(() => decodeXmlSource(be))).toContain('UTF-16BE');
  });

  it('refuses bytes that are not valid UTF-8, as PowerPoint does', () => {
    // A lone 0xE9 is "é" in windows-1252 and invalid as UTF-8. PowerPoint
    // refuses such a part with 0x80070570, so refusing it costs us no file.
    const latin1 = encoder.encode('<a:t>caf.</a:t>');
    latin1[latin1.length - 8] = 0xe9;
    expect(codeOf(() => decodeXmlSource(latin1))).toBe('ERR_UNSUPPORTED_ENCODING');
  });

  it('names the declared encoding when the declaration explains the failure', () => {
    const declared = encoder.encode('<?xml version="1.0" encoding="windows-1252"?><a>.</a>');
    declared[declared.indexOf(0x2e)] = 0xe9;
    expect(messageOf(() => decodeXmlSource(declared))).toContain('windows-1252');
  });

  it('accepts a part labelled windows-1252 whose content is ASCII', () => {
    // The lenient half of the same rule. Such a part decodes as UTF-8 without
    // complaint and re-encodes byte-identically, so refusing it on the strength
    // of a label would lose a file we handle perfectly.
    const bytes = encoder.encode('<?xml version="1.0" encoding="windows-1252"?><a>plain</a>');
    expect(decodeXmlSource(bytes).text).toContain('plain');
    expect(encodeXmlSource(decodeXmlSource(bytes).text)).toEqual(bytes);
  });

  it('refuses UTF-8 bytes that claim to be UTF-16', () => {
    // A document that contradicts itself: there is no reading under which the
    // label and the bytes agree, because decoding these as UTF-16 yields
    // garbage. PowerPoint refuses it. That is the distinction from the
    // windows-1252 case above, where ASCII content means the same thing under
    // either label and the file is perfectly readable.
    const lying = encoder.encode('<?xml version="1.0" encoding="UTF-16"?><a/>');
    expect(codeOf(() => decodeXmlSource(lying))).toBe('ERR_UNSUPPORTED_ENCODING');
  });

  it('refuses a truncated multi-byte sequence rather than substituting U+FFFD', () => {
    // Substitution is the default `TextDecoder` behaviour and it is silent data
    // loss: the replacement character re-encodes to three different bytes.
    const truncated = new Uint8Array([
      ...encoder.encode('<a>'),
      0xe2,
      0x80,
      ...encoder.encode('</a>'),
    ]);
    expect(codeOf(() => decodeXmlSource(truncated))).toBe('ERR_UNSUPPORTED_ENCODING');
  });

  it('refuses an overlong encoding', () => {
    // 0xC0 0xAF is an overlong "/" - the classic path-traversal smuggling
    // trick. PowerPoint refuses it too.
    const overlong = new Uint8Array([
      ...encoder.encode('<a>'),
      0xc0,
      0xaf,
      ...encoder.encode('</a>'),
    ]);
    expect(codeOf(() => decodeXmlSource(overlong))).toBe('ERR_UNSUPPORTED_ENCODING');
  });
});

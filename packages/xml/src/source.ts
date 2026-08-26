/**
 * Getting from the bytes of a part to the string the tokenizer scans.
 *
 * This is the only place in the package that knows bytes exist, and it is
 * short because of one measured fact: for **all 2834 XML parts in our
 * 37-package corpus**, `TextEncoder.encode(TextDecoder.decode(bytes))` is
 * byte-for-byte identical to `bytes`.
 *
 * That is not luck. Valid UTF-8 and sequences of Unicode scalar values are in
 * bijection - UTF-8 admits no overlong forms and no encoded surrogates, and
 * `TextEncoder` always emits the shortest form - so the round trip is lossless
 * by construction for any input that decodes without error. The measurement
 * confirms the reasoning rather than substituting for it.
 *
 * The consequence is the central design decision of this package: **the
 * tokenizer works in string space, and a span is a pair of UTF-16 code-unit
 * indices into the decoded source.** See `docs/adr/0004-the-xml-layer.md`.
 */

import { XmlError } from './errors.js';

/** Encodings this package can decode losslessly. */
export type XmlEncoding = 'utf-8';

/** A decoded part, plus what is needed to encode it back byte-identically. */
export interface XmlSource {
  /**
   * The decoded document, **including a leading U+FEFF if the bytes carried a
   * byte order mark**. Every span in the tree indexes into this string.
   */
  readonly text: string;
  readonly encoding: XmlEncoding;
  /** True if `text` begins with U+FEFF because the bytes began with EF BB BF. */
  readonly bom: boolean;
  /** Length of the original byte sequence, for diagnostics and budgeting. */
  readonly byteLength: number;
}

/**
 * `TextDecoder`'s `ignoreBOM` option is named backwards and it matters here.
 *
 * `ignoreBOM: false` - the default - means "consume the BOM and do not emit
 * it", so the decoder silently deletes the first three bytes. **1037 of the
 * 2834 XML parts in our corpus carry a UTF-8 BOM**, so a default `TextDecoder`
 * would corrupt 37% of them, and the corruption is invisible until an export
 * is compared byte-for-byte against its input. `ignoreBOM: true` means "treat
 * it as an ordinary U+FEFF", which is what we want.
 */
const DECODER_OPTIONS = { fatal: true, ignoreBOM: true } as const;

/** Pulled out of the declaration before we know we can decode the whole part. */
const ENCODING_PATTERN = /<\?xml\s[^?]*?encoding\s*=\s*("([^"]*)"|'([^']*)')/;

/**
 * Read the `encoding` pseudo-attribute without decoding the part.
 *
 * Only ever used to name the encoding in an error message, so it may be
 * approximate - but it has to work on bytes we have already failed to decode,
 * which is why it reads them as Latin-1 and skips NULs rather than going
 * anywhere near a `TextDecoder`.
 */
function sniffLabel(bytes: Uint8Array): string | undefined {
  const limit = Math.min(bytes.length, 256);
  let head = '';
  for (let i = 0; i < limit; i++) {
    const b = bytes[i]!;
    if (b === 0) continue; // so a UTF-16 declaration still reads
    head += String.fromCharCode(b);
  }
  const match = ENCODING_PATTERN.exec(head);
  return match ? (match[2] ?? match[3]) : undefined;
}

/**
 * Identify a byte order mark, or a UTF-16 document that has none.
 *
 * The no-BOM UTF-16 check is XML 1.0 Appendix F: a document must begin with
 * `<?xml` or an element, so a NUL as the first or second byte can only mean a
 * 16-bit encoding.
 */
function detectByteOrderMark(bytes: Uint8Array): string | undefined {
  if (bytes.length >= 4) {
    if (bytes[0] === 0x00 && bytes[1] === 0x00 && bytes[2] === 0xfe && bytes[3] === 0xff)
      return 'UTF-32BE';
    if (bytes[0] === 0xff && bytes[1] === 0xfe && bytes[2] === 0x00 && bytes[3] === 0x00)
      return 'UTF-32LE';
  }
  if (bytes.length >= 2) {
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return 'UTF-16BE';
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return 'UTF-16LE';
  }
  if (bytes.length >= 2 && bytes[0] === 0x3c && bytes[1] === 0x00) return 'UTF-16LE';
  if (bytes.length >= 2 && bytes[0] === 0x00 && bytes[1] === 0x3c) return 'UTF-16BE';
  return undefined;
}

function unsupported(label: string, declared: string | undefined): never {
  const suffix =
    declared !== undefined && declared.toLowerCase() !== label.toLowerCase()
      ? ' (the declaration says "' + declared + '")'
      : '';
  throw new XmlError(
    'ERR_UNSUPPORTED_ENCODING',
    'this part is encoded as ' +
      label +
      suffix +
      '; @pptx-studio/xml reads UTF-8 only. PowerPoint honours the encoding declaration and would open this file',
    { name: label },
  );
}

/**
 * Decode a part into the string the tokenizer scans.
 *
 * Refuses anything it cannot reproduce byte-for-byte, which is the same
 * position `@pptx-studio/opc` takes on the archive: better a typed error naming
 * the encoding than a silently mangled export. PowerPoint is more permissive
 * here than we are - it opens a slide part written in windows-1252 and reads
 * `café` correctly - and that gap is a known, tested limitation rather than an
 * oversight.
 */
export function decodeXmlSource(bytes: Uint8Array): XmlSource {
  const mark = detectByteOrderMark(bytes);
  if (mark !== undefined) unsupported(mark, sniffLabel(bytes));

  const bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;

  // A declaration naming a 16-bit encoding, over bytes that are plainly not
  // 16-bit, is a document that contradicts itself. PowerPoint refuses it, and
  // we agree - unlike the windows-1252 case just below, there is no reading
  // under which the two agree: decoding these bytes as UTF-16 yields garbage.
  const declared = sniffLabel(bytes);
  if (declared !== undefined && /^(utf-?16|ucs-?2)/i.test(declared)) {
    unsupported(declared.toUpperCase(), declared);
  }

  let text: string;
  try {
    text = new TextDecoder('utf-8', DECODER_OPTIONS).decode(bytes);
  } catch (cause) {
    const declared = sniffLabel(bytes);
    // A part labelled windows-1252 whose content happens to be pure ASCII
    // decodes as UTF-8 without complaint and round-trips perfectly, so we only
    // reach here for one that genuinely is not UTF-8.
    throw new XmlError(
      'ERR_UNSUPPORTED_ENCODING',
      declared !== undefined && !/^utf-?8$/i.test(declared)
        ? 'this part declares encoding "' +
            declared +
            '" and is not valid UTF-8; @pptx-studio/xml reads UTF-8 only'
        : 'this part is not valid UTF-8',
      { name: declared, cause },
    );
  }

  return { text, encoding: 'utf-8', bom, byteLength: bytes.length };
}

/**
 * Encode a source string back to bytes.
 *
 * Byte-identical to the input of {@link decodeXmlSource} whenever the document
 * has not been edited - the property sub-phase 0.5's round-trip gate rests on.
 */
export function encodeXmlSource(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

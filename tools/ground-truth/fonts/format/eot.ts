/**
 * Embedded OpenType - read and write.
 *
 * `ppt/fonts/fontN.fntdata` is an EOT file. Not a bare `.ttf`, and not Word's
 * ODTTF obfuscation, both of which are widely repeated and both of which are
 * wrong. This module is the instrument for sub-phase 0.7's experiments A and B:
 * A reads what PowerPoint wrote, B writes something and asks PowerPoint to read
 * it back.
 *
 * ## Layout
 *
 * All little-endian, unlike the SFNT it wraps.
 *
 * ```
 * 0x00 u32   EOTSize             total bytes of the EOT file
 * 0x04 u32   FontDataSize
 * 0x08 u32   Version             0x00010000 | 0x00020001 | 0x00020002
 * 0x0C u32   Flags               <- experiment A is entirely about this word
 * 0x10 u8[10] FontPANOSE
 * 0x1A u8    Charset             Windows charset byte; DEFAULT_CHARSET is 1
 * 0x1B u8    Italic
 * 0x1C u32   Weight              OS/2.usWeightClass
 * 0x20 u16   fsType              OS/2.fsType, copied honestly
 * 0x22 u16   MagicNumber         0x504C  ("PL", for Phil Lehman)
 * 0x24 u32   UnicodeRange1..4
 * 0x34 u32   CodePageRange1..2
 * 0x3C u32   CheckSumAdjustment  head.checkSumAdjustment
 * 0x40 u32   Reserved1..4        must be zero
 * 0x50 u16   Padding1            must be zero
 * 0x52 u16   FamilyNameSize      bytes, not characters
 *      ...   FamilyName          UTF-16LE, no terminator
 *      u16   Padding2 = 0
 *      u16   StyleNameSize,  StyleName
 *      u16   Padding3 = 0
 *      u16   VersionNameSize, VersionName
 *      u16   Padding4 = 0
 *      u16   FullNameSize, FullName
 *   version >= 0x00020001 adds: Padding5, RootStringSize, RootString
 *   version >= 0x00020002 adds: RootStringCheckSum, EUDCCodePage, Padding6,
 *                              SignatureSize, Signature, EUDCFlags,
 *                              EUDCFontSize, EUDCFontData
 * ```
 *
 * The font data is found as `bytes[EOTSize - FontDataSize .. EOTSize]`. That is
 * the spec-sanctioned way and it is not laziness: it means a reader never has
 * to walk the variable-length name tail, which is where the version differences
 * live and where a third-party writer is most likely to have got the padding
 * wrong.
 */

/**
 * The flags at 0x0C. The one that matters is `TTCOMPRESSED`: it means the
 * payload is MicroType Express (LZCOMP) compressed, and there is no MTX decoder
 * in JavaScript or WebAssembly anywhere.
 */
export const EOT_FLAG = {
  SUBSET: 0x00000001,
  TTCOMPRESSED: 0x00000004,
  FAILIFVARIATIONSIMULATED: 0x00000010,
  EMBEDEUDC: 0x00000020,
  VALIDATIONTESTS: 0x00000040,
  WEBOBJECT: 0x00000080,
  XORENCRYPTDATA: 0x10000000,
} as const;

export const EOT_MAGIC = 0x504c;
export const EOT_MAGIC_OFFSET = 0x22;

export const EOT_VERSION_1 = 0x00010000;
export const EOT_VERSION_2_1 = 0x00020001;
export const EOT_VERSION_2_2 = 0x00020002;

export interface EotHeader {
  readonly eotSize: number;
  readonly fontDataSize: number;
  readonly version: number;
  readonly flags: number;
  readonly panose: Uint8Array;
  readonly charset: number;
  readonly italic: number;
  readonly weight: number;
  readonly fsType: number;
  readonly magicNumber: number;
  readonly unicodeRange: readonly [number, number, number, number];
  readonly codePageRange: readonly [number, number];
  readonly checkSumAdjustment: number;
  readonly reserved: readonly [number, number, number, number];
  readonly familyName: string;
  readonly styleName: string;
  readonly versionName: string;
  readonly fullName: string;
  readonly rootString: string | undefined;
  /** Where the name tail ends. Should equal `eotSize - fontDataSize` for v1. */
  readonly headerEnd: number;
}

/** Every flag bit that is set, by name, plus any bit we have no name for. */
export function flagNames(flags: number): string[] {
  const names: string[] = [];
  let known = 0;
  for (const [name, bit] of Object.entries(EOT_FLAG)) {
    if ((flags & bit) !== 0) names.push(name);
    known |= bit;
  }
  const unknown = flags & ~known;
  if (unknown !== 0) names.push(`UNKNOWN(0x${unknown.toString(16).padStart(8, '0')})`);
  return names;
}

function utf16le(bytes: Uint8Array, at: number, length: number): string {
  let text = '';
  for (let i = 0; i + 1 < length; i += 2) {
    text += String.fromCharCode(bytes[at + i]! | (bytes[at + i + 1]! << 8));
  }
  return text;
}

export function readEot(bytes: Uint8Array): EotHeader {
  if (bytes.length < 0x54) throw new Error(`too short to be EOT: ${String(bytes.length)} bytes`);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const magicNumber = view.getUint16(EOT_MAGIC_OFFSET, true);
  if (magicNumber !== EOT_MAGIC) {
    throw new Error(
      `bad EOT magic at 0x22: 0x${magicNumber.toString(16)} (expected 0x${EOT_MAGIC.toString(16)})`,
    );
  }

  const version = view.getUint32(0x08, true);

  let at = 0x50;
  const padding1 = view.getUint16(at, true);
  at += 2;

  // Each name is `u16 size` then that many BYTES of UTF-16LE, then a u16 pad.
  const nextName = (): string => {
    const size = view.getUint16(at, true);
    at += 2;
    const text = utf16le(bytes, at, size);
    at += size;
    return text;
  };
  const skipPad = (): void => {
    at += 2;
  };

  const familyName = nextName();
  skipPad();
  const styleName = nextName();
  skipPad();
  const versionName = nextName();
  skipPad();
  const fullName = nextName();

  let rootString: string | undefined;
  if (version >= EOT_VERSION_2_1) {
    skipPad();
    rootString = nextName();
  }

  if (padding1 !== 0) throw new Error(`Padding1 is 0x${padding1.toString(16)}, must be zero`);

  return {
    eotSize: view.getUint32(0x00, true),
    fontDataSize: view.getUint32(0x04, true),
    version,
    flags: view.getUint32(0x0c, true),
    panose: bytes.slice(0x10, 0x1a),
    charset: bytes[0x1a]!,
    italic: bytes[0x1b]!,
    weight: view.getUint32(0x1c, true),
    fsType: view.getUint16(0x20, true),
    magicNumber,
    unicodeRange: [
      view.getUint32(0x24, true),
      view.getUint32(0x28, true),
      view.getUint32(0x2c, true),
      view.getUint32(0x30, true),
    ],
    codePageRange: [view.getUint32(0x34, true), view.getUint32(0x38, true)],
    checkSumAdjustment: view.getUint32(0x3c, true),
    reserved: [
      view.getUint32(0x40, true),
      view.getUint32(0x44, true),
      view.getUint32(0x48, true),
      view.getUint32(0x4c, true),
    ],
    familyName,
    styleName,
    versionName,
    fullName,
    rootString,
    headerEnd: at,
  };
}

/**
 * The wrapped font bytes.
 *
 * If `XORENCRYPTDATA` is set every byte is XORed with 0x50, and that must be
 * undone *before* decompression - doing it the other way round produces bytes
 * that look plausible and are not.
 */
export function eotFontData(bytes: Uint8Array, header: EotHeader): Uint8Array {
  const start = header.eotSize - header.fontDataSize;
  const data = bytes.subarray(start, header.eotSize);
  if ((header.flags & EOT_FLAG.XORENCRYPTDATA) === 0) return data;
  const out = new Uint8Array(data.length);
  for (let i = 0; i < data.length; i++) out[i] = data[i]! ^ 0x50;
  return out;
}

/* -------------------------------------------------------------------------- */
/* writing                                                                    */
/* -------------------------------------------------------------------------- */

export interface WriteEotOptions {
  readonly familyName: string;
  readonly styleName: string;
  readonly versionName: string;
  readonly fullName: string;
  readonly panose: Uint8Array;
  /** Windows charset byte. `DEFAULT_CHARSET` is 1. */
  readonly charset: number;
  readonly italic: boolean;
  readonly weight: number;
  /**
   * The real `OS/2.fsType`. `ttf2eot` leaves this zero, which falsely advertises
   * "Installable" - a claim the OpenType specification makes normative and which
   * we are not going to make on a font's behalf.
   */
  readonly fsType: number;
  readonly unicodeRange: readonly [number, number, number, number];
  readonly codePageRange: readonly [number, number];
  readonly checkSumAdjustment: number;
  /**
   * `EOT_VERSION_1` or `EOT_VERSION_2_2`. PowerPoint writes 2.2; whether it
   * *requires* 2.2 is one of the things experiment B is measuring.
   */
  readonly version?: number | undefined;
  /**
   * Include a terminating NUL inside the counted length of FamilyName and
   * StyleName. PowerPoint does this for those two names and not for
   * VersionName or FullName - an asymmetry with no obvious reason, which is
   * exactly the sort of thing a reader might be strict about.
   */
  readonly nulTerminateNames?: boolean | undefined;
}

function nameBytes(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
}

/**
 * The four bytes PowerPoint writes as RootStringCheckSum when the root string
 * is empty. They spell `BSGP`, which is not a checksum of anything - it is a
 * constant, and we copy it rather than invent one.
 */
const ROOT_STRING_CHECKSUM = [0x42, 0x53, 0x47, 0x50];
/** PowerPoint writes 1252 (Windows-1252) as EUDCCodePage. */
const EUDC_CODE_PAGE = 1252;

/**
 * Build an uncompressed EOT around an SFNT.
 *
 * Defaults to version 1, which is the smallest thing the format can express:
 * 0x00020001 and 0x00020002 exist to carry the RootString same-origin list and
 * the EUDC tail, neither of which means anything in a `.pptx`. Whether
 * PowerPoint agrees is what experiment B measures - so both shapes are
 * buildable from here.
 */
export function writeEot(sfnt: Uint8Array, options: WriteEotOptions): Uint8Array {
  if (!options.fullName.startsWith(options.familyName)) {
    // Internet Explorer's loader rejected fonts whose FullName did not begin
    // with the FamilyName. IE is gone; the invariant costs nothing and no
    // document says PowerPoint dropped it.
    throw new Error(
      `EOT FullName ${JSON.stringify(options.fullName)} must begin with FamilyName ` +
        `${JSON.stringify(options.familyName)}`,
    );
  }

  const version = options.version ?? EOT_VERSION_1;
  const nul = options.nulTerminateNames ?? false;

  // PowerPoint NUL-terminates FamilyName and StyleName but not VersionName or
  // FullName, and counts the NUL in the size.
  const family = nameBytes(options.familyName + (nul ? '\0' : ''));
  const style = nameBytes(options.styleName + (nul ? '\0' : ''));
  const versionName = nameBytes(options.versionName);
  const full = nameBytes(options.fullName);

  // Padding5 + RootStringSize, then, for 2.2 only, RootStringCheckSum,
  // EUDCCodePage, Padding6, SignatureSize, EUDCFlags, EUDCFontSize.
  const tailSize =
    version >= EOT_VERSION_2_1 ? (version >= EOT_VERSION_2_2 ? 2 + 2 + 20 : 2 + 2) : 0;

  const headerSize =
    0x52 + // fixed part, up to and including Padding1
    2 +
    family.length +
    2 + // Padding2
    2 +
    style.length +
    2 + // Padding3
    2 +
    versionName.length +
    2 + // Padding4
    2 +
    full.length +
    tailSize;

  const total = headerSize + sfnt.length;
  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);

  view.setUint32(0x00, total, true);
  view.setUint32(0x04, sfnt.length, true);
  view.setUint32(0x08, version, true);
  view.setUint32(0x0c, 0, true); // TTEMBED_RAW: no subset, no compression, no XOR
  out.set(options.panose.subarray(0, 10), 0x10);
  out[0x1a] = options.charset;
  out[0x1b] = options.italic ? 1 : 0;
  view.setUint32(0x1c, options.weight, true);
  view.setUint16(0x20, options.fsType, true);
  view.setUint16(0x22, EOT_MAGIC, true);
  for (let i = 0; i < 4; i++) view.setUint32(0x24 + i * 4, options.unicodeRange[i]!, true);
  view.setUint32(0x34, options.codePageRange[0], true);
  view.setUint32(0x38, options.codePageRange[1], true);
  view.setUint32(0x3c, options.checkSumAdjustment, true);
  // Reserved1..4 and Padding1 stay zero.

  let at = 0x52;
  const putName = (value: Uint8Array, pad: boolean): void => {
    view.setUint16(at, value.length, true);
    at += 2;
    out.set(value, at);
    at += value.length;
    if (pad) {
      view.setUint16(at, 0, true);
      at += 2;
    }
  };
  putName(family, true);
  putName(style, true);
  putName(versionName, true);
  putName(full, version >= EOT_VERSION_2_1);

  if (version >= EOT_VERSION_2_1) {
    view.setUint16(at, 0, true); // RootStringSize: no same-origin list
    at += 2;
  }
  if (version >= EOT_VERSION_2_2) {
    for (const byte of ROOT_STRING_CHECKSUM) out[at++] = byte;
    view.setUint32(at, EUDC_CODE_PAGE, true);
    at += 4;
    view.setUint16(at, 0, true); // Padding6
    at += 2;
    view.setUint16(at, 0, true); // SignatureSize
    at += 2;
    view.setUint32(at, 0, true); // EUDCFlags
    at += 4;
    view.setUint32(at, 0, true); // EUDCFontSize
    at += 4;
  }

  if (at !== headerSize)
    throw new Error(`header size mismatch: ${String(at)} vs ${String(headerSize)}`);
  out.set(sfnt, at);
  return out;
}

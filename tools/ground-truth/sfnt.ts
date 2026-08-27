/**
 * Just enough of the SFNT container to answer the ground-truth questions.
 *
 * Phase 8 gets a real reader in `@pptx-studio/fonts`. This one exists to
 * cross-examine what PowerPoint wrote: an EOT header carries a family name, a
 * weight, an italic flag and an `fsType`, and every one of those is a *copy* of
 * something inside the font it wraps. Reading both and comparing is how we find
 * out whether the copy can be trusted - which matters, because a reader that
 * believes the EOT header is much simpler than one that parses the SFNT.
 *
 * Byte order is big-endian throughout, which is the one thing everybody gets
 * wrong the first time.
 */

export interface SfntTable {
  readonly tag: string;
  readonly checksum: number;
  readonly offset: number;
  readonly length: number;
}

export interface Sfnt {
  /** `0x00010000` for TrueType outlines, `OTTO` for CFF. */
  readonly version: number;
  readonly tables: readonly SfntTable[];
  readonly bytes: Uint8Array;
}

export const SFNT_VERSION_TRUETYPE = 0x00010000;
export const SFNT_TAG_OTTO = 0x4f54544f;
export const SFNT_TAG_TTCF = 0x74746366;

export function readSfnt(bytes: Uint8Array): Sfnt {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const version = view.getUint32(0);
  if (version === SFNT_TAG_TTCF) throw new Error('TrueType Collection, not a single font');
  const numTables = view.getUint16(4);
  const tables: SfntTable[] = [];
  for (let i = 0; i < numTables; i++) {
    const at = 12 + i * 16;
    tables.push({
      tag: String.fromCharCode(bytes[at]!, bytes[at + 1]!, bytes[at + 2]!, bytes[at + 3]!),
      checksum: view.getUint32(at + 4),
      offset: view.getUint32(at + 8),
      length: view.getUint32(at + 12),
    });
  }
  return { version, tables, bytes };
}

export function table(font: Sfnt, tag: string): Uint8Array | undefined {
  const t = font.tables.find((x) => x.tag === tag);
  if (t === undefined) return undefined;
  return font.bytes.subarray(t.offset, t.offset + t.length);
}

/* -------------------------------------------------------------------------- */
/* name                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * The name IDs worth having a word for. ID 1 is the one PowerPoint writes into
 * `p:embeddedFont/@typeface` and into `a:latin/@typeface`; ID 6 is PostScript
 * and is *not* what matches.
 */
export const NAME_ID = {
  copyright: 0,
  family: 1,
  subfamily: 2,
  uniqueId: 3,
  fullName: 4,
  version: 5,
  postScript: 6,
  trademark: 7,
  typographicFamily: 16,
  typographicSubfamily: 17,
} as const;

/**
 * Read the `name` table into `nameID -> string`.
 *
 * Where a name ID appears more than once, Windows/Unicode BMP (platform 3,
 * encoding 1, language 0x0409) wins, because that is the record Windows itself
 * reads. Mac Roman records are decoded as Latin-1, which is wrong for the
 * high half and right for everything we will actually see.
 */
export function readNames(font: Sfnt): Map<number, string> {
  const name = table(font, 'name');
  const out = new Map<number, string>();
  if (name === undefined) return out;
  const view = new DataView(name.buffer, name.byteOffset, name.byteLength);
  const count = view.getUint16(2);
  const stringOffset = view.getUint16(4);

  // Higher score wins. Anything is better than nothing.
  const scoreOf = (platform: number, encoding: number, language: number): number => {
    if (platform === 3 && encoding === 1 && language === 0x0409) return 4;
    if (platform === 3 && encoding === 1) return 3;
    if (platform === 3) return 2;
    if (platform === 0) return 2;
    return 1;
  };
  const best = new Map<number, number>();

  for (let i = 0; i < count; i++) {
    const at = 6 + i * 12;
    const platform = view.getUint16(at);
    const encoding = view.getUint16(at + 2);
    const language = view.getUint16(at + 4);
    const nameId = view.getUint16(at + 6);
    const length = view.getUint16(at + 8);
    const offset = view.getUint16(at + 10);
    const raw = name.subarray(stringOffset + offset, stringOffset + offset + length);

    const utf16 = platform === 3 || platform === 0;
    let text = '';
    if (utf16) {
      for (let j = 0; j + 1 < raw.length; j += 2) {
        text += String.fromCharCode((raw[j]! << 8) | raw[j + 1]!);
      }
    } else {
      for (const byte of raw) text += String.fromCharCode(byte);
    }

    const score = scoreOf(platform, encoding, language);
    if (score >= (best.get(nameId) ?? -1)) {
      best.set(nameId, score);
      out.set(nameId, text);
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* OS/2 and head                                                              */
/* -------------------------------------------------------------------------- */

export interface Os2 {
  readonly version: number;
  /**
   * The embedding permission bits. Note the versioning trap: in OS/2 v0-v2 the
   * field is a set of independent permissions and the *least* restrictive wins,
   * so `fsType === 4` is not a reliable test for "preview and print only".
   */
  readonly fsType: number;
  readonly usWeightClass: number;
  readonly usWidthClass: number;
  readonly fsSelection: number;
  readonly panose: Uint8Array;
  readonly ulUnicodeRange: readonly [number, number, number, number];
  readonly ulCodePageRange: readonly [number, number] | undefined;
  readonly usFirstCharIndex: number;
  readonly usLastCharIndex: number;
}

export function readOs2(font: Sfnt): Os2 | undefined {
  const os2 = table(font, 'OS/2');
  if (os2 === undefined) return undefined;
  const view = new DataView(os2.buffer, os2.byteOffset, os2.byteLength);
  const version = view.getUint16(0);
  return {
    version,
    usWidthClass: view.getUint16(6),
    fsType: view.getUint16(8),
    usWeightClass: view.getUint16(4),
    panose: os2.subarray(32, 42),
    ulUnicodeRange: [
      view.getUint32(42),
      view.getUint32(46),
      view.getUint32(50),
      view.getUint32(54),
    ],
    fsSelection: view.getUint16(62),
    usFirstCharIndex: view.getUint16(64),
    usLastCharIndex: view.getUint16(66),
    ulCodePageRange:
      version >= 1 && os2.length >= 86 ? [view.getUint32(78), view.getUint32(82)] : undefined,
  };
}

export interface Head {
  readonly unitsPerEm: number;
  readonly checkSumAdjustment: number;
  readonly macStyle: number;
  readonly indexToLocFormat: number;
  readonly xMin: number;
  readonly yMin: number;
  readonly xMax: number;
  readonly yMax: number;
}

export function readHead(font: Sfnt): Head | undefined {
  const head = table(font, 'head');
  if (head === undefined) return undefined;
  const view = new DataView(head.buffer, head.byteOffset, head.byteLength);
  return {
    checkSumAdjustment: view.getUint32(8),
    unitsPerEm: view.getUint16(18),
    xMin: view.getInt16(36),
    yMin: view.getInt16(38),
    xMax: view.getInt16(40),
    yMax: view.getInt16(42),
    macStyle: view.getUint16(44),
    indexToLocFormat: view.getInt16(50),
  };
}

/** The number of glyphs, from `maxp`. */
export function numGlyphs(font: Sfnt): number | undefined {
  const maxp = table(font, 'maxp');
  if (maxp === undefined) return undefined;
  return new DataView(maxp.buffer, maxp.byteOffset, maxp.byteLength).getUint16(4);
}

/**
 * Every code point the `cmap` maps, read from the format 4 and format 12
 * subtables. Used to tell a subset font from a whole one - which is the
 * question experiment A asks and the file size only hints at.
 */
export function readCoverage(font: Sfnt): Set<number> {
  const covered = new Set<number>();
  const cmap = table(font, 'cmap');
  if (cmap === undefined) return covered;
  const view = new DataView(cmap.buffer, cmap.byteOffset, cmap.byteLength);
  const numSubtables = view.getUint16(2);

  for (let i = 0; i < numSubtables; i++) {
    const offset = view.getUint32(4 + i * 8 + 4);
    if (offset + 4 > cmap.length) continue;
    const format = view.getUint16(offset);

    if (format === 4) {
      const segCountX2 = view.getUint16(offset + 6);
      const segCount = segCountX2 / 2;
      const endAt = offset + 14;
      const startAt = endAt + segCountX2 + 2;
      const deltaAt = startAt + segCountX2;
      const rangeAt = deltaAt + segCountX2;
      for (let s = 0; s < segCount; s++) {
        const end = view.getUint16(endAt + s * 2);
        const start = view.getUint16(startAt + s * 2);
        if (start > end || start === 0xffff) continue;
        const rangeOffset = view.getUint16(rangeAt + s * 2);
        for (let cp = start; cp <= end; cp++) {
          if (rangeOffset === 0) {
            covered.add(cp);
          } else {
            const glyphAt = rangeAt + s * 2 + rangeOffset + (cp - start) * 2;
            if (glyphAt + 1 < cmap.length && view.getUint16(glyphAt) !== 0) covered.add(cp);
          }
        }
      }
    } else if (format === 12) {
      const nGroups = view.getUint32(offset + 12);
      for (let g = 0; g < nGroups; g++) {
        const at = offset + 16 + g * 12;
        const start = view.getUint32(at);
        const end = view.getUint32(at + 4);
        // Guard against a malformed group claiming the whole plane.
        for (let cp = start; cp <= end && cp - start < 0x10000; cp++) covered.add(cp);
      }
    }
  }
  return covered;
}

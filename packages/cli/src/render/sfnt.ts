/**
 * The four tables a measurement needs, read out of a font file.
 *
 * Not a font library: nothing here decodes an outline, and a CFF font is read
 * as happily as a glyf one because advances live in `hmtx` either way. What it
 * answers is the question the browser answers with `measureText`, which is the
 * one thing `render-svg` cannot get in Node. Every reading below was measured
 * against Chromium in experiment T13 rather than taken from the specification.
 * ADR 0042.
 */

import { RenderError } from './errors.js';

const OTTO = 0x4f54544f;
const TRUE_TYPE = 0x00010000;
const TTCF = 0x74746366;

/** `OS/2.fsSelection` bit 7: the typographic metrics are the ones to use. */
const USE_TYPO_METRICS = 0x0080;

/**
 * The rasteriser a Chromium reads a face's vertical metrics through.
 *
 * T13 scored `usWin` 30/30 against DirectWrite on fonts built to disagree; T14
 * scored `hhea` 94/94 against FreeType and 248/248 against CoreText on real
 * faces. One reader cannot answer for all three. ADR 0045, ADR 0053.
 */
export type FontBackend = 'directwrite' | 'freetype' | 'coretext';

/** FreeType on Linux, CoreText on macOS, DirectWrite elsewhere. ADR 0045, ADR 0053. */
export function backendFor(platform: string): FontBackend {
  if (platform === 'linux') return 'freetype';
  if (platform === 'darwin') return 'coretext';
  return 'directwrite';
}

/** One ascent and descent in font units, the descent positive below the baseline. */
export interface MetricPair {
  readonly ascent: number;
  readonly descent: number;
}

/** Every pair a browser could report a face box from, each as its table wrote it. */
export interface MetricCandidates {
  readonly hhea: MetricPair;
  /** `undefined` on a face with no `OS/2` table long enough to hold version 0. */
  readonly usWin: MetricPair | undefined;
  readonly sTypo: MetricPair | undefined;
  /** `OS/2.fsSelection` bit 7, USE_TYPO_METRICS. */
  readonly useTypoMetrics: boolean;
}

export interface FaceMetrics {
  /** `head.unitsPerEm`, the denominator of every number here. */
  readonly unitsPerEm: number;
  /** Above the baseline, in font units. */
  readonly ascent: number;
  /** Below the baseline, in font units, positive. */
  readonly descent: number;
  /** Which table `ascent` and `descent` came out of, for the diagnostics. */
  readonly source: 'hhea' | 'usWin' | 'sTypo';
  /** The rasteriser `source` was chosen for, which also fixes how the box rounds. */
  readonly backend: FontBackend;
  /** The `BASE` `ideo` coordinate under `DFLT`, in font units below the baseline. */
  readonly ideographic: number | undefined;
  /** The pairs `ascent` and `descent` were chosen from, so a rival reading is scorable. */
  readonly candidates: MetricCandidates;
}

export interface Face {
  /** `name` ID 1, which is what PowerPoint matches a typeface on. */
  readonly family: string;
  /** `name` ID 2: `Regular`, `Bold`, `Italic`, `Bold Italic`, or a style name. */
  readonly subfamily: string;
  /** `name` ID 16, present when the family splits into more than four styles. */
  readonly typographicFamily: string | undefined;
  readonly metrics: FaceMetrics;
  readonly bold: boolean;
  readonly italic: boolean;
  /** Code point to advance, in font units. */
  advanceOf(codePoint: number): number | undefined;
  /** The adjustment between two code points, in font units. Usually zero. */
  kernBetween(left: number, right: number): number;
}

interface Reader {
  readonly bytes: Uint8Array;
  readonly view: DataView;
}

function readerOf(bytes: Uint8Array): Reader {
  return { bytes, view: new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength) };
}

function u8(r: Reader, at: number): number {
  return r.view.getUint8(at);
}
function u16(r: Reader, at: number): number {
  return r.view.getUint16(at);
}
function i16(r: Reader, at: number): number {
  return r.view.getInt16(at);
}
function u32(r: Reader, at: number): number {
  return r.view.getUint32(at);
}

function tagAt(r: Reader, at: number): string {
  return String.fromCharCode(u8(r, at), u8(r, at + 1), u8(r, at + 2), u8(r, at + 3));
}

function unreadable(what: string, subject: string): never {
  throw new RenderError('CLI_FONT_UNREADABLE', what, subject);
}

/* -------------------------------------------------------------------------- */
/* the table directory                                                        */
/* -------------------------------------------------------------------------- */

type Tables = ReadonlyMap<string, Uint8Array>;

function directoryAt(bytes: Uint8Array, start: number, subject: string): Tables {
  const r = readerOf(bytes);
  if (start + 12 > bytes.length) unreadable('font directory runs past the file', subject);
  const count = u16(r, start + 4);
  const out = new Map<string, Uint8Array>();
  for (let i = 0; i < count; i++) {
    const at = start + 12 + i * 16;
    if (at + 16 > bytes.length) unreadable('table directory runs past the file', subject);
    const tag = tagAt(r, at);
    const offset = u32(r, at + 8);
    const length = u32(r, at + 12);
    // A length that runs past the end is a truncated file, not a fatal one:
    // clamping keeps a readable `name` usable when a trailing table is short.
    if (offset < bytes.length)
      out.set(tag, bytes.subarray(offset, Math.min(offset + length, bytes.length)));
  }
  return out;
}

/** Every font in the file: one for a plain font, several for a collection. */
export function fontsIn(bytes: Uint8Array, subject: string): readonly Tables[] {
  if (bytes.length < 12) unreadable('too short to be a font', subject);
  const r = readerOf(bytes);
  const version = u32(r, 0);
  if (version === TTCF) {
    const count = u32(r, 8);
    const out: Tables[] = [];
    for (let i = 0; i < count; i++) out.push(directoryAt(bytes, u32(r, 12 + i * 4), subject));
    return out;
  }
  if (version !== TRUE_TYPE && version !== OTTO) {
    unreadable(`0x${version.toString(16).padStart(8, '0')} is not an SFNT signature`, subject);
  }
  return [directoryAt(bytes, 0, subject)];
}

function required(tables: Tables, tag: string, subject: string): Uint8Array {
  const table = tables.get(tag);
  if (table === undefined) unreadable(`no ${tag} table`, subject);
  return table;
}

/* -------------------------------------------------------------------------- */
/* name                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * The names a face is known by, decoded from the Windows platform records.
 *
 * Platform 3 encoding 1 is UTF-16BE and is what every font shipped for Windows
 * carries; platform 1 is Mac Roman and is read only where a face has no
 * Windows record at all, which is rare enough to be worth the four lines.
 */
function namesOf(tables: Tables, subject: string): ReadonlyMap<number, string> {
  const table = required(tables, 'name', subject);
  const r = readerOf(table);
  if (table.length < 6) unreadable('name table is truncated', subject);
  const count = u16(r, 2);
  const storage = u16(r, 4);
  const out = new Map<number, string>();
  const seenWindows = new Set<number>();
  for (let i = 0; i < count; i++) {
    const at = 6 + i * 12;
    if (at + 12 > table.length) break;
    const platform = u16(r, at);
    const encoding = u16(r, at + 2);
    const nameId = u16(r, at + 6);
    const length = u16(r, at + 8);
    const offset = storage + u16(r, at + 10);
    if (offset + length > table.length) continue;
    const raw = table.subarray(offset, offset + length);
    if (platform === 3 && (encoding === 1 || encoding === 0)) {
      let text = '';
      for (let j = 0; j + 1 < raw.length; j += 2) {
        text += String.fromCharCode(((raw[j] ?? 0) << 8) | (raw[j + 1] ?? 0));
      }
      out.set(nameId, text);
      seenWindows.add(nameId);
    } else if (platform === 1 && encoding === 0 && !seenWindows.has(nameId)) {
      out.set(nameId, String.fromCharCode(...raw));
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* cmap                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * The best Unicode subtable in the font.
 *
 * Preference order is the one every shaper uses: a format 12 full-repertoire
 * table beats a format 4 BMP one, because a face that has both maps astral
 * code points only in the first.
 */
function cmapOf(tables: Tables, subject: string): (codePoint: number) => number {
  const table = required(tables, 'cmap', subject);
  const r = readerOf(table);
  if (table.length < 4) unreadable('cmap table is truncated', subject);
  const count = u16(r, 2);

  let best: { offset: number; score: number } | undefined;
  for (let i = 0; i < count; i++) {
    const at = 4 + i * 8;
    if (at + 8 > table.length) break;
    const platform = u16(r, at);
    const encoding = u16(r, at + 2);
    const offset = u32(r, at + 4);
    if (offset + 2 > table.length) continue;
    const unicode =
      (platform === 3 && (encoding === 10 || encoding === 1 || encoding === 0)) || platform === 0;
    if (!unicode) continue;
    const format = u16(readerOf(table.subarray(offset)), 0);
    const score = format === 12 ? 3 : format === 4 ? 2 : format === 6 || format === 0 ? 1 : 0;
    if (score > 0 && (best === undefined || score > best.score)) best = { offset, score };
  }
  if (best === undefined) unreadable('no Unicode cmap subtable', subject);

  const sub = table.subarray(best.offset);
  const s = readerOf(sub);
  const format = u16(s, 0);

  if (format === 4) {
    const segCount = u16(s, 6) / 2;
    const ends = 14;
    const starts = ends + segCount * 2 + 2;
    const deltas = starts + segCount * 2;
    const ranges = deltas + segCount * 2;
    return (codePoint: number): number => {
      if (codePoint > 0xffff) return 0;
      for (let seg = 0; seg < segCount; seg++) {
        if (u16(s, ends + seg * 2) < codePoint) continue;
        if (u16(s, starts + seg * 2) > codePoint) return 0;
        const rangeOffset = u16(s, ranges + seg * 2);
        if (rangeOffset === 0) return (codePoint + i16(s, deltas + seg * 2)) & 0xffff;
        const at = ranges + seg * 2 + rangeOffset + (codePoint - u16(s, starts + seg * 2)) * 2;
        if (at + 2 > sub.length) return 0;
        const glyph = u16(s, at);
        return glyph === 0 ? 0 : (glyph + i16(s, deltas + seg * 2)) & 0xffff;
      }
      return 0;
    };
  }

  if (format === 12) {
    const groups = u32(s, 12);
    return (codePoint: number): number => {
      let low = 0;
      let high = groups - 1;
      while (low <= high) {
        const mid = (low + high) >> 1;
        const at = 16 + mid * 12;
        const start = u32(s, at);
        const end = u32(s, at + 4);
        if (codePoint < start) high = mid - 1;
        else if (codePoint > end) low = mid + 1;
        else return u32(s, at + 8) + (codePoint - start);
      }
      return 0;
    };
  }

  if (format === 6) {
    const first = u16(s, 6);
    const entries = u16(s, 8);
    return (codePoint: number): number => {
      const at = codePoint - first;
      return at < 0 || at >= entries ? 0 : u16(s, 10 + at * 2);
    };
  }

  // Format 0: a 256-byte byte-to-glyph array.
  return (codePoint: number): number => (codePoint > 0xff ? 0 : u8(s, 6 + codePoint));
}

/* -------------------------------------------------------------------------- */
/* kerning                                                                    */
/* -------------------------------------------------------------------------- */

type Kerning = (left: number, right: number) => number;

const NO_KERNING: Kerning = () => 0;

function coverageIndex(r: Reader, at: number, glyph: number): number {
  const format = u16(r, at);
  if (format === 1) {
    const count = u16(r, at + 2);
    for (let i = 0; i < count; i++) if (u16(r, at + 4 + i * 2) === glyph) return i;
    return -1;
  }
  if (format !== 2) return -1;
  const ranges = u16(r, at + 2);
  for (let i = 0; i < ranges; i++) {
    const record = at + 4 + i * 6;
    if (glyph >= u16(r, record) && glyph <= u16(r, record + 2)) {
      return u16(r, record + 4) + (glyph - u16(r, record));
    }
  }
  return -1;
}

function classOf(r: Reader, at: number, glyph: number): number {
  const format = u16(r, at);
  if (format === 1) {
    const start = u16(r, at + 2);
    const count = u16(r, at + 4);
    const index = glyph - start;
    return index < 0 || index >= count ? 0 : u16(r, at + 6 + index * 2);
  }
  if (format !== 2) return 0;
  const ranges = u16(r, at + 2);
  for (let i = 0; i < ranges; i++) {
    const record = at + 4 + i * 6;
    if (glyph >= u16(r, record) && glyph <= u16(r, record + 2)) return u16(r, record + 4);
  }
  return 0;
}

/** The byte width of a GPOS value record, which is one 16-bit field per set bit. */
function valueSize(format: number): number {
  let bits = 0;
  for (let i = 0; i < 16; i++) if ((format & (1 << i)) !== 0) bits += 1;
  return bits * 2;
}

/** `XAdvance` is bit 2, and it is the only field a horizontal advance reads. */
const X_ADVANCE = 0x0004;

/** `LookupType` 2, pair adjustment: the only positioning that moves an advance. */
const PAIR_POS = 2;

/** `LookupType` 9, Extension Positioning, whose subtable names the real type. */
const EXTENSION_POS = 9;

function pairPosLookup(r: Reader, at: number): Kerning | undefined {
  const format = u16(r, at);
  const valueFormat1 = u16(r, at + 4);
  const valueFormat2 = u16(r, at + 6);
  if ((valueFormat1 & X_ADVANCE) === 0) return undefined;
  const size1 = valueSize(valueFormat1);
  const size2 = valueSize(valueFormat2);
  const coverage = at + u16(r, at + 2);

  if (format === 1) {
    const setCount = u16(r, at + 8);
    return (left: number, right: number): number => {
      const index = coverageIndex(r, coverage, left);
      if (index < 0 || index >= setCount) return 0;
      const set = at + u16(r, at + 10 + index * 2);
      const pairs = u16(r, set);
      for (let i = 0; i < pairs; i++) {
        const record = set + 2 + i * (2 + size1 + size2);
        if (u16(r, record) === right) return i16(r, record + 2);
      }
      return 0;
    };
  }

  if (format !== 2) return undefined;
  const classDef1 = at + u16(r, at + 8);
  const classDef2 = at + u16(r, at + 10);
  const class1Count = u16(r, at + 12);
  const class2Count = u16(r, at + 14);
  return (left: number, right: number): number => {
    if (coverageIndex(r, coverage, left) < 0) return 0;
    const c1 = classOf(r, classDef1, left);
    const c2 = classOf(r, classDef2, right);
    if (c1 >= class1Count || c2 >= class2Count) return 0;
    const record = at + 16 + (c1 * class2Count + c2) * (size1 + size2);
    return i16(r, record);
  };
}

/**
 * The pair adjustments behind an `ExtensionPosFormat1` subtable.
 *
 * The extension names the real lookup type and a 32-bit offset from its own
 * start, and may not itself target another, so exactly one level is followed.
 */
function extensionPairPos(r: Reader, at: number): Kerning | undefined {
  const end = r.bytes.length;
  if (at + 8 > end || u16(r, at) !== 1 || u16(r, at + 2) !== PAIR_POS) return undefined;
  const target = at + u32(r, at + 4);
  // A PairPos header is sixteen bytes at its widest and a 32-bit offset can
  // name anything, so one that runs past the table is read as no kerning.
  return target + 16 > end ? undefined : pairPosLookup(r, target);
}

/**
 * The `kern` feature's pair adjustments, if the font has GPOS.
 *
 * Only the `kern` feature: a font's GPOS also carries mark attachment and
 * cursive positioning, and neither changes an advance. Its lookups are type 2,
 * or the type 9 Extension a font compiler wraps them in.
 */
function gposKerning(tables: Tables): Kerning | undefined {
  const table = tables.get('GPOS');
  if (table === undefined || table.length < 10) return undefined;
  const r = readerOf(table);
  const featureList = u16(r, 6);
  const lookupList = u16(r, 8);
  if (featureList >= table.length || lookupList >= table.length) return undefined;

  const wanted = new Set<number>();
  const featureCount = u16(r, featureList);
  for (let i = 0; i < featureCount; i++) {
    const record = featureList + 2 + i * 6;
    if (tagAt(r, record) !== 'kern') continue;
    const feature = featureList + u16(r, record + 4);
    const lookups = u16(r, feature + 2);
    for (let j = 0; j < lookups; j++) wanted.add(u16(r, feature + 4 + j * 2));
  }
  if (wanted.size === 0) return undefined;

  const found: Kerning[] = [];
  const lookupCount = u16(r, lookupList);
  for (const index of wanted) {
    if (index >= lookupCount) continue;
    const lookup = lookupList + u16(r, lookupList + 2 + index * 2);
    const type = u16(r, lookup);
    if (type !== PAIR_POS && type !== EXTENSION_POS) continue;
    const subtables = u16(r, lookup + 4);
    for (let i = 0; i < subtables; i++) {
      const at = lookup + u16(r, lookup + 6 + i * 2);
      const pairs = type === PAIR_POS ? pairPosLookup(r, at) : extensionPairPos(r, at);
      if (pairs !== undefined) found.push(pairs);
    }
  }
  if (found.length === 0) return undefined;
  return (left, right) => {
    for (const lookup of found) {
      const adjust = lookup(left, right);
      if (adjust !== 0) return adjust;
    }
    return 0;
  };
}

/** The pre-OpenType `kern` table, format 0, horizontal subtables only. */
function legacyKerning(tables: Tables): Kerning | undefined {
  const table = tables.get('kern');
  if (table === undefined || table.length < 4) return undefined;
  const r = readerOf(table);
  const subtables = u16(r, 2);
  const pairs = new Map<number, number>();
  let at = 4;
  for (let i = 0; i < subtables && at + 14 <= table.length; i++) {
    const length = u16(r, at + 2);
    const coverage = u16(r, at + 4);
    const horizontal = (coverage & 0x0001) !== 0;
    const format = (coverage >> 8) & 0xff;
    if (horizontal && format === 0) {
      const count = u16(r, at + 6);
      for (let j = 0; j < count; j++) {
        const record = at + 14 + j * 6;
        if (record + 6 > table.length) break;
        pairs.set((u16(r, record) << 16) | u16(r, record + 2), i16(r, record + 4));
      }
    }
    at += length === 0 ? 14 : length;
  }
  if (pairs.size === 0) return undefined;
  return (left, right) => pairs.get((left << 16) | right) ?? 0;
}

/* -------------------------------------------------------------------------- */
/* BASE                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * The `BASE` horizontal axis' `ideo` coordinate under `DFLT`, in font units.
 *
 * One script and one tag: T13 built a font whose `DFLT`, `latn` and `hani`
 * coordinates all differ and Chromium took `DFLT`, and a second whose `BASE`
 * names every script but `DFLT` and Chromium read none of them. A table this
 * cannot follow is answered with `undefined`, which is what HarfBuzz's
 * sanitiser leaves the browser holding as well.
 */
function ideographicOf(tables: Tables): number | undefined {
  const table = tables.get('BASE');
  if (table === undefined || table.length < 8) return undefined;
  const r = readerOf(table);
  /** An offset that is neither NULL nor short of the bytes it points at. */
  const within = (offset: number, need: number): number | undefined =>
    offset > 0 && offset + need <= table.length ? offset : undefined;

  const axis = within(u16(r, 4), 4);
  if (axis === undefined) return undefined;
  const tagList = within(axis + u16(r, axis), 2);
  const scriptList = within(axis + u16(r, axis + 2), 2);
  if (tagList === undefined || scriptList === undefined) return undefined;

  const tagCount = u16(r, tagList);
  if (tagList + 2 + tagCount * 4 > table.length) return undefined;
  let tag = -1;
  for (let i = 0; i < tagCount && tag < 0; i++) {
    if (tagAt(r, tagList + 2 + i * 4) === 'ideo') tag = i;
  }
  if (tag < 0) return undefined;

  const scriptCount = u16(r, scriptList);
  if (scriptList + 2 + scriptCount * 6 > table.length) return undefined;
  let script: number | undefined;
  for (let i = 0; i < scriptCount && script === undefined; i++) {
    const record = scriptList + 2 + i * 6;
    if (tagAt(r, record) === 'DFLT') script = within(scriptList + u16(r, record + 4), 6);
  }
  if (script === undefined) return undefined;

  const values = within(script + u16(r, script), 4);
  if (values === undefined) return undefined;
  if (tag >= u16(r, values + 2) || values + 4 + tag * 2 + 2 > table.length) return undefined;
  const coord = within(values + u16(r, values + 4 + tag * 2), 4);
  if (coord === undefined) return undefined;
  // Formats 2 and 3 put the coordinate in the same field as format 1 and add a
  // device or glyph hint after it, so one read serves all three.
  const format = u16(r, coord);
  return format >= 1 && format <= 3 ? -i16(r, coord + 2) : undefined;
}

/* -------------------------------------------------------------------------- */
/* the face                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * The vertical metrics a browser reports for the face.
 *
 * Bit 7 wins on DirectWrite and FreeType, which otherwise answer `usWin` and
 * `hhea`; CoreText answers `hhea` either way, 248/248. ADR 0045, ADR 0053.
 */
function metricsOf(tables: Tables, subject: string, backend: FontBackend): FaceMetrics {
  const head = readerOf(required(tables, 'head', subject));
  const unitsPerEm = u16(head, 18);
  if (unitsPerEm <= 0) unreadable('head.unitsPerEm is zero', subject);

  const ideographic = ideographicOf(tables);
  const hhea = readerOf(required(tables, 'hhea', subject));
  const hheaPair: MetricPair = { ascent: i16(hhea, 4), descent: -i16(hhea, 6) };

  // 78 bytes is `OS/2` version 0, which is every field read below.
  const os2 = tables.get('OS/2');
  if (os2 === undefined || os2.length < 78) {
    // No OS/2 at all is a bare CJK or bitmap face; hhea is then the only source
    // there is, and the browser has nothing else to read either.
    return {
      unitsPerEm,
      ...hheaPair,
      source: 'hhea',
      backend,
      ideographic,
      candidates: { hhea: hheaPair, usWin: undefined, sTypo: undefined, useTypoMetrics: false },
    };
  }
  const r = readerOf(os2);
  const usWin: MetricPair = { ascent: u16(r, 74), descent: u16(r, 76) };
  const sTypo: MetricPair = { ascent: i16(r, 68), descent: -i16(r, 70) };
  const useTypoMetrics = (u16(r, 62) & USE_TYPO_METRICS) !== 0;
  const source =
    backend === 'coretext'
      ? 'hhea'
      : useTypoMetrics
        ? 'sTypo'
        : backend === 'freetype'
          ? 'hhea'
          : 'usWin';
  const chosen = source === 'sTypo' ? sTypo : source === 'hhea' ? hheaPair : usWin;
  return {
    unitsPerEm,
    ...chosen,
    source,
    backend,
    ideographic,
    candidates: { hhea: hheaPair, usWin, sTypo, useTypoMetrics },
  };
}

function advancesOf(tables: Tables, subject: string): (glyph: number) => number {
  const hhea = readerOf(required(tables, 'hhea', subject));
  const metrics = u16(hhea, 34);
  const hmtx = required(tables, 'hmtx', subject);
  const r = readerOf(hmtx);
  if (metrics === 0) unreadable('hhea.numberOfHMetrics is zero', subject);
  return (glyph: number): number => {
    const at = Math.min(glyph, metrics - 1) * 4;
    return at + 2 <= hmtx.length ? u16(r, at) : 0;
  };
}

const BOLD_STYLE = /bold|black|heavy|semibold|extrabold|demibold/i;
const ITALIC_STYLE = /italic|oblique/i;

/** Read one font out of an already-located table directory. */
export function faceOf(tables: Tables, subject: string, backend: FontBackend): Face {
  const names = namesOf(tables, subject);
  const family = names.get(1);
  if (family === undefined || family === '') unreadable('no family name', subject);
  const subfamily = names.get(2) ?? 'Regular';
  const cmap = cmapOf(tables, subject);
  const advance = advancesOf(tables, subject);

  // GPOS first, and only then the legacy table: T13 built a font whose two
  // disagreed and Chromium took the GPOS value.
  const kerning = gposKerning(tables) ?? legacyKerning(tables) ?? NO_KERNING;

  // `head.macStyle` is the fallback for a face whose subfamily is a design name
  // like `Condensed Light` rather than one of the four standard strings.
  const head = readerOf(required(tables, 'head', subject));
  const macStyle = u16(head, 44);

  return {
    family,
    subfamily,
    typographicFamily: names.get(16),
    metrics: metricsOf(tables, subject, backend),
    bold: BOLD_STYLE.test(subfamily) || (macStyle & 0x01) !== 0,
    italic: ITALIC_STYLE.test(subfamily) || (macStyle & 0x02) !== 0,
    advanceOf(codePoint: number): number | undefined {
      const glyph = cmap(codePoint);
      return glyph === 0 ? undefined : advance(glyph);
    },
    kernBetween(left: number, right: number): number {
      const a = cmap(left);
      const b = cmap(right);
      return a === 0 || b === 0 ? 0 : kerning(a, b);
    },
  };
}

/** Every face in a file, which is one unless the file is a collection. */
export function facesIn(bytes: Uint8Array, subject: string, backend: FontBackend): readonly Face[] {
  return fontsIn(bytes, subject).map((tables) => faceOf(tables, subject, backend));
}

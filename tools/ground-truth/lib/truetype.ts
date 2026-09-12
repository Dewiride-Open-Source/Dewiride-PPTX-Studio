/**
 * A TrueType font with the metrics the caller asks for.
 *
 * Real faces agree with themselves - Arial's `hhea.ascender` and its
 * `OS/2.usWinAscent` are both 1854 - so no real font can say which of them a
 * browser read. These fonts disagree on purpose, which is what makes the
 * question decidable in one measurement. Experiment T13, ADR 0042.
 */

/** `A`..`H` and space, which is every glyph any probe string here needs. */
const LETTERS = 'ABCDEFGH';

/**
 * `PUSHB[0] 0; POP`: a font program, so FreeType hints natively instead of
 * autohinting the advances (`FT_Load_Glyph`, ftobjs.c).
 */
const FONT_PROGRAM = new Uint8Array([0xb0, 0x00, 0x21]);

export interface Glyph {
  readonly advance: number;
  /** A filled rectangle in font units, or `undefined` for a blank glyph. */
  readonly box: { x0: number; y0: number; x1: number; y1: number } | undefined;
}

/** A kerning adjustment, in font units, applied between two characters. */
export interface KernPair {
  readonly left: string;
  readonly right: string;
  readonly adjust: number;
}

/** `BASE` baseline coordinates for one script, in font units. */
export type Baselines = Partial<Record<'hang' | 'icfb' | 'ideo' | 'romn', number>>;

/** The `BASE` script tags a probe may carry, in the order the table sorts them. */
export type BaseScriptTag = 'DFLT' | 'hani' | 'kana' | 'latn';

export interface FontSpec {
  readonly familyName: string;
  readonly unitsPerEm: number;
  /** `hhea.ascender` / `hhea.descender`. Descender is negative. */
  readonly hheaAscender: number;
  readonly hheaDescender: number;
  readonly hheaLineGap?: number;
  /** `OS/2.usWinAscent` / `usWinDescent`, both positive by definition. */
  readonly winAscent: number;
  readonly winDescent: number;
  /** `OS/2.sTypoAscender` / `sTypoDescender`. Descender is negative. */
  readonly typoAscender: number;
  readonly typoDescender: number;
  readonly typoLineGap?: number;
  /** `OS/2.fsSelection` bit 7, USE_TYPO_METRICS. */
  readonly useTypoMetrics?: boolean;
  /** Advance for every letter glyph, in font units. */
  readonly advance: number;
  readonly spaceAdvance: number;
  /** Pairs written into a legacy `kern` table, format 0. */
  readonly kern?: readonly KernPair[];
  /** Pairs written into a `GPOS` PairPos, format 1. */
  readonly gpos?: readonly KernPair[];
  /** Reach the PairPos through a type 9 Extension lookup, as compilers emit. */
  readonly gposExtension?: boolean;
  /**
   * A code point above the BMP, added in a second cmap subtable of format 12.
   *
   * Format 4 cannot hold one, so this is what separates a reader that picks the
   * best subtable from one that takes the first it recognises.
   */
  readonly astral?: number;
  /** A horizontal `BASE` axis, one coordinate set per script it names. */
  readonly base?: Readonly<Partial<Record<BaseScriptTag, Baselines>>>;
  /** A BMP ideograph mapped to the letter glyph, so a run in this font is Han. */
  readonly ideograph?: number;
}

export const BASE_SPEC: FontSpec = {
  familyName: 'PptxStudio Metrics',
  unitsPerEm: 1000,
  hheaAscender: 800,
  hheaDescender: -200,
  winAscent: 800,
  winDescent: 200,
  typoAscender: 800,
  typoDescender: -200,
  advance: 600,
  spaceAdvance: 400,
};

function glyphsOf(spec: FontSpec): Glyph[] {
  const out: Glyph[] = [
    { advance: spec.advance, box: { x0: 60, y0: 0, x1: 540, y1: 760 } },
    { advance: spec.spaceAdvance, box: undefined },
  ];
  for (let i = 0; i < LETTERS.length; i++) {
    out.push({ advance: spec.advance, box: { x0: 80, y0: 0, x1: 520, y1: 100 + i * 100 } });
  }
  return out;
}

/** The glyph id a character is mapped to, matching `cmapTable` below. */
export function glyphIdOf(char: string): number {
  if (char === ' ') return 1;
  const at = LETTERS.indexOf(char);
  if (at < 0) throw new Error(`no glyph for ${JSON.stringify(char)}`);
  return 2 + at;
}

class Writer {
  private readonly bytes: number[] = [];
  u8(v: number): this {
    this.bytes.push(v & 0xff);
    return this;
  }
  u16(v: number): this {
    return this.u8(v >> 8).u8(v);
  }
  i16(v: number): this {
    return this.u16(v < 0 ? v + 0x10000 : v);
  }
  u32(v: number): this {
    return this.u16((v >>> 16) & 0xffff).u16(v & 0xffff);
  }
  tag(t: string): this {
    for (let i = 0; i < 4; i++) this.u8(t.charCodeAt(i));
    return this;
  }
  utf16be(text: string): this {
    for (let i = 0; i < text.length; i++) this.u16(text.charCodeAt(i));
    return this;
  }
  zeros(n: number): this {
    for (let i = 0; i < n; i++) this.u8(0);
    return this;
  }
  get length(): number {
    return this.bytes.length;
  }
  done(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
}

function pad4(bytes: Uint8Array): Uint8Array {
  const over = bytes.length % 4;
  if (over === 0) return bytes;
  const out = new Uint8Array(bytes.length + (4 - over));
  out.set(bytes);
  return out;
}

function checksum(bytes: Uint8Array): number {
  const padded = pad4(bytes);
  let sum = 0;
  for (let i = 0; i < padded.length; i += 4) {
    sum =
      (sum +
        (((padded[i] ?? 0) << 24) |
          ((padded[i + 1] ?? 0) << 16) |
          ((padded[i + 2] ?? 0) << 8) |
          (padded[i + 3] ?? 0))) >>>
      0;
  }
  return sum >>> 0;
}

function glyfAndLoca(list: readonly Glyph[]): { glyf: Uint8Array; loca: Uint8Array } {
  const parts: Uint8Array[] = [];
  const offsets: number[] = [];
  let at = 0;
  for (const glyph of list) {
    offsets.push(at);
    if (glyph.box === undefined) continue;
    const { x0, y0, x1, y1 } = glyph.box;
    const w = new Writer();
    w.i16(1).i16(x0).i16(y0).i16(x1).i16(y1);
    w.u16(3); // endPtsOfContours: one contour, four points
    w.u16(0); // instructionLength
    // Four on-curve points, each an absolute x then y delta from the last.
    for (let i = 0; i < 4; i++) w.u8(0x01);
    const xs = [x0, x1, x1, x0];
    const ys = [y0, y0, y1, y1];
    let px = 0;
    for (const x of xs) {
      w.i16(x - px);
      px = x;
    }
    let py = 0;
    for (const y of ys) {
      w.i16(y - py);
      py = y;
    }
    const bytes = pad4(w.done());
    parts.push(bytes);
    at += bytes.length;
  }
  offsets.push(at);
  const glyf = new Uint8Array(at);
  let cursor = 0;
  for (const part of parts) {
    glyf.set(part, cursor);
    cursor += part.length;
  }
  const loca = new Writer();
  for (const offset of offsets) loca.u32(offset);
  return { glyf, loca: loca.done() };
}

/** The highest code point a font maps in the BMP, which `OS/2` also declares. */
function lastCharOf(spec: FontSpec): number {
  return spec.ideograph ?? 0x40 + LETTERS.length;
}

/** A format 4 subtable over space, `A`..`H` and the ideograph if there is one. */
function cmapFormat4(ideograph: number | undefined): Uint8Array {
  if (ideograph !== undefined && (ideograph <= 0x40 + LETTERS.length || ideograph >= 0xffff)) {
    throw new Error(`an ideograph at ${String(ideograph)} would not sort into the cmap`);
  }
  const ideo = ideograph === undefined ? [] : [ideograph];
  const starts = [0x20, 0x41, ...ideo, 0xffff];
  const ends = [0x20, 0x40 + LETTERS.length, ...ideo, 0xffff];
  // idDelta maps a code point to its glyph id: 0x20 -> 1, 0x41 -> 2.
  const deltas = [1 - 0x20, 2 - 0x41, ...ideo.map((c) => 2 - c), 1];
  const segments = starts.length;

  const sub = new Writer();
  sub.u16(4);
  sub.u16(16 + segments * 8);
  sub.u16(0); // language
  sub.u16(segments * 2);
  const searchRange = 2 * 2 ** Math.floor(Math.log2(segments));
  sub.u16(searchRange);
  sub.u16(Math.log2(searchRange / 2));
  sub.u16(segments * 2 - searchRange);
  for (const end of ends) sub.u16(end);
  sub.u16(0); // reservedPad
  for (const start of starts) sub.u16(start);
  for (const delta of deltas) sub.i16(delta);
  for (let i = 0; i < segments; i++) sub.u16(0); // idRangeOffset
  return sub.done();
}

/**
 * A format 12 subtable over the same characters plus one above the BMP.
 *
 * The astral code point is the whole point: format 4 cannot express it, so a
 * reader that prefers the BMP subtable answers "no glyph" and one that prefers
 * this answers with a width.
 */
function cmapFormat12(astral: number, ideograph: number | undefined): Uint8Array {
  const ideo: [number, number, number][] =
    ideograph === undefined ? [] : [[ideograph, ideograph, 2]];
  const groups: readonly [number, number, number][] = [
    [0x20, 0x20, 1],
    [0x41, 0x40 + LETTERS.length, 2],
    ...ideo,
    [astral, astral, 2],
  ];
  const sub = new Writer();
  sub.u16(12).u16(0);
  sub.u32(16 + groups.length * 12);
  sub.u32(0); // language
  sub.u32(groups.length);
  for (const [start, end, glyph] of groups) sub.u32(start).u32(end).u32(glyph);
  return sub.done();
}

/** The cmap, with the format 4 subtable first so a reader must choose on merit. */
function cmapTable(spec: FontSpec): Uint8Array {
  const { astral, ideograph } = spec;
  const subtables = [
    { platform: 3, encoding: 1, bytes: cmapFormat4(ideograph) },
    ...(astral === undefined
      ? []
      : [{ platform: 3, encoding: 10, bytes: cmapFormat12(astral, ideograph) }]),
  ];
  const headerLength = 4 + subtables.length * 8;
  const out = new Writer();
  out.u16(0).u16(subtables.length);
  let at = headerLength;
  for (const sub of subtables) {
    out.u16(sub.platform).u16(sub.encoding).u32(at);
    at += sub.bytes.length;
  }
  const header = out.done();
  const table = new Uint8Array(at);
  table.set(header);
  let cursor = headerLength;
  for (const sub of subtables) {
    table.set(sub.bytes, cursor);
    cursor += sub.bytes.length;
  }
  return table;
}

const NAME_STRINGS = (family: string): readonly [number, string][] => [
  [1, family],
  [2, 'Regular'],
  [3, `${family} 2026`],
  [4, family],
  [6, family.replace(/\s+/g, '')],
];

function nameTable(family: string): Uint8Array {
  const records = NAME_STRINGS(family);
  const storage = new Writer();
  const placed = records.map(([id, text]) => {
    const offset = storage.length;
    storage.utf16be(text);
    return { id, offset, length: text.length * 2 };
  });
  const out = new Writer();
  out
    .u16(0)
    .u16(placed.length)
    .u16(6 + placed.length * 12);
  for (const record of placed) {
    out.u16(3).u16(1).u16(0x0409).u16(record.id).u16(record.length).u16(record.offset);
  }
  const header = out.done();
  const strings = storage.done();
  const table = new Uint8Array(header.length + strings.length);
  table.set(header);
  table.set(strings, header.length);
  return table;
}

/** A `kern` table, format 0: the pre-OpenType way to say the same thing. */
function kernTable(pairs: readonly KernPair[]): Uint8Array {
  const sorted = [...pairs].sort(
    (a, b) => glyphIdOf(a.left) - glyphIdOf(b.left) || glyphIdOf(a.right) - glyphIdOf(b.right),
  );
  const sub = new Writer();
  sub
    .u16(0)
    .u16(14 + sorted.length * 6)
    .u16(1); // version, length, coverage: horizontal
  const searchRange = 6 * 2 ** Math.floor(Math.log2(sorted.length));
  sub.u16(sorted.length);
  sub.u16(searchRange);
  sub.u16(Math.log2(searchRange / 6));
  sub.u16(sorted.length * 6 - searchRange);
  for (const pair of sorted) {
    sub.u16(glyphIdOf(pair.left)).u16(glyphIdOf(pair.right)).i16(pair.adjust);
  }
  const subtable = sub.done();
  const out = new Writer();
  out.u16(0).u16(1);
  const header = out.done();
  const table = new Uint8Array(header.length + subtable.length);
  table.set(header);
  table.set(subtable, header.length);
  return table;
}

/**
 * A `GPOS` with one PairPos format 1 subtable under `kern`, latn/DFLT.
 *
 * Format 1 rather than 2 because it lists pairs literally, so what the fixture
 * says the font contains is what the bytes contain. Under `extension` the same
 * subtable is reached through an `ExtensionPosFormat1` of type 9.
 */
function gposTable(pairs: readonly KernPair[], extension: boolean): Uint8Array {
  const byLeft = new Map<number, KernPair[]>();
  for (const pair of pairs) {
    const left = glyphIdOf(pair.left);
    byLeft.set(left, [...(byLeft.get(left) ?? []), pair]);
  }
  const lefts = [...byLeft.keys()].sort((a, b) => a - b);

  // PairPos: header, then one offset per set, then the sets themselves.
  const headerLength = 10 + lefts.length * 2;
  const setBytes: Uint8Array[] = [];
  const setOffsets: number[] = [];
  let at = headerLength;
  for (const left of lefts) {
    const set = (byLeft.get(left) ?? []).sort((a, b) => glyphIdOf(a.right) - glyphIdOf(b.right));
    const w = new Writer();
    w.u16(set.length);
    for (const pair of set) w.u16(glyphIdOf(pair.right)).i16(pair.adjust);
    setOffsets.push(at);
    const bytes = w.done();
    setBytes.push(bytes);
    at += bytes.length;
  }
  const coverageAt = at;

  const pairPos = new Writer();
  pairPos.u16(1); // posFormat
  pairPos.u16(coverageAt);
  pairPos.u16(0x0004); // valueFormat1: XAdvance
  pairPos.u16(0x0000); // valueFormat2: nothing
  pairPos.u16(lefts.length);
  for (const offset of setOffsets) pairPos.u16(offset);
  for (const bytes of setBytes) for (const byte of bytes) pairPos.u8(byte);
  pairPos.u16(1).u16(lefts.length); // coverage format 1
  for (const left of lefts) pairPos.u16(left);
  const lookupSubtable = pairPos.done();

  const out = new Writer();
  out.u32(0x00010000);
  out.u16(10).u16(30).u16(44); // scriptList, featureList, lookupList

  // ScriptList: one script, one default language system, one feature.
  out.u16(1).tag('DFLT').u16(8);
  out.u16(4).u16(0); // Script: defaultLangSys at 4, no named systems
  out.u16(0).u16(0xffff).u16(1).u16(0); // LangSys: required none, one feature, index 0

  // FeatureList: one `kern` feature pointing at lookup 0.
  out.u16(1).tag('kern').u16(8);
  out.u16(0).u16(1).u16(0);

  // LookupList: one lookup, over one subtable eight bytes past its own start.
  out.u16(1).u16(4);
  out
    .u16(extension ? 9 : 2)
    .u16(0)
    .u16(1)
    .u16(8);
  // ExtensionPosFormat1: the real type, and a 32-bit offset from here.
  if (extension) out.u16(1).u16(2).u32(8);
  for (const byte of lookupSubtable) out.u8(byte);
  return out.done();
}

/** The baseline tags a `BASE` axis may carry, in the alphabetical order it requires. */
const BASE_TAGS = ['hang', 'icfb', 'ideo', 'romn'] as const;

/** The script tags a `BASE` axis may name, in the order the table sorts them. */
const BASE_SCRIPT_TAGS: readonly BaseScriptTag[] = ['DFLT', 'hani', 'kana', 'latn'];

/**
 * A horizontal `BASE` axis, one `BaseValues` per distinct coordinate set.
 *
 * The default baseline is `ideo` under an East Asian script and `romn`
 * elsewhere, and two scripts whose coordinates are equal share one `BaseScript`.
 */
function baseTable(base: Readonly<Partial<Record<BaseScriptTag, Baselines>>>): Uint8Array {
  const named = BASE_SCRIPT_TAGS.filter((tag) => base[tag] !== undefined);
  const first = base[named[0] ?? 'DFLT'];
  if (first === undefined) throw new Error('a BASE axis must name at least one script');
  const tags = BASE_TAGS.filter((tag) => first[tag] !== undefined);
  for (const script of named) {
    const own = BASE_TAGS.filter((tag) => base[script]?.[tag] !== undefined);
    if (own.join() !== tags.join()) {
      throw new Error('a BASE axis must give every script the same baseline tags');
    }
  }
  const romn = Math.max(tags.indexOf('romn'), 0);

  const sets = new Map<string, { defaultIndex: number; coords: readonly number[] }>();
  const scripts = named.map((tag) => {
    const set = base[tag] ?? {};
    const coords = tags.map((baseline) => set[baseline] ?? 0);
    const ideo = tags.indexOf('ideo');
    const eastAsian = tag === 'hani' || tag === 'kana';
    const defaultIndex = eastAsian && ideo >= 0 ? ideo : romn;
    const key = `${String(defaultIndex)}:${coords.join(',')}`;
    if (!sets.has(key)) sets.set(key, { defaultIndex, coords });
    return { tag, key };
  });
  const keys = [...sets.keys()];

  const axisAt = 8;
  const tagListAt = axisAt + 4;
  const scriptListAt = tagListAt + 2 + tags.length * 4;
  const scriptAt = scriptListAt + 2 + scripts.length * 6;
  const valuesAt = scriptAt + keys.length * 6;
  const valuesLength = 4 + tags.length * 2;
  const coordsAt = valuesAt + keys.length * valuesLength;

  const out = new Writer();
  out.u16(1).u16(0).u16(axisAt).u16(0);
  out.u16(tagListAt - axisAt).u16(scriptListAt - axisAt);
  out.u16(tags.length);
  for (const tag of tags) out.tag(tag);
  out.u16(scripts.length);
  for (const script of scripts) {
    out.tag(script.tag).u16(scriptAt + keys.indexOf(script.key) * 6 - scriptListAt);
  }
  for (let i = 0; i < keys.length; i++) {
    const own = scriptAt + i * 6;
    out
      .u16(valuesAt + i * valuesLength - own)
      .u16(0)
      .u16(0);
  }
  for (let i = 0; i < keys.length; i++) {
    const own = valuesAt + i * valuesLength;
    const set = sets.get(keys[i] ?? '');
    if (set === undefined) throw new Error('a BASE coordinate set went missing');
    out.u16(set.defaultIndex).u16(tags.length);
    for (let j = 0; j < tags.length; j++) {
      out.u16(coordsAt + (i * tags.length + j) * 4 - own);
    }
  }
  for (const key of keys) {
    for (const coordinate of sets.get(key)?.coords ?? []) out.u16(1).i16(coordinate);
  }
  const bytes = out.done();
  if (bytes.length !== coordsAt + keys.length * tags.length * 4) {
    throw new Error(`the BASE table came out ${String(bytes.length)} bytes long`);
  }
  return bytes;
}

export interface BuiltFont {
  readonly bytes: Uint8Array;
  readonly spec: FontSpec;
}

/** The tables a glyf-outline font needs, in the order a font directory sorts. */
export function buildFont(overrides: Partial<FontSpec> = {}): BuiltFont {
  const spec: FontSpec = { ...BASE_SPEC, ...overrides };
  const list = glyphsOf(spec);
  const { glyf, loca } = glyfAndLoca(list);

  const head = new Writer();
  head.u32(0x00010000).u32(0x00010000).u32(0); // version, revision, checkSumAdjustment
  head.u32(0x5f0f3cf5).u16(0).u16(spec.unitsPerEm);
  head.zeros(16); // created, modified
  head.i16(0).i16(spec.hheaDescender).i16(spec.advance).i16(spec.hheaAscender); // bbox
  head.u16(0).u16(8).i16(2); // macStyle, lowestRecPPEM, fontDirectionHint
  head.i16(1).i16(0); // indexToLocFormat: long. glyphDataFormat

  const hhea = new Writer();
  hhea.u32(0x00010000);
  hhea
    .i16(spec.hheaAscender)
    .i16(spec.hheaDescender)
    .i16(spec.hheaLineGap ?? 0);
  hhea.u16(spec.advance);
  hhea.i16(0).i16(0).i16(spec.advance);
  hhea.i16(1).i16(0).i16(0);
  hhea.zeros(8);
  hhea.i16(0);
  hhea.u16(list.length); // numberOfHMetrics: a full metric for every glyph

  const hmtx = new Writer();
  for (const glyph of list) hmtx.u16(glyph.advance).i16(glyph.box?.x0 ?? 0);

  const maxp = new Writer();
  maxp.u32(0x00010000).u16(list.length);
  maxp.u16(4).u16(1).u16(0).u16(0).u16(2).u16(0).u16(0).u16(0).u16(0).u16(1); // ..maxStackElements
  maxp.u16(0).u16(0).u16(0).u16(0);

  const os2 = new Writer();
  os2.u16(4); // version 4: everything read here is present from version 0
  os2.i16(spec.advance).u16(400).u16(5).u16(0); // xAvgCharWidth, weight, width, fsType
  os2.i16(650).i16(700).i16(-150).i16(200).i16(400).i16(300).i16(100).i16(-50).i16(400).i16(300);
  os2.i16(0); // sFamilyClass
  os2.zeros(10); // panose
  os2.u32(1).u32(0).u32(0).u32(0); // ulUnicodeRange: Basic Latin
  os2.tag('PXST');
  os2.u16(spec.useTypoMetrics === true ? 0x00c0 : 0x0040); // fsSelection: REGULAR (+ USE_TYPO)
  os2.u16(0x20).u16(lastCharOf(spec));
  os2
    .i16(spec.typoAscender)
    .i16(spec.typoDescender)
    .i16(spec.typoLineGap ?? 0);
  os2.u16(spec.winAscent).u16(spec.winDescent);
  os2.u32(1).u32(0); // ulCodePageRange
  os2.i16(500).i16(700).i16(0).i16(0).u16(0); // sxHeight, sCapHeight, default/break char, maxContext

  const post = new Writer();
  post.u32(0x00030000).u32(0).i16(-100).i16(50).u32(0).zeros(16);

  const tables: { tag: string; bytes: Uint8Array }[] = [
    { tag: 'OS/2', bytes: os2.done() },
    { tag: 'cmap', bytes: cmapTable(spec) },
    { tag: 'fpgm', bytes: FONT_PROGRAM },
    { tag: 'glyf', bytes: glyf },
    { tag: 'head', bytes: head.done() },
    { tag: 'hhea', bytes: hhea.done() },
    { tag: 'hmtx', bytes: hmtx.done() },
    { tag: 'loca', bytes: loca },
    { tag: 'maxp', bytes: maxp.done() },
    { tag: 'name', bytes: nameTable(spec.familyName) },
    { tag: 'post', bytes: post.done() },
  ];
  if (spec.base !== undefined) tables.push({ tag: 'BASE', bytes: baseTable(spec.base) });
  if (spec.gpos !== undefined) {
    tables.push({ tag: 'GPOS', bytes: gposTable(spec.gpos, spec.gposExtension === true) });
  }
  if (spec.kern !== undefined) tables.push({ tag: 'kern', bytes: kernTable(spec.kern) });
  tables.sort((a, b) => (a.tag < b.tag ? -1 : 1));

  const count = tables.length;
  const searchRange = 16 * 2 ** Math.floor(Math.log2(count));
  const directory = new Writer();
  directory.u32(0x00010000).u16(count);
  directory
    .u16(searchRange)
    .u16(Math.log2(searchRange / 16))
    .u16(count * 16 - searchRange);

  let offset = 12 + count * 16;
  const placed = tables.map((table) => {
    const at = offset;
    offset += pad4(table.bytes).length;
    return { ...table, offset: at };
  });
  for (const table of placed) {
    directory.tag(table.tag).u32(checksum(table.bytes)).u32(table.offset).u32(table.bytes.length);
  }

  const font = new Uint8Array(offset);
  font.set(directory.done());
  for (const table of placed) font.set(pad4(table.bytes), table.offset);

  // checkSumAdjustment: 0xB1B0AFBA less the checksum of the whole file with the
  // field zeroed, which it is at this point.
  const headOffset = placed.find((t) => t.tag === 'head')?.offset ?? 0;
  const adjustment = (0xb1b0afba - checksum(font)) >>> 0;
  new DataView(font.buffer).setUint32(headOffset + 8, adjustment);
  return { bytes: font, spec };
}

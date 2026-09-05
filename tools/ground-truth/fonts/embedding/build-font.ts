/**
 * Build a synthetic TrueType font from nothing.
 *
 * Experiment B asks whether PowerPoint will accept an EOT we wrote. Wrapping a
 * font off this machine would answer a different question badly: the font is
 * already installed, so PowerPoint could render it perfectly while ignoring our
 * bytes entirely, and we would call that a pass.
 *
 * So the probe font is authored here. Two consequences, both wanted:
 *
 *   - Its family name exists nowhere on earth, so if the glyphs appear, they
 *     came out of the file.
 *   - Every glyph is a solid bar of a distinct height, so "ABCDEFGH" renders as
 *     a staircase. No fallback font can accidentally produce that, which makes
 *     the pass/fail decidable by comparing pixels rather than by squinting.
 *
 * It is also CC0 by construction, which is the difference between a fixture we
 * can commit and one we cannot.
 *
 * The tables are the minimum a glyf-outline font needs: head, hhea, hmtx, maxp,
 * name, OS/2, post, cmap, glyf, loca. No hinting, no kerning, no CFF.
 */

export const FAMILY_NAME = 'PptxStudio Probe';
export const UNITS_PER_EM = 1000;

/** `A`..`H`, plus space. Glyph 0 is `.notdef`. */
const LETTERS = 'ABCDEFGH';
const ADVANCE = 600;
const SPACE_ADVANCE = 400;

const ASCENDER = 800;
const DESCENDER = -200;

interface Glyph {
  readonly advance: number;
  /** A rectangle in font units, or `undefined` for an empty glyph. */
  readonly box: { x0: number; y0: number; x1: number; y1: number } | undefined;
}

function glyphs(): Glyph[] {
  const out: Glyph[] = [
    // .notdef - a solid block, distinct from every letter.
    { advance: ADVANCE, box: { x0: 60, y0: 0, x1: 540, y1: 760 } },
    // space
    { advance: SPACE_ADVANCE, box: undefined },
  ];
  // A staircase: each letter one step taller than the last.
  for (let i = 0; i < LETTERS.length; i++) {
    const height = 100 + i * 100; // 100..800
    out.push({ advance: ADVANCE, box: { x0: 80, y0: 0, x1: 520, y1: height } });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* binary writer                                                              */
/* -------------------------------------------------------------------------- */

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

/** Pad to a 4-byte boundary with zeros, as the table directory requires. */
function pad4(bytes: Uint8Array): Uint8Array {
  const extra = (4 - (bytes.length % 4)) % 4;
  if (extra === 0) return bytes;
  const out = new Uint8Array(bytes.length + extra);
  out.set(bytes);
  return out;
}

/** The sum of a table's big-endian uint32 words, modulo 2^32. */
function checksum(bytes: Uint8Array): number {
  const padded = pad4(bytes);
  let sum = 0;
  for (let i = 0; i < padded.length; i += 4) {
    sum =
      (sum +
        ((padded[i]! << 24) | (padded[i + 1]! << 16) | (padded[i + 2]! << 8) | padded[i + 3]!)) >>>
      0;
  }
  return sum >>> 0;
}

/* -------------------------------------------------------------------------- */
/* tables                                                                     */
/* -------------------------------------------------------------------------- */

function glyfAndLoca(list: readonly Glyph[]): { glyf: Uint8Array; loca: Uint8Array } {
  const parts: Uint8Array[] = [];
  const offsets: number[] = [0];
  let at = 0;

  for (const glyph of list) {
    if (glyph.box === undefined) {
      // An empty glyph is encoded as a zero-length range in loca. Not a glyph
      // record with no contours - a *zero-length* one.
      offsets.push(at);
      continue;
    }
    const { x0, y0, x1, y1 } = glyph.box;
    const w = new Writer();
    w.i16(1); // numberOfContours
    w.i16(x0).i16(y0).i16(x1).i16(y1);
    w.u16(3); // endPtsOfContours: one contour, last point index 3
    w.u16(0); // instructionLength

    // Four on-curve points, clockwise with y up. Flag 0x01 is ON_CURVE with
    // both coordinates stored as int16 deltas - the plainest encoding there is.
    for (let i = 0; i < 4; i++) w.u8(0x01);
    const xs = [x0, 0, x1 - x0, 0];
    const ys = [y0, y1 - y0, 0, y0 - y1];
    for (const dx of xs) w.i16(dx);
    for (const dy of ys) w.i16(dy);

    const record = pad4(w.done());
    parts.push(record);
    at += record.length;
    offsets.push(at);
  }

  const glyf = new Uint8Array(at);
  let cursor = 0;
  for (const part of parts) {
    glyf.set(part, cursor);
    cursor += part.length;
  }

  // Long format: offsets are byte offsets, not halves. One entry more than
  // there are glyphs.
  const loca = new Writer();
  for (const offset of offsets) loca.u32(offset);
  return { glyf, loca: loca.done() };
}

function cmapTable(): Uint8Array {
  // Three segments: space, the letters, and the mandatory 0xFFFF terminator.
  const segments = [
    { start: 0x20, end: 0x20, delta: 1 - 0x20 },
    { start: 0x41, end: 0x40 + LETTERS.length, delta: 2 - 0x41 },
    { start: 0xffff, end: 0xffff, delta: 1 },
  ];
  const segCount = segments.length;
  const segCountX2 = segCount * 2;
  const searchRange = 2 * Math.pow(2, Math.floor(Math.log2(segCount)));
  const entrySelector = Math.log2(searchRange / 2);
  const rangeShift = segCountX2 - searchRange;

  const sub = new Writer();
  const subLength = 16 + segCount * 8;
  sub.u16(4).u16(subLength).u16(0);
  sub.u16(segCountX2).u16(searchRange).u16(entrySelector).u16(rangeShift);
  for (const s of segments) sub.u16(s.end);
  sub.u16(0); // reservedPad
  for (const s of segments) sub.u16(s.start);
  for (const s of segments) sub.i16(s.delta);
  for (const _ of segments) sub.u16(0); // idRangeOffset: delta-only
  const subtable = sub.done();

  // Two encoding records pointing at the same subtable: (3,1) is what Windows
  // reads, (0,3) is what a Unicode-first reader looks for.
  const headerSize = 4 + 2 * 8;
  const w = new Writer();
  w.u16(0).u16(2);
  w.u16(0).u16(3).u32(headerSize);
  w.u16(3).u16(1).u32(headerSize);
  const header = w.done();

  const out = new Uint8Array(header.length + subtable.length);
  out.set(header);
  out.set(subtable, header.length);
  return out;
}

function nameTable(familyName: string): Uint8Array {
  const records: [number, string][] = [
    [0, 'Released under CC0 1.0. Authored for PPTX Studio ground-truth testing.'],
    [1, familyName],
    [2, 'Regular'],
    [3, `${familyName}; Version 1.000`],
    [4, familyName],
    [5, 'Version 1.000'],
    [6, familyName.replace(/ /g, '')],
  ];

  const strings = new Writer();
  const offsets: { id: number; offset: number; length: number }[] = [];
  for (const [id, text] of records) {
    const offset = strings.length;
    strings.utf16be(text);
    offsets.push({ id, offset, length: text.length * 2 });
  }
  const stringBytes = strings.done();

  const w = new Writer();
  w.u16(0); // format 0
  w.u16(offsets.length);
  w.u16(6 + offsets.length * 12); // stringOffset
  for (const r of offsets) {
    // Platform 3 (Windows), encoding 1 (Unicode BMP), language 0x0409 (en-US).
    // This is the record Windows reads, and name ID 1 out of it is what
    // PowerPoint matches a typeface string against.
    w.u16(3).u16(1).u16(0x0409).u16(r.id).u16(r.length).u16(r.offset);
  }
  const header = w.done();

  const out = new Uint8Array(header.length + stringBytes.length);
  out.set(header);
  out.set(stringBytes, header.length);
  return out;
}

export interface BuiltFont {
  readonly bytes: Uint8Array;
  readonly familyName: string;
  readonly fullName: string;
  readonly styleName: string;
  readonly versionName: string;
  readonly panose: Uint8Array;
  readonly fsType: number;
  readonly weight: number;
  readonly unicodeRange: [number, number, number, number];
  readonly codePageRange: [number, number];
  readonly checkSumAdjustment: number;
  /** The characters the cmap covers, for the test text. */
  readonly sample: string;
}

export function buildProbeFont(familyName: string = FAMILY_NAME): BuiltFont {
  const list = glyphs();
  const { glyf, loca } = glyfAndLoca(list);

  const maxX = Math.max(...list.map((g) => g.box?.x1 ?? 0));
  const minX = Math.min(...list.map((g) => g.box?.x0 ?? 0));
  const maxY = Math.max(...list.map((g) => g.box?.y1 ?? 0));

  // PANOSE: family 2 (Latin text), the rest unclassified. Copied into the EOT
  // header and into p:embeddedFont/@panose, so it has to be real.
  const panose = new Uint8Array([2, 0, 5, 9, 0, 0, 0, 0, 0, 0]);
  const fsType = 0; // Installable Embedding, honestly - we wrote this font.
  const weight = 400;
  // Unicode range bit 0 is Basic Latin. Code page bit 0 is Latin-1.
  const unicodeRange: [number, number, number, number] = [1, 0, 0, 0];
  const codePageRange: [number, number] = [1, 0];

  const head = new Writer();
  head.u32(0x00010000); // version
  head.u32(0x00010000); // fontRevision 1.0
  head.u32(0); // checkSumAdjustment - filled in at the very end
  head.u32(0x5f0f3cf5); // magicNumber
  head.u16(0x000b); // flags: baseline at y=0, lsb at x=0, instructions may depend on ppem
  head.u16(UNITS_PER_EM);
  head.zeros(8).zeros(8); // created, modified - zero, so the build is reproducible
  head.i16(minX).i16(DESCENDER).i16(maxX).i16(maxY);
  head.u16(0); // macStyle
  head.u16(8); // lowestRecPPEM
  head.i16(2); // fontDirectionHint
  head.i16(1); // indexToLocFormat: long
  head.i16(0); // glyphDataFormat

  const hhea = new Writer();
  hhea.u32(0x00010000);
  hhea.i16(ASCENDER).i16(DESCENDER).i16(200);
  hhea.u16(ADVANCE);
  hhea.i16(minX).i16(0).i16(maxX);
  hhea.i16(1).i16(0).i16(0);
  hhea.zeros(8);
  hhea.i16(0);
  hhea.u16(list.length); // numberOfHMetrics: a full metric for every glyph

  const hmtx = new Writer();
  for (const g of list) hmtx.u16(g.advance).i16(g.box?.x0 ?? 0);

  const maxp = new Writer();
  maxp.u32(0x00010000);
  maxp.u16(list.length);
  maxp.u16(4); // maxPoints
  maxp.u16(1); // maxContours
  maxp.u16(0).u16(0);
  maxp.u16(2); // maxZones
  maxp.u16(0).u16(0).u16(0).u16(0).u16(0).u16(0).u16(0).u16(0);

  const os2 = new Writer();
  os2.u16(4); // version 4
  os2.i16(ADVANCE); // xAvgCharWidth
  os2.u16(weight).u16(5);
  os2.u16(fsType);
  os2.i16(650).i16(700).i16(0).i16(140); // subscript
  os2.i16(650).i16(700).i16(0).i16(480); // superscript
  os2.i16(50).i16(250); // strikeout
  os2.i16(0); // sFamilyClass
  for (const b of panose) os2.u8(b);
  for (const r of unicodeRange) os2.u32(r);
  os2.tag('PPTX'); // achVendID
  os2.u16(0x0040); // fsSelection: REGULAR
  os2.u16(0x20).u16(0x40 + LETTERS.length);
  os2.i16(ASCENDER).i16(DESCENDER).i16(200);
  os2.u16(ASCENDER).u16(-DESCENDER);
  os2.u32(codePageRange[0]).u32(codePageRange[1]);
  os2.i16(500).i16(700); // sxHeight, sCapHeight
  os2.u16(0).u16(0x20); // usDefaultChar, usBreakChar
  os2.u16(1); // usMaxContext

  const post = new Writer();
  post.u32(0x00030000); // version 3.0: no glyph names
  post.u32(0); // italicAngle
  post.i16(-100).i16(50);
  post.u32(1); // isFixedPitch
  post.u32(0).u32(0).u32(0).u32(0);

  const tables: [string, Uint8Array][] = [
    ['OS/2', os2.done()],
    ['cmap', cmapTable()],
    ['glyf', glyf],
    ['head', head.done()],
    ['hhea', hhea.done()],
    ['hmtx', hmtx.done()],
    ['loca', loca],
    ['maxp', maxp.done()],
    ['name', nameTable(familyName)],
    ['post', post.done()],
  ];
  // The directory must be sorted by tag. `OS/2` sorts before `cmap` because
  // uppercase is lower in ASCII, which is easy to get wrong by sorting
  // case-insensitively.
  tables.sort((a, b) => (a[0] < b[0] ? -1 : 1));

  const numTables = tables.length;
  const entrySelector = Math.floor(Math.log2(numTables));
  const searchRange = Math.pow(2, entrySelector) * 16;
  const rangeShift = numTables * 16 - searchRange;

  const directorySize = 12 + numTables * 16;
  let offset = directorySize;
  const placed = tables.map(([tag, bytes]) => {
    const entry = { tag, bytes, offset, length: bytes.length };
    offset += pad4(bytes).length;
    return entry;
  });

  const dir = new Writer();
  dir.u32(0x00010000).u16(numTables).u16(searchRange).u16(entrySelector).u16(rangeShift);
  for (const t of placed) {
    dir.tag(t.tag).u32(checksum(t.bytes)).u32(t.offset).u32(t.length);
  }

  const total = offset;
  const font = new Uint8Array(total);
  font.set(dir.done(), 0);
  for (const t of placed) font.set(pad4(t.bytes), t.offset);

  // head.checkSumAdjustment is the whole file's checksum subtracted from a
  // constant, computed with the field itself zero - which it currently is.
  const headEntry = placed.find((t) => t.tag === 'head')!;
  const adjustment = (0xb1b0afba - checksum(font)) >>> 0;
  const view = new DataView(font.buffer);
  view.setUint32(headEntry.offset + 8, adjustment);

  return {
    bytes: font,
    familyName,
    fullName: familyName,
    styleName: 'Regular',
    versionName: 'Version 1.000',
    panose,
    fsType,
    weight,
    unicodeRange,
    codePageRange,
    checkSumAdjustment: adjustment,
    sample: LETTERS,
  };
}

/**
 * Experiment T11, step 3 - fit the WordArt cell and write the fixture.
 *
 * ```
 * node tools/ground-truth/text/wordart-columns/analyse.ts <work-dir>
 * ```
 *
 * Every question is scored against its rivals and throws on a miss or a tie
 * rather than emitting a fixture it cannot fit. The one thing that is recorded
 * rather than ruled is the font box itself: see ADR 0040.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { repoPath } from '../../../repo/root.ts';
import { readBmp, type Bitmap } from '../../lib/bmp.ts';
import { readEmf, type EmfText } from '../../lib/emf.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse.ts <work-dir>');
const work: string = arg;

/* -------------------------------------------------------------------------- */
/* the readings                                                               */
/* -------------------------------------------------------------------------- */

interface PlacedProbe {
  readonly id: string;
  readonly family: string;
  readonly vert: string;
  readonly face: string;
  readonly sizePt: number;
  readonly text: string;
  readonly paragraphs: number;
  readonly wrap: string;
  readonly secondRun?: { readonly text: string; readonly sizePt: number } | undefined;
  readonly raster: boolean;
  readonly slide: number;
  readonly rect: { readonly x: number; readonly y: number; readonly w: number; readonly h: number };
}

interface Inputs {
  readonly slidePt: { readonly widthPt: number; readonly heightPt: number };
  readonly framePt: { readonly widthPt: number; readonly heightPt: number };
  readonly probes: readonly PlacedProbe[];
}

interface Readings {
  readonly opened: boolean;
  readonly slides: readonly {
    readonly slide: number;
    readonly emf: string | null;
    readonly bmp: string | null;
  }[];
  readonly shapes: readonly {
    readonly id: string;
    readonly boundWidth: number | null;
    readonly boundHeight: number | null;
    readonly boundLeft: number | null;
    readonly boundTop: number | null;
  }[];
}

interface BrowserFace {
  readonly face: string;
  readonly ascent: number;
  readonly descent: number;
  readonly advances: Readonly<Record<string, number>>;
}

const inputs = JSON.parse(readFileSync(join(work, 'column-inputs.json'), 'utf8')) as Inputs;
const readings = JSON.parse(readFileSync(join(work, 'column-readings.json'), 'utf8')) as Readings;
const browser = (
  JSON.parse(readFileSync(join(work, 'column-browser.json'), 'utf8')) as {
    readonly faces: readonly BrowserFace[];
  }
).faces;

if (!readings.opened) throw new Error('PowerPoint refused the probe deck');

const boxOf = new Map(browser.map((row) => [row.face, row.ascent + row.descent]));
const browserFace = new Map(browser.map((row) => [row.face, row]));

const streams = new Map<number, readonly EmfText[]>();
function textsOf(slide: number): readonly EmfText[] {
  const cached = streams.get(slide);
  if (cached !== undefined) return cached;
  const entry = readings.slides.find((row) => row.slide === slide);
  if (entry?.emf == null) throw new Error(`slide ${String(slide)} has no drawing stream`);
  const texts = readEmf(new Uint8Array(readFileSync(join(work, entry.emf)))).texts;
  streams.set(slide, texts);
  return texts;
}

const bitmaps = new Map<number, Bitmap>();
function bitmapOf(slide: number): Bitmap {
  const cached = bitmaps.get(slide);
  if (cached !== undefined) return cached;
  const entry = readings.slides.find((row) => row.slide === slide);
  if (entry?.bmp == null) throw new Error(`slide ${String(slide)} has no bitmap`);
  const bmp = readBmp(new Uint8Array(readFileSync(join(work, entry.bmp))));
  bitmaps.set(slide, bmp);
  return bmp;
}

/**
 * Logical units per point in the drawing stream, fitted rather than assumed.
 *
 * The scale family is a row of horizontal runs across the slide, each with its
 * pen on its own left edge, so the set is one line through the origin.
 */
function unitsPerPoint(): { units: number; residualPt: number } {
  let top = 0;
  let bottom = 0;
  const seen: { at: number; wanted: number }[] = [];
  for (const probe of inputs.probes) {
    if (probe.family !== 'scale') continue;
    for (const text of textsOf(probe.slide)) {
      if (Math.abs(text.transform[0]) < 0.5) continue;
      if (Math.abs(text.transform[4] / 8.3333333 - probe.rect.x) > 2) continue;
      top += text.transform[4] * probe.rect.x;
      bottom += probe.rect.x * probe.rect.x;
      seen.push({ at: text.transform[4], wanted: probe.rect.x });
    }
  }
  if (seen.length < 4) throw new Error('the scale family did not pin the stream scale');
  const units = top / bottom;
  const residualPt = Math.max(...seen.map((row) => Math.abs(row.at / units - row.wanted)));
  return { units, residualPt };
}

const { units: UNITS, residualPt: SCALE_RESIDUAL } = unitsPerPoint();
if (SCALE_RESIDUAL > 0.01) {
  throw new Error(`the stream scale is not a fit: ${SCALE_RESIDUAL.toFixed(4)}pt off`);
}

/** Every text record whose pen lands inside a probe's rectangle, in points. */
function pensOf(probe: PlacedProbe): readonly { x: number; y: number; text: EmfText }[] {
  const out: { x: number; y: number; text: EmfText }[] = [];
  for (const text of textsOf(probe.slide)) {
    const x = text.transform[4] / UNITS;
    const y = text.transform[5] / UNITS;
    const inside =
      x >= probe.rect.x - 2 &&
      x <= probe.rect.x + probe.rect.w + 2 &&
      y >= probe.rect.y - 2 &&
      y <= probe.rect.y + probe.rect.h + 2;
    if (inside) out.push({ x: x - probe.rect.x, y: y - probe.rect.y, text });
  }
  return out.sort((a, b) => a.text.index - b.text.index);
}

/* -------------------------------------------------------------------------- */
/* one row per probe                                                          */
/* -------------------------------------------------------------------------- */

interface Glyph {
  readonly glyph: string;
  readonly penXPt: number;
  readonly penYPt: number;
}

interface Row {
  readonly id: string;
  readonly family: string;
  readonly vert: string;
  readonly text: string;
  readonly paragraphs: number;
  readonly wrap: string;
  readonly secondRun: { readonly text: string; readonly sizePt: number } | null;
  readonly face: string;
  readonly sizePt: number;
  readonly columns: readonly (readonly Glyph[])[];
  readonly boundWidthPt: number;
  readonly boundHeightPt: number;
  readonly drewFaces: readonly string[];
  readonly turned: boolean;
}

/** The drawn characters, or the probe's own text where the record carries glyph ids. */
function charactersOf(pen: { text: EmfText }, fallback: string): readonly string[] {
  return pen.text.glyphIndices ? [...fallback] : [...pen.text.text];
}

/**
 * The glyphs of one probe, split into columns wherever the pen returns to the top.
 *
 * PowerPoint draws a stacked frame in reading order, one record per character,
 * so a pen that moves back up is the only column boundary there is.
 */
function rowOf(probe: PlacedProbe): Row {
  const pens = pensOf(probe);
  const columns: Glyph[][] = [];
  let current: Glyph[] = [];
  let previousY = -Infinity;
  let at = 0;
  for (const pen of pens) {
    const characters = charactersOf(pen, [...probe.text][at] ?? probe.text);
    at += characters.length;
    if (pen.y < previousY - 0.5 && current.length > 0) {
      columns.push(current);
      current = [];
    }
    previousY = pen.y;
    current.push({ glyph: characters.join(''), penXPt: pen.x, penYPt: pen.y });
  }
  if (current.length > 0) columns.push(current);

  const shape = readings.shapes.find((row) => row.id === probe.id);
  return {
    id: probe.id,
    family: probe.family,
    vert: probe.vert,
    text: probe.text,
    paragraphs: probe.paragraphs,
    wrap: probe.wrap,
    secondRun: probe.secondRun ?? null,
    face: probe.face,
    sizePt: probe.sizePt,
    columns,
    boundWidthPt: shape?.boundWidth ?? NaN,
    boundHeightPt: shape?.boundHeight ?? NaN,
    drewFaces: [...new Set(pens.map((pen) => pen.text.font?.face ?? '?'))],
    turned: pens.some(
      (pen) => Math.abs(pen.text.transform[1]) > 0.01 || (pen.text.font?.escapement ?? 0) !== 0,
    ),
  };
}

const rows = inputs.probes.filter((probe) => !probe.raster).map(rowOf);
const stacked = rows.filter((row) => row.vert !== 'horz');

/** The stacked probes whose column is set in one size, which is all but two. */
const oneSize = stacked.filter((row) => row.family !== 'sizes');

/** The pitch between consecutive cells in a column, and how far it varies. */
function pitchOf(row: Row): { pitchPt: number; spreadPt: number } | null {
  const gaps: number[] = [];
  for (const column of row.columns) {
    for (let at = 1; at < column.length; at += 1) {
      gaps.push(column[at]!.penYPt - column[at - 1]!.penYPt);
    }
  }
  if (gaps.length === 0) return null;
  const pitchPt = gaps.reduce((sum, gap) => sum + gap, 0) / gaps.length;
  return { pitchPt, spreadPt: Math.max(...gaps) - Math.min(...gaps) };
}

/* -------------------------------------------------------------------------- */
/* the questions                                                              */
/* -------------------------------------------------------------------------- */

/** The reading the column width is scored against: seven sixths of the font box. */
const COLUMN_OF_BOX = 7 / 6;

interface Candidate {
  readonly name: string;
  readonly hits: number;
}

interface Question {
  readonly key: string;
  readonly asks: string;
  readonly rows: number;
  readonly winner: string;
  readonly candidates: readonly Candidate[];
}

const questions: Question[] = [];

/** Scores every reading and throws unless one of them fits every row alone. */
function score<T>(
  key: string,
  asks: string,
  subjects: readonly T[],
  candidates: readonly { name: string; fits: (subject: T) => boolean }[],
): void {
  const scored = candidates
    .map((candidate) => ({ name: candidate.name, hits: subjects.filter(candidate.fits).length }))
    .sort((a, b) => b.hits - a.hits);
  const best = scored[0];
  if (best === undefined) throw new Error(`${key}: nothing to score`);
  if (best.hits < subjects.length) {
    throw new Error(
      `${key}: the best reading "${best.name}" fits ${String(best.hits)} of ` +
        `${String(subjects.length)}; the fixture is not written from a reading that does not fit`,
    );
  }
  if (scored.filter((row) => row.hits === best.hits).length > 1) {
    throw new Error(`${key}: ${best.name} is tied, so the probes do not separate the readings`);
  }
  questions.push({
    key,
    asks,
    rows: subjects.length,
    winner: best.name,
    candidates: scored,
  });
}

/* the cell is square: the pitch down a column is the width of the column */

const squareRows = oneSize
  .map((row) => ({ row, pitch: pitchOf(row) }))
  .filter((entry): entry is { row: Row; pitch: { pitchPt: number; spreadPt: number } } => {
    return entry.pitch !== null && entry.row.family !== 'break';
  })
  .map((entry) => ({
    face: entry.row.face,
    sizePt: entry.row.sizePt,
    pitchPt: entry.pitch.pitchPt,
    columnPt: entry.row.boundWidthPt / entry.row.columns.length,
    emPt: entry.row.sizePt,
  }));

/** Two logical units: a pen is written as an integer, so no reading is finer. */
const PEN_PT = 2 / UNITS;

score(
  'cell-is-square',
  'whether a stacked character advances by the width of its own column',
  squareRows,
  [
    {
      name: 'the cell is square: the pitch is the column width',
      fits: (r) => Math.abs(r.pitchPt - r.columnPt) <= PEN_PT,
    },
    { name: 'the pitch is one em', fits: (r) => Math.abs(r.pitchPt - r.emPt) <= PEN_PT },
    {
      name: 'the pitch is 1.2 of the em, as a horizontal line is',
      fits: (r) => Math.abs(r.pitchPt - 1.2 * r.emPt) <= PEN_PT,
    },
    {
      name: 'the pitch is 1.2 of the column',
      fits: (r) => Math.abs(r.pitchPt - 1.2 * r.columnPt) <= PEN_PT,
    },
  ],
);

/* the glyph is upright, and it is drawn in the plain face */

const uprightProbes = inputs.probes.filter((probe) => probe.raster);

/** The ink of one probe, in points from its own corner. */
function inkOf(probe: PlacedProbe): { widthPt: number; heightPt: number } {
  const bmp = bitmapOf(probe.slide);
  const perPoint = bmp.width / inputs.slidePt.widthPt;
  let x0 = Infinity;
  let x1 = -Infinity;
  let y0 = Infinity;
  let y1 = -Infinity;
  const right = Math.min(bmp.width, Math.ceil((probe.rect.x + probe.rect.w) * perPoint));
  const bottom = Math.min(bmp.height, Math.ceil((probe.rect.y + probe.rect.h) * perPoint));
  for (let y = Math.floor(probe.rect.y * perPoint); y < bottom; y += 1) {
    for (let x = Math.floor(probe.rect.x * perPoint); x < right; x += 1) {
      const [r, g, b] = bmp.pixel(x, y);
      if (r >= 200 && g >= 200 && b >= 200) continue;
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
  }
  if (x1 < x0) throw new Error(`${probe.id} drew no ink`);
  return { widthPt: (x1 - x0 + 1) / perPoint, heightPt: (y1 - y0 + 1) / perPoint };
}

const inks = uprightProbes.map((probe) => ({
  id: probe.id,
  vert: probe.vert,
  face: probe.face,
  text: probe.text,
  ...inkOf(probe),
}));

/** The horizontal reading of each glyph, which every other frame is scored against. */
const horizontalInk = new Map(
  inks.filter((row) => row.vert === 'horz').map((row) => [row.face, row]),
);

score('glyph-orientation', 'which way a character stacked in a WordArt column points', inks, [
  {
    name: 'a stacked character keeps the orientation it has in a horizontal frame',
    fits: (row) => {
      const flat = horizontalInk.get(row.face);
      if (flat === undefined) return false;
      return (
        Math.abs(row.widthPt - flat.widthPt) <= 0.5 && Math.abs(row.heightPt - flat.heightPt) <= 0.5
      );
    },
  },
  {
    name: 'a stacked character is turned a quarter, as a `vert` frame turns it',
    fits: (row) => {
      const flat = horizontalInk.get(row.face);
      if (flat === undefined) return false;
      return (
        Math.abs(row.widthPt - flat.heightPt) <= 0.5 && Math.abs(row.heightPt - flat.widthPt) <= 0.5
      );
    },
  },
]);

score(
  'vertical-face',
  'whether GDI is asked for the `@`-prefixed vertical variant of the face',
  stacked,
  [
    { name: 'the plain face, in every WordArt frame', fits: (row) => !row.turned },
    { name: 'the vertical variant, as an eaVert frame asks for', fits: (row) => row.turned },
  ],
);

/* which edge the columns stack from */

const columnRows = stacked.filter((row) => row.columns.length > 1);

score(
  'column-order',
  'which edge of the frame the first column of a WordArt frame sits against',
  columnRows,
  [
    {
      name: '`wordArtVert` stacks rightward from the left, `wordArtVertRtl` leftward from the right',
      fits: (row) => {
        const first = row.columns[0]?.[0]?.penXPt ?? NaN;
        const last = row.columns[row.columns.length - 1]?.[0]?.penXPt ?? NaN;
        return row.vert === 'wordArtVert' ? first < last : first > last;
      },
    },
    {
      name: 'both stack rightward from the left',
      fits: (row) =>
        (row.columns[0]?.[0]?.penXPt ?? NaN) < (row.columns.at(-1)?.[0]?.penXPt ?? NaN),
    },
    {
      name: 'both stack leftward from the right',
      fits: (row) =>
        (row.columns[0]?.[0]?.penXPt ?? NaN) > (row.columns.at(-1)?.[0]?.penXPt ?? NaN),
    },
  ],
);

/* the pitch belongs to the face, not to the character in the cell */

const mixedRows = stacked.filter((row) => row.family === 'mixed');

score(
  'pitch-per-face',
  'whether a Latin character in a CJK column takes a cell of its own size',
  mixedRows,
  [
    {
      name: 'every cell in a column is the same size, whatever character it holds',
      fits: (row) => (pitchOf(row)?.spreadPt ?? Infinity) <= 0.05 * row.sizePt,
    },
    {
      name: 'a cell is sized from the character it holds',
      fits: (row) => (pitchOf(row)?.spreadPt ?? 0) > 0.05 * row.sizePt,
    },
  ],
);

/** The size each glyph of a mixed-size probe was set in, in reading order. */
function sizesOf(row: Row): readonly number[] {
  const first = [...row.text].map(() => row.sizePt);
  const second = [...(row.secondRun?.text ?? '')].map(() => row.secondRun?.sizePt ?? row.sizePt);
  return [...first, ...second];
}

const sizeRows = stacked.filter((row) => row.family === 'sizes');

/** Two logical units, as everywhere else the pens are scored. */
function near(a: number, b: number): boolean {
  return Math.abs(a - b) <= 4 / UNITS;
}

score(
  'mixed-sizes',
  'whose cell a column takes when its runs are set in different sizes',
  sizeRows,
  [
    {
      name: 'each run stacks in a cell of its own size',
      fits: (row) => {
        const cell = (sizePt: number): number =>
          (boxOf.get(row.face) ?? NaN) * COLUMN_OF_BOX * sizePt;
        const sizes = sizesOf(row);
        const column = row.columns[0] ?? [];
        return column.every((glyph, at) => {
          const next = column[at + 1];
          if (next === undefined || sizes[at] !== sizes[at + 1]) return true;
          return near(next.penYPt - glyph.penYPt, cell(sizes[at] ?? 0));
        });
      },
    },
    {
      name: "the whole column stacks in the widest run's cell",
      fits: (row) => {
        const widest = (boxOf.get(row.face) ?? NaN) * COLUMN_OF_BOX * Math.max(...sizesOf(row));
        const column = row.columns[0] ?? [];
        return column.every((glyph, at) => {
          const next = column[at + 1];
          return next === undefined || near(next.penYPt - glyph.penYPt, widest);
        });
      },
    },
    {
      name: "the whole column stacks in the first run's cell",
      fits: (row) => {
        const first = (boxOf.get(row.face) ?? NaN) * COLUMN_OF_BOX * row.sizePt;
        const column = row.columns[0] ?? [];
        return column.every((glyph, at) => {
          const next = column[at + 1];
          return next === undefined || near(next.penYPt - glyph.penYPt, first);
        });
      },
    },
  ],
);

score('mixed-centring', 'which box a glyph of a mixed-size column is centred across', sizeRows, [
  {
    name: 'the column, which is the widest cell in it',
    fits: (row) => {
      const box = boxOf.get(row.face) ?? NaN;
      const columnPt = box * COLUMN_OF_BOX * Math.max(...sizesOf(row));
      const sizes = sizesOf(row);
      return (row.columns[0] ?? []).every((glyph, at) => {
        const sizePt = sizes[at] ?? 0;
        const advance = (browserFace.get(row.face)?.advances[glyph.glyph] ?? 0) * sizePt;
        return near(glyph.penXPt, (columnPt - advance) / 2);
      });
    },
  },
  {
    name: 'its own cell',
    fits: (row) => {
      const box = boxOf.get(row.face) ?? NaN;
      const sizes = sizesOf(row);
      return (row.columns[0] ?? []).every((glyph, at) => {
        const sizePt = sizes[at] ?? 0;
        const cellPt = box * COLUMN_OF_BOX * sizePt;
        const advance = (browserFace.get(row.face)?.advances[glyph.glyph] ?? 0) * sizePt;
        return near(glyph.penXPt, (cellPt - advance) / 2);
      });
    },
  },
]);

/* -------------------------------------------------------------------------- */
/* the cell, per face                                                         */
/* -------------------------------------------------------------------------- */

interface AxisRow {
  readonly face: string;
  readonly readings: number;
  readonly columnEm: number;
  readonly browserBoxEm: number;
  readonly predictedEm: number;
  readonly columnError: number;
  readonly baselineFromBottomEm: number;
  readonly browserDescent: number;
  readonly baselineError: number;
  readonly centringError: number;
}

const axes: AxisRow[] = [];
for (const face of [...new Set(oneSize.map((row) => row.face))]) {
  const mine = oneSize.filter((row) => row.face === face && row.family !== 'break');
  const columnEms = mine.map((row) => row.boundWidthPt / row.columns.length / row.sizePt);
  const columnEm = columnEms.reduce((sum, value) => sum + value, 0) / columnEms.length;
  const box = boxOf.get(face) ?? NaN;
  const metrics = browserFace.get(face);

  // The baseline of the first cell, measured up from that cell's bottom edge.
  const fromBottom = mine.map((row) => {
    const first = row.columns[0]?.[0];
    if (first === undefined) return NaN;
    return columnEm - first.penYPt / row.sizePt;
  });
  const baselineFromBottomEm =
    fromBottom.reduce((sum, value) => sum + value, 0) / fromBottom.length;

  // The glyph's own advance box, centred across the column. A right-to-left
  // frame draws its rightmost column first, so the edge comes from where the
  // column sits and not from the order it was drawn in.
  const centring = mine.flatMap((row) => {
    const cellPt = columnEm * row.sizePt;
    const blockLeftPt =
      row.vert === 'wordArtVert' ? 0 : inputs.framePt.widthPt - row.columns.length * cellPt;
    const meanX = (column: readonly Glyph[]): number =>
      column.reduce((sum, glyph) => sum + glyph.penXPt, 0) / column.length;
    const ranked = [...row.columns].sort((a, b) => meanX(a) - meanX(b));
    return ranked.flatMap((column, rank) =>
      column.map((glyph) => {
        const advance = metrics?.advances[glyph.glyph];
        if (advance === undefined) return NaN;
        const columnLeft = blockLeftPt + rank * cellPt;
        return (glyph.penXPt - columnLeft) / row.sizePt - (columnEm - advance) / 2;
      }),
    );
  });
  const clean = centring.filter((value) => Number.isFinite(value));

  axes.push({
    face,
    readings: mine.length,
    columnEm: round(columnEm),
    browserBoxEm: round(box),
    predictedEm: round(box * COLUMN_OF_BOX),
    columnError: round(columnEm - box * COLUMN_OF_BOX),
    baselineFromBottomEm: round(baselineFromBottomEm),
    browserDescent: round(metrics?.descent ?? NaN),
    baselineError: round(baselineFromBottomEm - (metrics?.descent ?? NaN)),
    centringError: round(Math.max(...clean.map(Math.abs))),
  });
}

/** The faces the browser's own font box is the box PowerPoint measured from. */
const FITS = 0.005;
const fitting = axes.filter((row) => Math.abs(row.columnError) <= FITS);
if (fitting.length < 6) {
  throw new Error(
    `seven sixths of the font box fits only ${String(fitting.length)} face(s); ` +
      'that is not a rule, it is a coincidence',
  );
}

/** A fortieth of the em: tighter than the two logical units the smallest probe rounds to. */
const READABLE = 0.025;

const worstBaseline = Math.max(...axes.map((row) => Math.abs(row.baselineError)));
if (worstBaseline > READABLE) {
  throw new Error(`the baseline is ${worstBaseline.toFixed(4)} em out; it is not the descent`);
}

const worstCentring = Math.max(...axes.map((row) => row.centringError));
if (worstCentring > READABLE) {
  throw new Error(
    `a glyph sits ${worstCentring.toFixed(4)} em from where centring its advance box puts it`,
  );
}

/**
 * How far PowerPoint's own pens stray from a stack of identical cells.
 *
 * It lays each cell out against a whole device pixel before it draws, so a
 * column's pitches are not all equal to the last thousandth. Nothing measured
 * against these pens can be held tighter than this.
 */
const penJitterEm = Math.max(
  ...oneSize.flatMap((row) =>
    row.columns.flatMap((column) => {
      if (column.length < 3) return [0];
      const first = column[0]?.penYPt ?? 0;
      const last = column[column.length - 1]?.penYPt ?? 0;
      const pitchPt = (last - first) / (column.length - 1);
      return column.map(
        (glyph, at) => Math.abs(glyph.penYPt - (first + at * pitchPt)) / row.sizePt,
      );
    }),
  ),
);

/**
 * How far PowerPoint's own cell strays from seven sixths of the font box.
 *
 * It rounds the cell before stacking, and at twelve points that rounding is
 * a hundredth of the em - Courier New stacks at 1.31 where the box asks for
 * 1.3218. At eighteen and twenty-eight it is a ten-thousandth.
 */
const cellRounding = oneSize
  .filter((row) => row.family !== 'break')
  .flatMap((row) => {
    const box = boxOf.get(row.face) ?? NaN;
    return row.columns
      .filter((column) => column.length >= 2)
      .map((column) => {
        const first = column[0]?.penYPt ?? 0;
        const last = column[column.length - 1]?.penYPt ?? 0;
        const pitchEm = (last - first) / (column.length - 1) / row.sizePt;
        return {
          face: row.face,
          sizePt: row.sizePt,
          offBy: Math.abs(pitchEm - box * COLUMN_OF_BOX),
        };
      });
  });

/** Only the faces whose own box is the box PowerPoint measured from. */
const fittingFaces = new Set(fitting.map((row) => row.face));

const cellRoundingEm = Math.max(
  ...cellRounding.filter((row) => fittingFaces.has(row.face)).map((row) => row.offBy),
);

/* -------------------------------------------------------------------------- */
/* how many cells fit in a column                                             */
/* -------------------------------------------------------------------------- */

const breakRows = stacked.filter((row) => row.family === 'break');
const capacity = breakRows.map((row) => {
  const cellPt = (row.boundWidthPt / row.columns.length) * 1;
  return {
    id: row.id,
    vert: row.vert,
    columns: row.columns.length,
    perColumn: row.columns.map((column) => column.length),
    cellPt: round(cellPt),
    framePt: inputs.framePt.heightPt,
    floorOfFrame: Math.floor(inputs.framePt.heightPt / cellPt),
  };
});
for (const row of capacity) {
  if (row.perColumn[0] !== row.floorOfFrame) {
    throw new Error(
      `${row.id}: the first column holds ${String(row.perColumn[0])} cells, ` +
        `not the ${String(row.floorOfFrame)} the frame has room for`,
    );
  }
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

function round(value: number): number {
  return Number.isFinite(value) ? Math.round(value * 1e4) / 1e4 : NaN;
}

const fixture = {
  $comment:
    'Ground truth for sub-phase 3.6: the cell a WordArt column stacks characters in. ' +
    'Every number was read out of PowerPoint through an exported EMF drawing stream, ' +
    'a bitmap at four pixels to the point, or TextRange2.BoundWidth. ' +
    'GENERATED by tools/ground-truth/text/wordart-columns/analyse.ts - then prettier --write ' +
    'this file before committing. See docs/adr/phase-3-text/0040-the-wordart-column.md.',
  experiment: 'T11',
  subPhase: '3.6',
  generatedBy: 'tools/ground-truth/text/wordart-columns/analyse.ts',
  unitsPerPoint: round(UNITS),
  scaleResidualPt: round(SCALE_RESIDUAL),
  columnOfFontBox: COLUMN_OF_BOX,
  penJitterEm: round(penJitterEm),
  cellRoundingEm: round(cellRoundingEm),
  cellRounding: [...new Set(cellRounding.map((row) => row.sizePt))]
    .sort((a, b) => a - b)
    .map((sizePt) => ({
      sizePt,
      worstEm: round(
        Math.max(
          ...cellRounding
            .filter((row) => row.sizePt === sizePt && fittingFaces.has(row.face))
            .map((row) => row.offBy),
        ),
      ),
    })),
  questions,
  axes,
  capacity,
  ink: inks.map((row) => ({
    id: row.id,
    vert: row.vert,
    face: row.face,
    text: row.text,
    widthPt: round(row.widthPt),
    heightPt: round(row.heightPt),
  })),
  browser: browser.map((row) => ({
    face: row.face,
    ascent: row.ascent,
    descent: row.descent,
    boxEm: round(row.ascent + row.descent),
    advances: row.advances,
  })),
  probes: rows.map((row) => ({
    id: row.id,
    family: row.family,
    vert: row.vert,
    text: row.text,
    paragraphs: row.paragraphs,
    wrap: row.wrap,
    secondRun: row.secondRun,
    face: row.face,
    sizePt: row.sizePt,
    boundWidthPt: round(row.boundWidthPt),
    boundHeightPt: round(row.boundHeightPt),
    columns: row.columns.map((column) =>
      column.map((glyph) => ({
        glyph: glyph.glyph,
        penXPt: round(glyph.penXPt),
        penYPt: round(glyph.penYPt),
      })),
    ),
    drewFaces: row.drewFaces,
    turned: row.turned,
  })),
};

const out = repoPath('corpus/ground-truth/wordart-columns.json');
writeFileSync(out, `${JSON.stringify(fixture, null, 2)}\n`);

process.stdout.write(`wordart-columns.json: ${String(rows.length)} probe(s)\n`);
for (const question of questions) {
  process.stdout.write(
    `  ${question.key.padEnd(18)} ${question.winner} (${String(question.rows)}/${String(question.rows)})\n`,
  );
}
for (const row of axes) {
  const verdict = Math.abs(row.columnError) <= FITS ? 'fits' : 'OFF';
  process.stdout.write(
    `  ${row.face.padEnd(18)} column ${row.columnEm.toFixed(4)} em vs ` +
      `${row.predictedEm.toFixed(4)} predicted  ${verdict} (${row.columnError.toFixed(4)})\n`,
  );
}

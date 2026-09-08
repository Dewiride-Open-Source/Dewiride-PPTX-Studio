/**
 * Experiment T10, step 3 - score the readings and write the fixture.
 *
 * ```
 * node tools/ground-truth/text/vertical-glyphs/analyse.ts <work-dir>
 * ```
 *
 * Each question is scored against every candidate reading, including the wrong
 * one anybody would write first, and a question that is meant to be exact throws
 * rather than emitting a fixture it cannot fit. The cross axis is the one
 * question with no exact reading; it is emitted as measurements and residuals
 * rather than as a rule, and ADR 0039 says so out loud.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp } from '../../lib/bmp.ts';
import { readEmf, type EmfText } from '../../lib/emf.ts';
import { repoPath } from '../../../repo/root.ts';
import { MIXED, type Probe } from './probes.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse.ts <work-dir>');
const work: string = arg;

interface PlacedProbe extends Probe {
  readonly slide: number;
  readonly rect: { readonly x: number; readonly y: number; readonly w: number; readonly h: number };
}

interface Inputs {
  readonly slidePt: { readonly widthPt: number; readonly heightPt: number };
  readonly probes: readonly PlacedProbe[];
}

interface Readings {
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly rasterWidth: number;
  readonly slides: readonly {
    readonly slide: number;
    readonly emf: string | null;
    readonly bmp: string | null;
  }[];
  readonly shapes: readonly {
    readonly id: string;
    readonly boundLeft: number | null;
    readonly boundTop: number | null;
    readonly boundWidth: number | null;
    readonly boundHeight: number | null;
    readonly orientation: number | null;
  }[];
}

const inputs = JSON.parse(readFileSync(join(work, 'glyph-inputs.json'), 'utf8')) as Inputs;
const readings = JSON.parse(readFileSync(join(work, 'glyph-readings.json'), 'utf8')) as Readings;

if (!readings.opened || readings.repaired === true) {
  throw new Error('PowerPoint refused or repaired the probe deck; there is nothing to measure');
}

/* -------------------------------------------------------------------------- */
/* the two instruments                                                        */
/* -------------------------------------------------------------------------- */

interface Box {
  readonly x0: number;
  readonly x1: number;
  readonly y0: number;
  readonly y1: number;
}

/** Anything darker than this is ink; the glyph is black on white. */
const INK_SUM = 600;

const bitmaps = new Map<number, ReturnType<typeof readBmp>>();
function bitmapOf(slide: number): ReturnType<typeof readBmp> {
  const cached = bitmaps.get(slide);
  if (cached !== undefined) return cached;
  const entry = readings.slides.find((row) => row.slide === slide);
  if (entry?.bmp == null) throw new Error(`slide ${String(slide)} has no bitmap`);
  const bmp = readBmp(new Uint8Array(readFileSync(join(work, entry.bmp))));
  bitmaps.set(slide, bmp);
  return bmp;
}

/** The ink inside one probe's rectangle, in points from the slide's corner. */
function inkOf(probe: PlacedProbe): Box | null {
  const bmp = bitmapOf(probe.slide);
  const perPoint = bmp.width / inputs.slidePt.widthPt;
  const left = Math.floor(probe.rect.x * perPoint);
  const right = Math.min(bmp.width, Math.ceil((probe.rect.x + probe.rect.w) * perPoint));
  const top = Math.floor(probe.rect.y * perPoint);
  const bottom = Math.min(bmp.height, Math.ceil((probe.rect.y + probe.rect.h) * perPoint));
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (let y = top; y < bottom; y += 1) {
    for (let x = left; x < right; x += 1) {
      const [r, g, b] = bmp.pixel(x, y);
      if (r + g + b >= INK_SUM) continue;
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
  }
  if (x0 === Infinity) return null;
  return {
    x0: x0 / perPoint - probe.rect.x,
    x1: (x1 + 1) / perPoint - probe.rect.x,
    y0: y0 / perPoint - probe.rect.y,
    y1: (y1 + 1) / perPoint - probe.rect.y,
  };
}

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

/**
 * Logical units per point in the drawing stream, fitted rather than assumed.
 *
 * Every `horz` probe puts its pen on its own left edge, so the whole set is one
 * line through the origin and its residual says whether the fit is a fit.
 */
function unitsPerPoint(): number {
  let top = 0;
  let bottom = 0;
  for (const probe of inputs.probes) {
    if (probe.vert !== 'horz') continue;
    for (const text of textsOf(probe.slide)) {
      if (Math.abs(text.transform[0]) < 0.5) continue;
      const wanted = probe.rect.x;
      if (Math.abs(text.transform[4] / 8.3333333 - wanted) > 2) continue;
      top += text.transform[4] * wanted;
      bottom += wanted * wanted;
    }
  }
  if (bottom === 0) throw new Error('no horizontal probe pinned the stream scale');
  return top / bottom;
}

const UNITS = unitsPerPoint();

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
/* the derived rows                                                           */
/* -------------------------------------------------------------------------- */

/** The stacking axis turns with the block; these two run the text downward. */
const UPRIGHT_TYPES = new Set(['eaVert', 'mongolianVert']);

interface Row {
  readonly id: string;
  readonly family: string;
  readonly vert: string;
  readonly face: string;
  readonly sizePt: number;
  readonly text: string;
  /** The face GDI was asked for, per drawn stretch. */
  readonly drew: readonly {
    readonly text: string;
    readonly face: string;
    readonly verticalFace: boolean;
    /** `ETO_GLYPH_INDEX`: the record holds glyph ids, so `text` is not characters. */
    readonly glyphIndices: boolean;
  }[];
  /** The pen of the first record, in points from the frame's corner. */
  readonly pen: readonly [number, number] | null;
  /** The per-character advances of the first record, in points. */
  readonly advances: readonly number[];
  readonly ink: Box | null;
  /** The stream's own world transform, which is where the quarter turn shows. */
  readonly turned: boolean;
}

const rows: Row[] = inputs.probes.map((probe) => {
  const pens = pensOf(probe);
  const first = pens[0];
  return {
    id: probe.id,
    family: probe.family,
    vert: probe.vert,
    face: probe.face,
    sizePt: probe.sizePt,
    text: probe.text,
    drew: pens.map((entry) => ({
      text: entry.text.text,
      face: entry.text.font?.face ?? '',
      verticalFace: entry.text.font?.verticalFace === true,
      glyphIndices: entry.text.glyphIndices,
    })),
    pen: first === undefined ? null : [round(first.x), round(first.y)],
    advances: (first?.text.dx ?? []).map((value) => round(value / UNITS)),
    ink: probe.raster ? roundBox(inkOf(probe)) : null,
    turned: first === undefined ? false : Math.abs(first.text.transform[0]) < 0.5,
  };
});

function round(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function roundBox(box: Box | null): Box | null {
  if (box === null) return null;
  return { x0: round(box.x0), x1: round(box.x1), y0: round(box.y0), y1: round(box.y1) };
}

const byId = new Map(rows.map((row) => [row.id, row]));

/* -------------------------------------------------------------------------- */
/* the questions                                                              */
/* -------------------------------------------------------------------------- */

interface Candidate {
  readonly name: string;
  readonly predict: (row: Row) => boolean | null;
}

interface Question {
  readonly key: string;
  readonly asks: string;
  readonly rows: readonly Row[];
  readonly truth: (row: Row) => boolean;
  readonly candidates: readonly Candidate[];
  /** A question that must fit perfectly, or the fixture is not written. */
  readonly exact: boolean;
}

/**
 * A glyph is upright when its ink keeps the shape it has drawn horizontally.
 *
 * Scored against the transpose rather than against a tolerance: `一` is 81.75 by
 * 5.25 one way round and 5.25 by 81.75 the other, so which of the two the ink is
 * nearer to is not a close call, and it survives the few points of difference
 * hinting makes between one orientation of an outline and the other.
 */
function isUpright(row: Row): boolean {
  const flat = byId.get(row.id.replace(row.vert, 'horz'));
  if (row.ink === null || flat?.ink == null) throw new Error(`${row.id} has no ink to compare`);
  const wide = row.ink.x1 - row.ink.x0;
  const tall = row.ink.y1 - row.ink.y0;
  const flatWide = flat.ink.x1 - flat.ink.x0;
  const flatTall = flat.ink.y1 - flat.ink.y0;
  const kept = Math.abs(wide - flatWide) + Math.abs(tall - flatTall);
  const turned = Math.abs(wide - flatTall) + Math.abs(tall - flatWide);
  if (kept === turned) throw new Error(`${row.id} is square, so its ink says nothing`);
  return kept < turned;
}

/**
 * The characters a stretch drew.
 *
 * A shaped run carries glyph ids rather than characters, so its slot has to come
 * from the probe's own text. Only the halfwidth katakana is shaped, and only
 * under the vertical face: the `@` variant has its own presentation forms.
 */
function charactersOf(row: Row, stretch: Row['drew'][number]): string {
  return stretch.glyphIndices ? row.text : stretch.text;
}

/** Whether a character belongs to the slot `a:ea` names, by the 3.5 tables. */
function isEastAsian(text: string): boolean {
  const code = text.codePointAt(0) ?? 0;
  const ranges: readonly (readonly [number, number])[] = [
    [0x1100, 0x11ff],
    [0x2e80, 0x2eff],
    [0x3000, 0x303f],
    [0x3040, 0x30ff],
    [0x3100, 0x312f],
    [0x3130, 0x318f],
    [0x3190, 0x319f],
    [0x31f0, 0x31ff],
    [0x3200, 0x33ff],
    [0x3400, 0x4dbf],
    [0x4e00, 0x9fff],
    [0xac00, 0xd7af],
    [0xf900, 0xfaff],
    [0xff00, 0xffef],
  ];
  return ranges.some(([from, to]) => code >= from && code <= to);
}

const geometry = rows.filter((row) => ['cell', 'glyph', 'latin'].includes(row.family));

/** The orientation question needs one glyph: two of them turn the block's shape. */
const single = geometry.filter((row) => [...row.text].length === 1);

const QUESTIONS: readonly Question[] = [
  {
    key: 'glyph-orientation',
    asks: 'which glyphs keep their own orientation when the block is turned',
    rows: single,
    truth: isUpright,
    exact: true,
    candidates: [
      {
        name: 'an East Asian glyph in an eaVert or mongolianVert frame',
        predict: (row) =>
          row.vert === 'horz' || (UPRIGHT_TYPES.has(row.vert) && isEastAsian(row.text)),
      },
      {
        name: 'every glyph in an eaVert or mongolianVert frame',
        predict: (row) => UPRIGHT_TYPES.has(row.vert) || row.vert === 'horz',
      },
      {
        name: 'nothing is ever upright but a horizontal frame',
        predict: (row) => row.vert === 'horz',
      },
      {
        name: 'every East Asian glyph, whatever the frame',
        predict: (row) => isEastAsian(row.text),
      },
    ],
  },
  {
    key: 'vertical-face',
    asks: 'when PowerPoint asks GDI for the @-prefixed vertical variant of a face',
    rows: rows.filter((row) => row.drew.length > 0),
    truth: (row) => row.drew.some((stretch) => stretch.verticalFace),
    exact: true,
    candidates: [
      {
        name: 'a stretch in the a:ea slot, in an eaVert or mongolianVert frame',
        predict: (row) =>
          UPRIGHT_TYPES.has(row.vert) &&
          row.drew.some((stretch) => isEastAsian(charactersOf(row, stretch))),
      },
      {
        name: 'any stretch in an eaVert or mongolianVert frame',
        predict: (row) => UPRIGHT_TYPES.has(row.vert),
      },
      { name: 'any turned frame', predict: (row) => row.vert !== 'horz' },
    ],
  },
  {
    key: 'face-per-stretch',
    asks: 'whether the face is chosen per script run rather than per frame',
    rows: rows.filter((row) => row.family === 'face'),
    truth: (row) =>
      row.drew.every(
        (stretch) =>
          stretch.verticalFace ===
          (UPRIGHT_TYPES.has(row.vert) && isEastAsian(charactersOf(row, stretch))),
      ),
    exact: true,
    candidates: [
      { name: 'per script run', predict: () => true },
      {
        name: 'per frame, so every stretch shares one face',
        predict: (row) =>
          row.drew.every((stretch) => stretch.verticalFace === row.drew[0]?.verticalFace),
      },
    ],
  },
];

interface Scored {
  readonly key: string;
  readonly asks: string;
  readonly rows: number;
  readonly winner: string;
  readonly candidates: readonly { readonly name: string; readonly hits: number }[];
}

const scored: Scored[] = [];
for (const question of QUESTIONS) {
  const marks = question.candidates.map((candidate) => ({
    name: candidate.name,
    hits: question.rows.filter((row) => candidate.predict(row) === question.truth(row)).length,
  }));
  const best = [...marks].sort((a, b) => b.hits - a.hits)[0];
  if (best === undefined) throw new Error(`${question.key} scored no candidate`);
  if (question.exact && best.hits !== question.rows.length) {
    throw new Error(
      `${question.key}: the best reading "${best.name}" fits ${String(best.hits)} of ` +
        `${String(question.rows.length)}; the fixture is not written from a reading that does not fit`,
    );
  }
  if (marks.filter((mark) => mark.hits === best.hits).length > 1) {
    throw new Error(
      `${question.key}: ${String(best.hits)} is a tie, so the probes do not separate the readings`,
    );
  }
  scored.push({
    key: question.key,
    asks: question.asks,
    rows: question.rows.length,
    winner: best.name,
    candidates: marks,
  });
}

/* -------------------------------------------------------------------------- */
/* the cell, and the one axis with no rule                                    */
/* -------------------------------------------------------------------------- */

/**
 * Where an upright glyph sits in its cell, as fractions of the em.
 *
 * `alongEm` is the alphabetic baseline from the cell's leading edge and
 * `descentEm` the em box below it; both are differences between the turned probe
 * and its own horizontal control, so the glyph's own outline cancels. `acrossEm`
 * is the glyph's x origin from the leading edge of the line box, and is the one
 * reading here that needs the block's placement as well.
 */
interface Cell {
  readonly id: string;
  readonly face: string;
  readonly sizePt: number;
  readonly vert: string;
  readonly alongEm: number;
  readonly descentEm: number;
  readonly acrossEm: number;
  /** The advance box centred across the line box, which is the candidate. */
  readonly centredEm: number;
}

/** The line box's leading edge on the slide: the edge its own top turns into. */
function leadingEdgePt(probe: PlacedProbe): number {
  // A quarter clockwise sends the laid-out top to the right, so the leading edge
  // is the right of the line box - the frame's own right where the block is
  // anchored there, and one line box in from the left where it is not.
  return probe.vert === 'mongolianVert' ? 1.2 * probe.sizePt : probe.rect.w;
}

const em = (value: number): number => Math.round(value * 10000) / 10000;

const cells: Cell[] = [];
for (const row of geometry) {
  if (!UPRIGHT_TYPES.has(row.vert) || !isEastAsian(row.text)) continue;
  const probe = inputs.probes.find((entry) => entry.id === row.id);
  const flat = byId.get(row.id.replace(row.vert, 'horz'));
  if (probe === undefined) continue;
  if (row.ink === null || row.pen === null || flat?.ink == null || flat.pen === null) continue;
  // Turned, the run goes down the slide and the cross axis runs back towards the
  // leading edge: pre_x = screen_y - penY and pre_y = penX - screen_x.
  const alongEm = (row.ink.y0 - row.pen[1] - (flat.ink.y0 - flat.pen[1])) / row.sizePt;
  const descentEm = (row.pen[0] - row.ink.x1 + (flat.ink.x1 - flat.pen[0])) / row.sizePt;
  const penAcross = leadingEdgePt(probe) - row.pen[0];
  const advanceEm = (row.advances[0] ?? row.sizePt) / row.sizePt;
  cells.push({
    id: row.id,
    face: row.face,
    sizePt: row.sizePt,
    vert: row.vert,
    alongEm: em(alongEm),
    descentEm: em(descentEm),
    acrossEm: em(penAcross / row.sizePt + descentEm),
    centredEm: em((1.2 + advanceEm) / 2),
  });
}

/* -------------------------------------------------------------------------- */
/* the two axes, and what a browser can reproduce of them                     */
/* -------------------------------------------------------------------------- */

interface BrowserFace {
  readonly face: string;
  readonly ascent: number;
  readonly descent: number;
  readonly ideographic: number;
}

/** Chromium's own metrics, when `measure-in-browser.ts` has been run. */
function browserFaces(): readonly BrowserFace[] {
  try {
    const raw = readFileSync(join(work, 'glyph-browser.json'), 'utf8');
    return (JSON.parse(raw) as { faces: readonly BrowserFace[] }).faces;
  } catch {
    return [];
  }
}

const browser = browserFaces();
const browserOf = new Map(browser.map((face) => [face.face, face]));

/**
 * Each axis of the upright cell, beside the readings a renderer could use.
 *
 * Neither candidate fits every face, and the fixture says so rather than
 * choosing: `1 - ideographicBaseline` is exact on the three faces whose cell is
 * an em and is 0.11 and 0.16 out on the two whose cell is not, and the advance
 * box centred across the line box is exact on MS Gothic and SimSun and 0.04 out
 * on Yu Gothic. ADR 0039 carries the table and the open question.
 */
const axes = [...new Set(cells.map((row) => row.face))].map((face) => {
  const rows = cells.filter((row) => row.face === face);
  const mean = (pick: (row: Cell) => number): number =>
    em(rows.reduce((sum, row) => sum + pick(row), 0) / rows.length);
  const metrics = browserOf.get(face);
  const alongEm = mean((row) => row.alongEm);
  const acrossEm = mean((row) => row.acrossEm);
  return {
    face,
    readings: rows.length,
    alongEm,
    descentEm: mean((row) => row.descentEm),
    cellEm: em(alongEm + mean((row) => row.descentEm)),
    acrossEm,
    centredEm: mean((row) => row.centredEm),
    acrossError: em(acrossEm - mean((row) => row.centredEm)),
    browserIdeographic: metrics === undefined ? null : em(metrics.ideographic),
    alongError: metrics === undefined ? null : em(alongEm - (1 - metrics.ideographic)),
  };
});

/** The line pitch, from two pens rather than from a bounding box. */
const pitches = rows
  .filter((row) => row.family === 'pitch' && row.drew.length >= 2)
  .map((row) => {
    const pens = pensOf(inputs.probes.find((probe) => probe.id === row.id) as PlacedProbe);
    const xs = pens.map((entry) => entry.x).sort((a, b) => b - a);
    return {
      face: row.face,
      sizePt: row.sizePt,
      pitchPt: round((xs[0] ?? 0) - (xs[1] ?? 0)),
      pitchEm: round(((xs[0] ?? 0) - (xs[1] ?? 0)) / row.sizePt),
    };
  });

/**
 * The cell pitch, from the two-glyph probes: the ink of a pair less the ink of
 * one, along whichever axis the text runs in.
 */
const advances = rows
  .filter((row) => row.family === 'advance' && row.ink !== null)
  .map((row) => {
    const one = byId.get(`cell-Yu Gothic-96-${row.vert}`);
    if (one?.ink == null || row.ink === null) throw new Error(`${row.id} has no single-glyph twin`);
    const along = row.vert === 'horz' ? 'x' : 'y';
    const pair = along === 'x' ? row.ink.x1 - row.ink.x0 : row.ink.y1 - row.ink.y0;
    const solo = along === 'x' ? one.ink.x1 - one.ink.x0 : one.ink.y1 - one.ink.y0;
    return { vert: row.vert, sizePt: row.sizePt, pitchPt: round(pair - solo) };
  });

const pitchSpread = advances.map((row) => row.pitchPt);
if (Math.max(...pitchSpread) - Math.min(...pitchSpread) > 0.5) {
  throw new Error(`the cell pitch is not the same in every direction: ${pitchSpread.join(', ')}`);
}

const badPitch = pitches.filter((row) => Math.abs(row.pitchEm - 1.2) > 0.005);
if (badPitch.length > 0) {
  throw new Error(
    `the line pitch is not 1.2 of the em on ${badPitch.map((row) => row.face).join(', ')}`,
  );
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const fixture = {
  $comment:
    'GENERATED by tools/ground-truth/text/vertical-glyphs/analyse.ts. Do not edit by hand. ' +
    'Experiment T10, sub-phase 3.6, ADR 0039. Run build-deck.ts, read.ps1, ' +
    'measure-in-browser.ts and analyse.ts to regenerate, then prettier --write this file ' +
    'before committing.',
  experiment: 'T10',
  subPhase: '3.6',
  generatedBy: 'tools/ground-truth/text/vertical-glyphs/analyse.ts',
  unitsPerPoint: round(UNITS),
  questions: scored,
  linePitch: { readingEm: 1.2, rows: pitches },
  cellPitch: { readingEm: 1, rows: advances },
  cells,
  axes,
  browser,
  mixed: MIXED,
  probes: rows,
};

const out = repoPath('corpus/ground-truth/vertical-glyphs.json');
writeFileSync(out, `${JSON.stringify(fixture, null, 2)}\n`);
process.stdout.write(
  `vertical-glyphs.json: ${String(rows.length)} probe(s), ${String(scored.length)} question(s), ` +
    `${String(cells.length)} cell reading(s), ${String(pitches.length)} pitch reading(s)\n`,
);
for (const question of scored) {
  process.stdout.write(
    `  ${question.key.padEnd(20)} ${String(question.rows).padStart(4)} rows  ${question.winner}\n`,
  );
}
process.stdout.write('  face                along    cell  across centred  along err across err\n');
for (const row of axes) {
  const show = (value: number | null): string =>
    (value === null ? '-' : value.toFixed(4)).padStart(9);
  process.stdout.write(
    `  ${row.face.padEnd(18)}${show(row.alongEm)}${show(row.cellEm)}${show(row.acrossEm)}` +
      `${show(row.centredEm)}${show(row.alongError)}${show(row.acrossError)}\n`,
  );
}

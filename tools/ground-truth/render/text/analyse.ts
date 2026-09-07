/**
 * Experiment T8, step 4 - score the candidate readings and write the fixture.
 *
 * ```
 * node tools/ground-truth/render/text/analyse.ts <work-dir> [--fixture <path>]
 * ```
 *
 * `choose` throws unless exactly one candidate fits every row of a question, so
 * a rule that is merely the best of a bad set cannot reach the fixture. The
 * candidates include, for each question, the reading a renderer written by hand
 * would have used - the CSS half-leading baseline, the counter-flip on both
 * axes, the run boundary that stops a kern - because knowing which of those is
 * wrong is most of the value.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { format, resolveConfig } from 'prettier';

import { repoPath } from '../../../repo/root.ts';
import { inkBox, readBmp, type Bitmap } from '../../lib/bmp.ts';
import { readEmf, type EmfText } from '../../lib/emf.ts';

import { FACES, SIZES, type Rect } from './probes.ts';

/* -------------------------------------------------------------------------- */
/* arguments and inputs                                                       */
/* -------------------------------------------------------------------------- */

const dirArg = process.argv[2];
if (dirArg === undefined) {
  throw new Error('usage: analyse.ts <dir> [--fixture corpus/ground-truth/text-rendering.json]');
}
const dir: string = resolve(dirArg);
const fixtureFlag = process.argv.indexOf('--fixture');
const fixturePath =
  fixtureFlag < 0
    ? repoPath('corpus/ground-truth/text-rendering.json')
    : resolve(process.argv[fixtureFlag + 1] ?? '');

interface ProbeIndex {
  readonly id: string;
  readonly question: string;
  readonly deck: string;
  readonly rect: Rect;
  readonly bitmap: boolean;
  readonly asks: string;
}

interface Bounds {
  readonly index: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly text: string | null;
}

interface ShapeReading {
  readonly name: string;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly rotation: number | null;
  readonly textLeft: number | null;
  readonly textTop: number | null;
  readonly text: string | null;
  readonly lines: readonly Bounds[];
  readonly runs: readonly (Bounds & { readonly font: Record<string, number | null> })[];
}

interface SlideReading {
  readonly slide: number;
  readonly probe: string;
  readonly shapes: readonly ShapeReading[];
  readonly emf: string | null;
  readonly bmp: string | null;
}

interface DeckReading {
  readonly deck: string;
  readonly question: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly slides: readonly SlideReading[];
}

const inputs = JSON.parse(readFileSync(join(dir, 'text-render-inputs.json'), 'utf8')) as {
  probes: readonly ProbeIndex[];
};
const readings = JSON.parse(readFileSync(join(dir, 'text-render-readings.json'), 'utf8')) as {
  decks: readonly DeckReading[];
};
const browser = JSON.parse(readFileSync(join(dir, 'text-render-browser.json'), 'utf8')) as {
  chromium: string;
  faces: Record<string, { ladder: Record<string, Record<string, number>> }>;
  decorations: Record<
    string,
    Record<'underline' | 'strike', { offsetEm: number; thicknessEm: number }>
  >;
};

for (const deck of readings.decks) {
  if (!deck.opened) throw new Error(`deck ${deck.deck} was refused; there is nothing to analyse`);
  if (deck.repaired === true)
    throw new Error(`deck ${deck.deck} was repaired; it measures a repair`);
}

const probeById = new Map(inputs.probes.map((probe) => [probe.id, probe]));

/* -------------------------------------------------------------------------- */
/* units                                                                      */
/* -------------------------------------------------------------------------- */

/** Every probe slide exports at 1920 x 1080 over 960 x 540 points. */
const PX_PER_POINT = 2;

/**
 * EMF logical units per point, checked rather than assumed.
 *
 * `lfHeight` is the em size in logical units, so 34 probes at four sizes are 34
 * chances for this to be something else.
 */
const UNITS_PER_POINT = 600 / 72;

/* -------------------------------------------------------------------------- */
/* one row per probe                                                          */
/* -------------------------------------------------------------------------- */

interface Row {
  readonly probe: ProbeIndex;
  readonly shape: ShapeReading;
  readonly texts: readonly EmfText[];
  readonly bitmap: Bitmap | null;
}

const rowsByQuestion = new Map<string, Row[]>();

for (const deck of readings.decks) {
  for (const slide of deck.slides) {
    const probe = probeById.get(slide.probe);
    if (probe === undefined) throw new Error(`reading names probe ${slide.probe}, which has none`);
    const shape = slide.shapes.find((s) => s.name === probe.id);
    if (shape === undefined) throw new Error(`slide for ${probe.id} holds no shape of that name`);
    if (slide.emf === null) throw new Error(`${probe.id} exported no EMF`);
    const emf = readEmf(new Uint8Array(readFileSync(join(dir, slide.emf))));
    const bitmap =
      slide.bmp === null ? null : readBmp(new Uint8Array(readFileSync(join(dir, slide.bmp))));
    const bucket = rowsByQuestion.get(probe.question) ?? [];
    bucket.push({ probe, shape, texts: emf.texts, bitmap });
    rowsByQuestion.set(probe.question, bucket);
  }
}

function rowsFor(question: string): readonly Row[] {
  const rows = rowsByQuestion.get(question);
  if (rows === undefined || rows.length === 0) throw new Error(`no rows for question ${question}`);
  return rows;
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Candidate<T> {
  readonly name: string;
  readonly says: string;
  readonly predict: (row: T) => unknown;
}

interface Scored {
  readonly question: string;
  readonly asks: string;
  readonly answer: string;
  readonly says: string;
  readonly scores: Record<string, string>;
  readonly total: number;
}

const scored: Scored[] = [];

/**
 * The one candidate that fits every row, or an error naming the alternatives.
 *
 * A question whose best candidate misses even one row has not been answered:
 * emitting the best of a bad set is how a guess becomes a fixture.
 */
function choose<T>(
  question: string,
  asks: string,
  rows: readonly T[],
  observe: (row: T) => unknown,
  candidates: readonly Candidate<T>[],
): Candidate<T> {
  const same = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);
  const scores: Record<string, string> = {};
  const winners: Candidate<T>[] = [];
  for (const candidate of candidates) {
    let hits = 0;
    for (const row of rows) if (same(candidate.predict(row), observe(row))) hits++;
    scores[candidate.name] = `${String(hits)}/${String(rows.length)}`;
    if (hits === rows.length) winners.push(candidate);
  }
  if (winners.length !== 1) {
    const table = Object.entries(scores)
      .map(([name, score]) => `    ${name.padEnd(28)} ${score}`)
      .join('\n');
    throw new Error(
      `question "${question}" has ${String(winners.length)} candidate(s) fitting all ` +
        `${String(rows.length)} row(s), and needs exactly one:\n${table}`,
    );
  }
  const winner = winners[0];
  if (winner === undefined) throw new Error('unreachable');
  scored.push({
    question,
    asks,
    answer: winner.name,
    says: winner.says,
    scores,
    total: rows.length,
  });
  return winner;
}

/** The spaces a line ends with, which an alignment does not count. */
const TRAILING_SPACES = / +$/;

/**
 * Record that every candidate misses at least one row, and none is the answer.
 *
 * Some questions are settled by a measurement rather than by a rule: the useful
 * result is that the two implementations a renderer would reach for are both
 * wrong, and the table of measured values is what it has to carry instead.
 */
function refute<T>(
  question: string,
  asks: string,
  says: string,
  rows: readonly T[],
  candidates: readonly Candidate<T>[],
): void {
  const scores: Record<string, string> = {};
  const survivors: string[] = [];
  for (const candidate of candidates) {
    let hits = 0;
    for (const row of rows) if (candidate.predict(row) === true) hits++;
    scores[candidate.name] = `${String(hits)}/${String(rows.length)}`;
    if (hits === rows.length) survivors.push(candidate.name);
  }
  if (survivors.length > 0) {
    throw new Error(
      `question "${question}" was written down as refuted, but ${survivors.join(', ')} fits every row`,
    );
  }
  scored.push({ question, asks, answer: 'none-of-these', says, scores, total: rows.length });
}

/** A length in points, rounded to the precision the instrument actually has. */
function pt(value: number, places = 4): number {
  const factor = 10 ** places;
  return Math.round(value * factor) / factor;
}

/** Whether two lengths agree to within one unit of the instrument. */
function near(a: number, b: number, tolerance: number): boolean {
  return Math.abs(a - b) <= tolerance;
}

/* -------------------------------------------------------------------------- */
/* the turn a text block is drawn at                                          */
/* -------------------------------------------------------------------------- */

/** The quadrant a text record's world transform turns through, or `null`. */
function quadrantTurn(text: EmfText): { turn: number; mirrored: boolean } | null {
  const [m11, m12, m21, m22] = text.transform;
  const mirrored = m11 * m22 - m12 * m21 < 0;
  const table: readonly (readonly [number, number, number, number, number])[] = [
    [1, 0, 0, 1, 0],
    [0, 1, -1, 0, 90],
    [-1, 0, 0, -1, 180],
    [0, -1, 1, 0, 270],
  ];
  for (const [a, b, c, d, turn] of table) {
    if (near(m11, a, 1e-3) && near(m12, b, 1e-3) && near(m21, c, 1e-3) && near(m22, d, 1e-3)) {
      return { turn, mirrored };
    }
  }
  return null;
}

/** The centre of a bitmap's ink, in slide points. */
function inkCentre(bitmap: Bitmap): { x: number; y: number } {
  const box = inkBox(bitmap);
  if (box === null) throw new Error('the slide drew no ink at all');
  return {
    x: (box.left + box.right + 1) / 2 / PX_PER_POINT,
    y: (box.top + box.bottom + 1) / 2 / PX_PER_POINT,
  };
}

function frameCentre(rect: Rect): { x: number; y: number } {
  return { x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 };
}

/* -------------------------------------------------------------------------- */
/* Q1 - flip                                                                   */
/* -------------------------------------------------------------------------- */

interface FlipRow {
  readonly row: Row;
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  /** The turn relative to the same rotation with no flip: 0 or 180. */
  readonly extraTurn: number;
  readonly mirrored: boolean;
}

function parseFlipId(id: string): { rot: number; flipH: boolean; flipV: boolean } {
  const match = /^flip-r(\d+)-(none|h|v|hv)$/.exec(id);
  if (match === null) throw new Error(`${id} is not a flip probe id`);
  const axes = match[2] ?? '';
  return {
    rot: Number(match[1]),
    flipH: axes.includes('h'),
    flipV: axes.includes('v'),
  };
}

const flipRows: FlipRow[] = (() => {
  const byId = new Map(rowsFor('flip').map((row) => [row.probe.id, row]));
  const out: FlipRow[] = [];
  for (const row of rowsFor('flip')) {
    const spec = parseFlipId(row.probe.id);
    const base = byId.get(`flip-r${String(spec.rot)}-none`);
    if (base === undefined) throw new Error(`no unflipped probe at ${String(spec.rot)} degrees`);
    const mine = row.texts[0];
    const theirs = base.texts[0];
    if (mine !== undefined && theirs !== undefined) {
      const a = quadrantTurn(mine);
      const b = quadrantTurn(theirs);
      if (a === null || b === null) throw new Error(`${row.probe.id} is not on a quadrant`);
      out.push({ ...spec, row, extraTurn: (a.turn - b.turn + 360) % 360, mirrored: a.mirrored });
      continue;
    }
    // Off a quadrant PowerPoint draws the glyphs as filled paths and the EMF
    // carries no text record, so the ink's own centre answers instead.
    if (row.bitmap === null || base.bitmap === null) {
      throw new Error(`${row.probe.id} has neither a text record nor a bitmap`);
    }
    const mineCentre = inkCentre(row.bitmap);
    const baseCentre = inkCentre(base.bitmap);
    const centre = frameCentre(row.probe.rect);
    const turned = { x: 2 * centre.x - baseCentre.x, y: 2 * centre.y - baseCentre.y };
    const isSame = near(mineCentre.x, baseCentre.x, 1) && near(mineCentre.y, baseCentre.y, 1);
    const isTurned = near(mineCentre.x, turned.x, 1) && near(mineCentre.y, turned.y, 1);
    if (isSame === isTurned) throw new Error(`${row.probe.id} is neither turned nor not`);
    out.push({ ...spec, row, extraTurn: isTurned ? 180 : 0, mirrored: false });
  }
  return out;
})();

const flipAnswer = choose<FlipRow>(
  'flip',
  'what a flipped shape does to its text',
  flipRows,
  (row) => ({ extraTurn: row.extraTurn, mirrored: row.mirrored }),
  [
    {
      name: 'flipH-ignored-flipV-turns',
      says: 'flipH does nothing to the text; flipV turns the whole block 180 degrees; neither mirrors',
      predict: (row) => ({ extraTurn: row.flipV ? 180 : 0, mirrored: false }),
    },
    {
      name: 'both-flips-ignored',
      says: 'the text is drawn as if the shape carried no flip at all',
      predict: () => ({ extraTurn: 0, mirrored: false }),
    },
    {
      name: 'both-flips-turn',
      says: 'each flip turns the block 180 degrees, so the two cancel',
      predict: (row) => ({
        extraTurn: ((row.flipH ? 180 : 0) + (row.flipV ? 180 : 0)) % 360,
        mirrored: false,
      }),
    },
    {
      name: 'flipV-ignored-flipH-turns',
      says: 'the mirror image of the answer: flipV does nothing and flipH turns',
      predict: (row) => ({ extraTurn: row.flipH ? 180 : 0, mirrored: false }),
    },
    {
      name: 'text-mirrors-with-shape',
      says: 'the glyphs are mirrored exactly as the outline is',
      predict: (row) => ({ extraTurn: 0, mirrored: row.flipH !== row.flipV }),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q2 - rotation                                                              */
/* -------------------------------------------------------------------------- */

interface RotRow {
  readonly row: Row;
  readonly rot: number;
  readonly upright: boolean;
  readonly bodyRot: number;
  readonly angle: number;
}

/**
 * The angle that carries the unturned ink to where this probe's ink is.
 *
 * The ink box's extents are the same at an angle and its negation, so the
 * *offset* from the frame centre is what has to be fitted: it is the one
 * observable off a quadrant that tells 30 degrees from 330.
 */
function angleFromInk(
  row: Row,
  unturnedOffset: { x: number; y: number },
  angles: readonly number[],
): number {
  if (row.bitmap === null) throw new Error(`${row.probe.id} has no bitmap`);
  const centre = frameCentre(row.probe.rect);
  const seen = inkCentre(row.bitmap);
  const offset = { x: seen.x - centre.x, y: seen.y - centre.y };
  let best: { angle: number; error: number } | null = null;
  for (const angle of angles) {
    const radians = (angle * Math.PI) / 180;
    const x = unturnedOffset.x * Math.cos(radians) - unturnedOffset.y * Math.sin(radians);
    const y = unturnedOffset.x * Math.sin(radians) + unturnedOffset.y * Math.cos(radians);
    const error = Math.hypot(x - offset.x, y - offset.y);
    if (best === null || error < best.error) best = { angle, error };
  }
  if (best === null) throw new Error('no angles offered');
  // A rotated run's ink box is not the rotated ink box, so the fit is a few
  // points loose; the nearest rival angle is 85pt away at this radius.
  if (best.error > 12) {
    throw new Error(`${row.probe.id}: no offered angle puts the ink within 12pt`);
  }
  return ((best.angle % 360) + 360) % 360;
}

const rotRows: RotRow[] = (() => {
  const rows = rowsFor('rotation');
  const flat = rows.find((row) => row.probe.id === 'rot-0-unset');
  if (flat?.bitmap == null) throw new Error('no unrotated rotation probe with a bitmap');
  const flatCentre = frameCentre(flat.probe.rect);
  const flatInk = inkCentre(flat.bitmap);
  const unturned = { x: flatInk.x - flatCentre.x, y: flatInk.y - flatCentre.y };
  const zero = flat.texts[0];
  if (zero === undefined) throw new Error('the unrotated probe drew no text record');
  const zeroTurn = quadrantTurn(zero)?.turn ?? 0;

  const candidateAngles = [0, 30, 45, 60, 75, 90, 105, 180, 270, 330];
  return rows.map((row) => {
    const spec = /^rot-(\d+)-(unset|up|down)$/.exec(row.probe.id);
    const body = /^rot-body-(neg)?(\d+)$/.exec(row.probe.id);
    const both = row.probe.id === 'rot-both-30-45';
    const rot = spec !== null ? Number(spec[1]) : both ? 30 : 0;
    const upright = spec?.[2] === 'up';
    const bodyRot = body !== null ? (body[1] === 'neg' ? -1 : 1) * Number(body[2]) : both ? 45 : 0;
    const text = row.texts[0];
    const angle =
      text === undefined
        ? angleFromInk(row, unturned, candidateAngles)
        : ((quadrantTurn(text)?.turn ?? 0) - zeroTurn + 360) % 360;
    return { row, rot, upright, bodyRot, angle };
  });
})();

const turn = (value: number): number => ((value % 360) + 360) % 360;

choose<RotRow>('rotation', 'the angle the glyphs are drawn at', rotRows, (row) => row.angle, [
  {
    name: 'upright-drops-the-shape-turn',
    says: 'the text turns by the shape rotation plus a:bodyPr/@rot, and @upright="1" drops the shape rotation',
    predict: (row) => turn((row.upright ? 0 : row.rot) + row.bodyRot),
  },
  {
    name: 'always-the-sum',
    says: 'the two rotations always add and @upright changes nothing',
    predict: (row) => turn(row.rot + row.bodyRot),
  },
  {
    name: 'shape-rotation-only',
    says: 'a:bodyPr/@rot does not turn the text',
    predict: (row) => turn(row.rot),
  },
  {
    name: 'body-rotation-only',
    says: 'the text follows a:bodyPr/@rot and not the shape',
    predict: (row) => turn(row.bodyRot),
  },
  {
    name: 'text-never-turns',
    says: 'the glyphs stay upright whatever the shape does',
    predict: () => 0,
  },
]);

/* -------------------------------------------------------------------------- */
/* Q2b - the box upright text is laid out in                                  */
/* -------------------------------------------------------------------------- */

/** Whether a quadrant-snapped rotation swaps which extent reaches which axis. */
function swapsExtents(degrees: number): boolean {
  const quadrant = Math.round(turn(degrees) / 90) % 4;
  return quadrant === 1 || quadrant === 3;
}

interface UprightRow {
  readonly row: Row;
  readonly rot: number;
  /** The left edge of the box the text was laid out in, in points. */
  readonly boxLeft: number;
}

const uprightRows: UprightRow[] = rotRows
  .filter((row) => row.upright)
  .map((row) => {
    const text = row.row.texts[0];
    if (text === undefined) throw new Error(`${row.row.probe.id} drew no text record`);
    return { row: row.row, rot: row.rot, boxLeft: pt(text.origin[0] / UNITS_PER_POINT) };
  });

choose<UprightRow>(
  'upright-box',
  'the rectangle a:bodyPr/@upright lays the text out in',
  uprightRows,
  (row) => row.boxLeft,
  [
    {
      name: 'extents-swap-on-a-quadrant',
      says: "the frame's own extents, swapped when the rotation snaps to 90 or 270, centred on the frame",
      predict: (row) => {
        const rect = row.row.probe.rect;
        const width = swapsExtents(row.rot) ? rect.h : rect.w;
        return pt(rect.x + rect.w / 2 - width / 2);
      },
    },
    {
      name: 'never-swapped',
      says: 'the frame as written, whatever the rotation',
      predict: (row) => pt(row.row.probe.rect.x),
    },
    {
      name: 'always-swapped',
      says: 'the extents swap at every angle',
      predict: (row) => {
        const rect = row.row.probe.rect;
        return pt(rect.x + rect.w / 2 - rect.h / 2);
      },
    },
    {
      name: 'rotated-bounding-box',
      says: 'the true axis-aligned box of the turned frame, at every angle',
      predict: (row) => {
        const rect = row.row.probe.rect;
        const radians = (row.rot * Math.PI) / 180;
        const width = rect.w * Math.abs(Math.cos(radians)) + rect.h * Math.abs(Math.sin(radians));
        return pt(rect.x + rect.w / 2 - width / 2);
      },
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q3 - the baseline inside the line box                                      */
/* -------------------------------------------------------------------------- */

/** `MEASURED_FACE_METRICS` from sub-phase 3.4, as the rival this question has. */
const AUTOFIT_B: Readonly<Record<string, number>> = {
  Arial: 0.227618991,
  'Times New Roman': 0.234424476,
  Verdana: 0.207355431,
  'Courier New': 0.300000058,
  Georgia: 0.231489897,
  Tahoma: 0.205318787,
};
const APPROXIMATE_B = AUTOFIT_B['Arial'] ?? 0;

/** The face's box, as a fraction of the em, from the least quantised reading. */
const LADDER_PX = 1000;

function faceBox(face: string): { ascent: number; descent: number; ink: number } {
  const entry = browser.faces[face];
  const at = entry?.ladder[String(LADDER_PX)];
  if (at === undefined) throw new Error(`no browser metrics for ${face}`);
  const ascent = (at['fontAscent'] ?? 0) / LADDER_PX;
  const descent = (at['fontDescent'] ?? 0) / LADDER_PX;
  const ink = (at['inkAscent'] ?? 0) / LADDER_PX;
  if (ascent <= 0 || descent <= 0) throw new Error(`${face} reported no font box`);
  return { ascent, descent, ink };
}

interface BaseRow {
  readonly row: Row;
  readonly face: string;
  /** The largest size on the line, which is what the box turns out to follow. */
  readonly size: number;
  readonly sizes: readonly number[];
  readonly lineHeight: number;
  /** How far the baseline sits below the top of the line box, in points. */
  readonly drop: number;
}

/** One EMF unit is 0.12pt, which is the whole precision this question has. */
const UNIT_PT = 1 / UNITS_PER_POINT;

const baseRows: BaseRow[] = rowsFor('baseline')
  .filter((row) => row.shape.lines.length === 1 && row.texts.length > 0)
  .map((row) => {
    const text = row.texts[0];
    const line = row.shape.lines[0];
    if (text === undefined || line?.top == null || line.height == null) {
      throw new Error(`${row.probe.id} is missing a line box`);
    }
    // Every run on a line shares one baseline, so a mixed line is one reading.
    for (const other of row.texts) {
      if (Math.abs(other.origin[1] - text.origin[1]) > 1) {
        throw new Error(`${row.probe.id} drew two baselines on one line`);
      }
    }
    const stated = row.shape.runs.map((run) => run.font['size'] ?? 0);
    const sizes = row.texts.map((each) => {
      // `lfHeight` is the em size in whole logical units, so it is the size a
      // run asked for to within half a unit - and that is 159 chances for the
      // unit to be something other than a six-hundredth of an inch.
      const height = -(each.font?.height ?? 0) / UNITS_PER_POINT;
      const nearest = stated.find((size) => Math.abs(size - height) <= 0.5 * UNIT_PT);
      if (nearest === undefined) {
        throw new Error(
          `${row.probe.id}: lfHeight ${String(each.font?.height)} is ${height.toFixed(3)}pt, which no run asked for`,
        );
      }
      return nearest;
    });
    return {
      row,
      face: text.font?.face ?? '',
      size: Math.max(...sizes),
      sizes,
      lineHeight: line.height,
      drop: pt(text.origin[1] / UNITS_PER_POINT - line.top),
    };
  });

choose<BaseRow>(
  'line-box',
  'which run sets the height of a line holding several sizes',
  baseRows,
  () => true,
  [
    {
      name: 'the-largest-run',
      says: 'the line box is 1.2 times the largest size on the line, whatever the order',
      predict: (row) => near(row.lineHeight, 1.2 * Math.max(...row.sizes), 0.01),
    },
    {
      name: 'the-first-run',
      says: 'the run that opens the line sets the box',
      predict: (row) => near(row.lineHeight, 1.2 * (row.sizes[0] ?? 0), 0.01),
    },
    {
      name: 'the-last-run',
      says: 'the run that closes the line sets the box',
      predict: (row) => near(row.lineHeight, 1.2 * (row.sizes.at(-1) ?? 0), 0.01),
    },
    {
      name: 'the-smallest-run',
      says: 'the line box shrinks to the smallest size on it',
      predict: (row) => near(row.lineHeight, 1.2 * Math.min(...row.sizes), 0.01),
    },
  ],
);

const baselineAnswer = choose<BaseRow>(
  'baseline',
  'how far below the top of the line box the baseline sits',
  baseRows,
  () => true,
  [
    {
      name: 'font-box-share',
      says: "the line box is split between ascent and descent in the face's own proportion",
      predict: (row) => {
        const box = faceBox(row.face);
        return near(row.drop, (row.lineHeight * box.ascent) / (box.ascent + box.descent), UNIT_PT);
      },
    },
    {
      name: 'autofit-coefficient',
      says: "the line box less 3.4's per-face last-line coefficient times the size",
      predict: (row) =>
        near(row.drop, row.lineHeight - (AUTOFIT_B[row.face] ?? APPROXIMATE_B) * row.size, UNIT_PT),
    },
    {
      name: 'css-half-leading',
      says: 'the CSS model: the leading split evenly above and below the font box',
      predict: (row) => {
        const box = faceBox(row.face);
        const leading = row.lineHeight - (box.ascent + box.descent) * row.size;
        return near(row.drop, leading / 2 + box.ascent * row.size, UNIT_PT);
      },
    },
    {
      name: 'ascent-from-the-top',
      says: "the face's ascent measured down from the top of the line box",
      predict: (row) => near(row.drop, faceBox(row.face).ascent * row.size, UNIT_PT),
    },
    {
      name: 'ink-box-share',
      says: 'the same proportion, taken from the ink of Hxg rather than the font box',
      predict: (row) => {
        const box = faceBox(row.face);
        const descent = box.ascent + box.descent - box.ink;
        return near(row.drop, (row.lineHeight * box.ink) / (box.ink + descent), UNIT_PT);
      },
    },
    {
      name: 'four-fifths-of-the-box',
      says: 'a face-independent fraction of the line box',
      predict: (row) => near(row.drop, row.lineHeight * 0.81, UNIT_PT),
    },
  ],
);

interface AdvanceRow {
  readonly probe: string;
  readonly gap: number;
  readonly lineHeight: number;
  readonly nextHeight: number;
}

/** Every place two lines follow one another, across every package. */
const advanceRows: AdvanceRow[] = [...rowsByQuestion.values()].flat().flatMap((row) => {
  const out: AdvanceRow[] = [];
  row.shape.lines.forEach((line, index) => {
    const next = row.shape.lines[index + 1];
    if (next?.top == null || line.top == null || line.height == null || next.height == null) return;
    out.push({
      probe: `${row.probe.id}#${String(index)}`,
      gap: pt(next.top - line.top, 4),
      lineHeight: line.height,
      nextHeight: next.height,
    });
  });
  return out;
});

choose<AdvanceRow>(
  'line-advance',
  'how far below one line the next one sits',
  advanceRows,
  () => true,
  [
    {
      name: 'one-line-box',
      says: "the next line's top is exactly this line's box below this line's top",
      predict: (row) => near(row.gap, row.lineHeight, 0.01),
    },
    {
      name: 'the-next-line-box',
      says: "the advance is the *following* line's box",
      predict: (row) => near(row.gap, row.nextHeight, 0.01),
    },
    {
      name: 'the-em-and-no-leading',
      says: 'lines stack at the font size with no leading between them',
      predict: (row) => near(row.gap, row.lineHeight / 1.2, 0.01),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q4 - where a line starts, and which lines stretch                          */
/* -------------------------------------------------------------------------- */

interface AlignRow {
  readonly row: Row;
  readonly algn: string;
  readonly wrapped: boolean;
  readonly lineIndex: number;
  readonly lastLine: boolean;
  /** The line's own left edge, relative to the content box, in points. */
  readonly offset: number;
  /** The same line's width when the paragraph is left-aligned. */
  readonly naturalWidth: number;
  /** The same, less whatever spaces the line ends with. */
  readonly anchorWidth: number;
  readonly width: number;
  readonly contentWidth: number;
}

/**
 * One space's advance, from the pair of probes that differ by exactly one.
 *
 * Derived rather than measured in a browser: the question is what PowerPoint
 * subtracts, so the width has to be PowerPoint's own.
 */
function spaceAdvance(rows: readonly Row[]): number {
  const width = (id: string): number => {
    const line = rows.find((row) => row.probe.id === id)?.shape.lines[0];
    if (line?.width == null) throw new Error(`${id} has no line width`);
    return line.width;
  };
  const one = width('align-trail-r-one-space') - width('align-trail-r-bare');
  const three = (width('align-trail-r-three-spaces') - width('align-trail-r-bare')) / 3;
  if (!near(one, three, 0.01)) {
    throw new Error(
      `one space measures ${String(one)}pt and three measure ${String(three)}pt each`,
    );
  }
  return one;
}

const alignRows: AlignRow[] = (() => {
  const rows = rowsFor('align');
  const space = spaceAdvance(rows);
  const naturalOf = new Map<string, readonly number[]>();
  const trailingOf = new Map<string, readonly number[]>();
  for (const kind of ['one', 'wrap'] as const) {
    const left = rows.find((row) => row.probe.id === `align-${kind}-l`);
    if (left === undefined) throw new Error(`no left-aligned ${kind} probe`);
    naturalOf.set(
      kind,
      left.shape.lines.map((line) => {
        if (line.width == null) throw new Error('a line reported no width');
        return line.width;
      }),
    );
    trailingOf.set(
      kind,
      left.shape.lines.map((line) => TRAILING_SPACES.exec(line.text ?? '')?.[0].length ?? 0),
    );
  }

  const out: AlignRow[] = [];
  for (const row of rows) {
    const match = /^align-(one|wrap)-(\w+)$/.exec(row.probe.id);
    if (match === null) continue;
    const kind = match[1] === 'wrap' ? 'wrap' : 'one';
    const natural = naturalOf.get(kind);
    const trailing = trailingOf.get(kind);
    if (natural === undefined || trailing === undefined) throw new Error('unreachable');
    row.shape.lines.forEach((line, index) => {
      if (line.left == null || line.width == null) throw new Error('a line reported no box');
      const width = natural[index];
      const spaces = trailing[index];
      if (width === undefined || spaces === undefined) {
        throw new Error(`${row.probe.id} has more lines than the left one`);
      }
      out.push({
        row,
        algn: match[2] ?? '',
        wrapped: kind === 'wrap',
        lineIndex: index,
        lastLine: index === row.shape.lines.length - 1,
        offset: pt(line.left - row.probe.rect.x, 3),
        naturalWidth: width,
        anchorWidth: width - spaces * space,
        width: line.width,
        contentWidth: row.probe.rect.w,
      });
    });
  }
  return out;
})();

const STRETCHES_LAST = new Set(['dist']);
const STRETCHES = new Set(['just', 'justLow', 'dist', 'thaiDist']);

/**
 * A hundredth of a point.
 *
 * `BoundWidth` overstates a line by about five thousandths of a point, so an
 * offset computed from it lands beside the one PowerPoint anchored on; the
 * nearest rival reading here is three points away.
 */
const ALIGN_TOLERANCE = 0.01;

/** Where a candidate's rule puts a line, given the width it anchors. */
function anchored(row: AlignRow, width: number): number {
  if (row.algn === 'ctr') return (row.contentWidth - width) / 2;
  if (row.algn === 'r') return row.contentWidth - width;
  return 0;
}

choose<AlignRow>('align', 'where each line starts inside the content box', alignRows, () => true, [
  {
    name: 'anchor-without-the-trailing-space',
    says: 'a justifying alignment starts at the left edge; ctr and r anchor the line less whatever spaces it ends with',
    predict: (row) => near(row.offset, anchored(row, row.anchorWidth), ALIGN_TOLERANCE),
  },
  {
    name: 'anchor-the-whole-line',
    says: 'the trailing spaces are part of the width an alignment anchors',
    predict: (row) => near(row.offset, anchored(row, row.naturalWidth), ALIGN_TOLERANCE),
  },
  {
    name: 'centre-is-left',
    says: 'ctr and r anchor nothing and every line starts at the left edge',
    predict: (row) => near(row.offset, 0, ALIGN_TOLERANCE),
  },
  {
    name: 'anchor-the-stretched-width',
    says: 'ctr and r are computed from the width the line actually reports',
    predict: (row) => near(row.offset, anchored(row, row.width), ALIGN_TOLERANCE),
  },
]);

interface StretchRow extends AlignRow {
  readonly stretched: boolean;
}

const stretchRows: StretchRow[] = alignRows.map((row) => ({
  ...row,
  stretched: !near(row.width, row.naturalWidth, 0.5),
}));

choose<StretchRow>(
  'justify',
  'which lines a justifying alignment stretches',
  stretchRows,
  (row) => row.stretched,
  [
    {
      name: 'dist-stretches-the-last-line',
      says: 'just, justLow and thaiDist leave the last line alone; dist stretches every line',
      predict: (row) => STRETCHES.has(row.algn) && (STRETCHES_LAST.has(row.algn) || !row.lastLine),
    },
    {
      name: 'every-justify-spares-the-last',
      says: 'no justifying alignment stretches a last line',
      predict: (row) => STRETCHES.has(row.algn) && !row.lastLine,
    },
    {
      name: 'every-justify-stretches-everything',
      says: 'a justifying alignment stretches every line it has',
      predict: (row) => STRETCHES.has(row.algn),
    },
    {
      name: 'nothing-stretches',
      says: 'justification is not implemented at all',
      predict: () => false,
    },
  ],
);

interface TrailRow {
  readonly row: Row;
  readonly algn: string;
  readonly spaces: number;
  readonly offset: number;
  readonly bareWidth: number;
  readonly ownWidth: number;
  readonly contentWidth: number;
}

const trailRows: TrailRow[] = (() => {
  const rows = rowsFor('align');
  const out: TrailRow[] = [];
  for (const row of rows) {
    const match = /^align-trail-(\w+)-(bare|one-space|three-spaces)$/.exec(row.probe.id);
    if (match === null) continue;
    const algn = match[1] ?? '';
    const bare = rows.find((other) => other.probe.id === `align-trail-${algn}-bare`);
    const line = row.shape.lines[0];
    const bareLine = bare?.shape.lines[0];
    if (line?.left == null || line.width == null || bareLine?.width == null) {
      throw new Error(`${row.probe.id} has no line box`);
    }
    out.push({
      row,
      algn,
      spaces: match[2] === 'bare' ? 0 : match[2] === 'one-space' ? 1 : 3,
      offset: pt(line.left - row.probe.rect.x, 3),
      bareWidth: bareLine.width,
      ownWidth: line.width,
      contentWidth: row.probe.rect.w,
    });
  }
  return out;
})();

choose<TrailRow>(
  'trailing-space',
  'whether a trailing space counts towards the width an alignment anchors',
  trailRows,
  () => true,
  [
    {
      name: 'trailing-space-hangs',
      says: 'alignment uses the width without the trailing spaces, which hang past the edge',
      predict: (row) =>
        near(
          row.offset,
          row.algn === 'ctr'
            ? (row.contentWidth - row.bareWidth) / 2
            : row.contentWidth - row.bareWidth,
          ALIGN_TOLERANCE,
        ),
    },
    {
      name: 'trailing-space-counts',
      says: 'the spaces are part of the line and move it',
      predict: (row) =>
        near(
          row.offset,
          row.algn === 'ctr'
            ? (row.contentWidth - row.ownWidth) / 2
            : row.contentWidth - row.ownWidth,
          ALIGN_TOLERANCE,
        ),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q5 - underline and strikethrough                                           */
/* -------------------------------------------------------------------------- */

interface Rule {
  /** The rule's top edge, in points below the baseline; negative is above it. */
  readonly offset: number;
  readonly thickness: number;
  readonly left: number;
  readonly right: number;
}

/**
 * Every full-width horizontal rule on a slide, in points.
 *
 * A row of glyphs can span the text's whole width but is never nine tenths ink,
 * which is what separates a rule from the middle of a lowercase word.
 */
function rulesIn(bitmap: Bitmap, baseline: number): readonly Rule[] {
  const spans: { y: number; left: number; right: number; count: number }[] = [];
  for (let y = 0; y < bitmap.height; y++) {
    let left = -1;
    let right = -1;
    let count = 0;
    for (let x = 0; x < bitmap.width; x++) {
      const [r, g, b] = bitmap.pixel(x, y);
      if (r >= 200 && g >= 200 && b >= 200) continue;
      if (left < 0) left = x;
      right = x;
      count++;
    }
    if (left >= 0) spans.push({ y, left, right, count });
  }
  if (spans.length === 0) return [];
  const widest = Math.max(...spans.map((span) => span.right - span.left));
  const solid = spans.filter(
    (span) => span.right - span.left >= widest - 4 && span.count > (span.right - span.left) * 0.9,
  );

  const rules: Rule[] = [];
  let run: { from: number; to: number; left: number; right: number } | null = null;
  for (const span of solid) {
    if (run !== null && span.y === run.to + 1) {
      run.to = span.y;
      run.left = Math.min(run.left, span.left);
      run.right = Math.max(run.right, span.right);
      continue;
    }
    if (run !== null) rules.push(ruleOf(run, baseline));
    run = { from: span.y, to: span.y, left: span.left, right: span.right };
  }
  if (run !== null) rules.push(ruleOf(run, baseline));
  return rules;
}

function ruleOf(
  run: { from: number; to: number; left: number; right: number },
  baseline: number,
): Rule {
  return {
    offset: pt(run.from / PX_PER_POINT - baseline, 3),
    thickness: pt((run.to - run.from + 1) / PX_PER_POINT, 3),
    left: pt(run.left / PX_PER_POINT, 3),
    right: pt(run.right / PX_PER_POINT, 3),
  };
}

interface DecRow {
  readonly row: Row;
  readonly kind: 'u' | 's' | 'face' | 'trail';
  readonly style: string;
  readonly face: string;
  readonly size: number;
  readonly baseline: number;
  readonly rules: readonly Rule[];
  /** The rightmost glyph ink above the first rule, in points. */
  readonly glyphRight: number | null;
}

/** The rightmost ink strictly above `y`, which on these probes is the last glyph. */
function inkRightAbove(bitmap: Bitmap, y: number): number | null {
  let right = -1;
  const limit = Math.min(bitmap.height, Math.max(0, Math.floor(y * PX_PER_POINT)));
  for (let row = 0; row < limit; row++) {
    for (let x = bitmap.width - 1; x > right; x--) {
      const [r, g, b] = bitmap.pixel(x, row);
      if (r < 200 || g < 200 || b < 200) {
        right = x;
        break;
      }
    }
  }
  return right < 0 ? null : pt((right + 1) / PX_PER_POINT, 3);
}

const decRows: DecRow[] = rowsFor('decoration').map((row) => {
  const text = row.texts[0];
  if (text === undefined) throw new Error(`${row.probe.id} drew no text record`);
  if (row.bitmap === null) throw new Error(`${row.probe.id} has no bitmap`);
  const size = pt(-(text.font?.height ?? 0) / UNITS_PER_POINT, 2);
  const baseline = text.origin[1] / UNITS_PER_POINT;
  const match = /^dec-(u|s|face|trail)-([A-Za-z-]+?)(?:-(\d+))?$/.exec(row.probe.id);
  const faceKind = /^dec-face-(u|s)-/.exec(row.probe.id);
  if (match === null) throw new Error(`${row.probe.id} is not a decoration probe id`);
  const rules = rulesIn(row.bitmap, baseline);
  const first = rules[0];
  return {
    row,
    kind: (match[1] ?? 'u') as DecRow['kind'],
    style: faceKind === null ? (match[2] ?? '') : (faceKind[1] ?? ''),
    face: text.font?.face ?? '',
    size,
    baseline,
    rules,
    glyphRight: first === undefined ? null : inkRightAbove(row.bitmap, baseline + first.offset),
  };
});

/** A rule's offset and thickness as a fraction of the em, which is what scales. */
function emOf(rule: Rule, size: number): { offset: number; thickness: number } {
  return { offset: rule.offset / size, thickness: rule.thickness / size };
}

interface DecScaleRow {
  readonly id: string;
  readonly style: string;
  readonly size: number;
  readonly offsetEm: number;
  readonly thicknessEm: number;
  readonly reference: { offsetEm: number; thicknessEm: number };
}

/** One pixel at the smallest probed size, which is the coarsest reading here. */
const DEC_TOLERANCE = 1 / PX_PER_POINT / 18;

const decScaleRows: DecScaleRow[] = (() => {
  const at40 = new Map<string, DecRow>();
  for (const row of decRows) if (near(row.size, 40, 0.1)) at40.set(`${row.kind}-${row.style}`, row);
  const out: DecScaleRow[] = [];
  for (const row of decRows) {
    if (row.kind === 'face' || row.kind === 'trail') continue;
    if (row.style === 'trailing-space') continue;
    const reference = at40.get(`${row.kind}-${row.style}`);
    const mine = row.rules[0];
    const theirs = reference?.rules[0];
    if (mine === undefined || theirs === undefined || reference === undefined) continue;
    out.push({
      id: row.row.probe.id,
      style: row.style,
      size: row.size,
      ...emOf(mine, row.size),
      offsetEm: emOf(mine, row.size).offset,
      thicknessEm: emOf(mine, row.size).thickness,
      reference: {
        offsetEm: emOf(theirs, reference.size).offset,
        thicknessEm: emOf(theirs, reference.size).thickness,
      },
    });
  }
  return out;
})();

choose<DecScaleRow>(
  'decoration-scale',
  'whether a rule keeps its place when the size changes',
  decScaleRows,
  () => true,
  [
    {
      name: 'proportional-to-the-size',
      says: 'offset and thickness are both fractions of the em, so a rule scales with the text',
      predict: (row) =>
        near(row.offsetEm, row.reference.offsetEm, DEC_TOLERANCE) &&
        near(row.thicknessEm, row.reference.thicknessEm, DEC_TOLERANCE),
    },
    {
      name: 'a-fixed-length',
      says: 'the rule is the same number of points whatever the size',
      predict: (row) =>
        near(row.offsetEm * row.size, row.reference.offsetEm * 40, 0.5) &&
        near(row.thicknessEm * row.size, row.reference.thicknessEm * 40, 0.5),
    },
  ],
);

interface DecFaceRow {
  readonly probe: string;
  readonly kind: 'u' | 's';
  readonly face: string;
  readonly size: number;
  readonly offset: number;
  readonly thickness: number;
  readonly arialOffsetEm: number;
  readonly chromium: { offsetEm: number; thicknessEm: number };
}

/** Half a point, which at 1920 pixels over 960 is one pixel. */
const RULE_TOLERANCE = 0.5;

const decFaceRows: DecFaceRow[] = (() => {
  const arialOf = new Map<string, number>();
  for (const kind of ['u', 's'] as const) {
    const arial = decRows.find((row) => row.row.probe.id === `dec-face-${kind}-Arial`);
    const rule = arial?.rules[0];
    if (arial === undefined || rule === undefined) throw new Error(`no Arial ${kind} row`);
    arialOf.set(kind, emOf(rule, arial.size).offset);
  }
  return decRows
    .filter((row) => row.kind === 'face')
    .map((row) => {
      const rule = row.rules[0];
      const arialOffsetEm = arialOf.get(row.style);
      const drawn = browser.decorations[row.face];
      if (rule === undefined || arialOffsetEm === undefined || drawn === undefined) {
        throw new Error(`${row.row.probe.id} has no rule or no browser reading`);
      }
      return {
        probe: row.row.probe.id,
        kind: row.style === 's' ? 's' : 'u',
        face: row.face,
        size: row.size,
        offset: rule.offset,
        thickness: rule.thickness,
        arialOffsetEm,
        chromium: row.style === 's' ? drawn.strike : drawn.underline,
      };
    });
})();

refute<DecFaceRow>(
  'decoration-face',
  'whether either cheap way of placing a rule reproduces PowerPoint',
  "neither: the browser's own text-decoration is in the wrong place for every face, and one offset does not fit eight faces, so the offsets are a measured table",
  decFaceRows,
  [
    {
      name: 'the-browser-already-agrees',
      says: 'text-decoration puts the rule where PowerPoint does, so a renderer can just ask for it',
      predict: (row) => near(row.offset, row.chromium.offsetEm * row.size, RULE_TOLERANCE),
    },
    {
      name: 'one-offset-for-every-face',
      says: "every face's rule sits the same fraction of the em from the baseline",
      predict: (row) => near(row.offset, row.arialOffsetEm * row.size, RULE_TOLERANCE),
    },
    {
      name: 'the-browser-thickness-agrees',
      says: 'the rule is as thick as text-decoration draws it',
      predict: (row) => near(row.thickness, row.chromium.thicknessEm * row.size, RULE_TOLERANCE),
    },
  ],
);

interface DecTrailRow {
  readonly id: string;
  /** How far the rule runs past the last glyph, in points. */
  readonly overhang: number;
  readonly spaces: number;
  readonly spaceAdvance: number;
}

const decTrailRows: DecTrailRow[] = (() => {
  const byId = new Map(decRows.map((row) => [row.row.probe.id, row]));
  const out: DecTrailRow[] = [];
  for (const id of ['one', 'three', 'georgia']) {
    const spaced = byId.get(`dec-trail-${id}`);
    const bare = byId.get(`dec-trail-${id}-bare`);
    const spacedRule = spaced?.rules[0];
    const bareRule = bare?.rules[0];
    if (
      spaced === undefined ||
      bare === undefined ||
      spacedRule === undefined ||
      bareRule === undefined
    ) {
      throw new Error(`dec-trail-${id} did not draw a rule`);
    }
    if (spaced.glyphRight === null) throw new Error(`dec-trail-${id} drew no glyphs`);
    // The bare probe is the same string with the spaces deleted, so the width
    // its rule covers is the width a rule that skips them would cover.
    out.push({
      id,
      overhang: pt(spacedRule.right - spaced.glyphRight, 3),
      spaces: id === 'three' ? 3 : 1,
      spaceAdvance: pt(spacedRule.right - bareRule.right, 3),
    });
  }
  return out;
})();

choose<DecTrailRow>(
  'decoration-extent',
  'whether an underline runs under the spaces a line ends with',
  decTrailRows,
  () => true,
  [
    {
      name: 'the-rule-stops-at-the-last-glyph',
      says: 'the spaces a line ends with are not underlined',
      predict: (row) => near(row.spaceAdvance, 0, 0.6) && near(row.overhang, 0, 1.5),
    },
    {
      name: 'the-rule-covers-the-spaces',
      says: 'the rule runs the full advance of the line, trailing spaces included',
      predict: (row) => row.spaceAdvance > 0.6 * row.spaces,
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q6 - baseline shift and capitalisation                                     */
/* -------------------------------------------------------------------------- */

interface ShiftRow {
  readonly row: Row;
  readonly percent: number;
  readonly statedSize: number;
  readonly drawnSize: number;
  readonly rise: number;
}

const shiftRows: ShiftRow[] = rowsFor('shift')
  .filter((row) => row.probe.id.startsWith('shift-up') || row.probe.id.startsWith('shift-down'))
  .map((row) => {
    const match = /^shift-(up|down)(\d+)$/.exec(row.probe.id);
    const plain = row.texts[0];
    const shifted = row.texts[1];
    if (match === null || plain === undefined || shifted === undefined) {
      throw new Error(`${row.probe.id} did not draw two records`);
    }
    return {
      row,
      percent: (match[1] === 'down' ? -1 : 1) * (Number(match[2]) / 1000),
      statedSize: pt(-(plain.font?.height ?? 0) / UNITS_PER_POINT, 2),
      drawnSize: pt(-(shifted.font?.height ?? 0) / UNITS_PER_POINT, 2),
      rise: pt((plain.origin[1] - shifted.origin[1]) / UNITS_PER_POINT, 3),
    };
  });

choose<ShiftRow>(
  'baseline-shift',
  'how far a:rPr/@baseline moves a run and at what size it is drawn',
  shiftRows,
  () => true,
  [
    {
      name: 'percent-of-the-stated-size',
      says: 'the run rises by @baseline per cent of the size it asked for, drawn at two thirds of it',
      predict: (row) =>
        near(row.rise, (row.percent / 100) * row.statedSize, 2 * UNIT_PT) &&
        near(row.drawnSize, (2 / 3) * row.statedSize, 0.15),
    },
    {
      name: 'percent-of-the-drawn-size',
      says: 'the rise is a fraction of the reduced size',
      predict: (row) =>
        near(row.rise, (row.percent / 100) * row.drawnSize, 2 * UNIT_PT) &&
        near(row.drawnSize, (2 / 3) * row.statedSize, 0.15),
    },
    {
      name: 'percent-of-the-line-box',
      says: 'the rise is a fraction of the 1.2 line box',
      predict: (row) =>
        near(row.rise, (row.percent / 100) * 1.2 * row.statedSize, 2 * UNIT_PT) &&
        near(row.drawnSize, (2 / 3) * row.statedSize, 0.15),
    },
    {
      name: 'shifted-runs-keep-their-size',
      says: 'the rise is right but the glyphs are not shrunk',
      predict: (row) =>
        near(row.rise, (row.percent / 100) * row.statedSize, 2 * UNIT_PT) &&
        near(row.drawnSize, row.statedSize, 0.15),
    },
  ],
);

/**
 * A drawn string without its spaces.
 *
 * A space needs no ink, so PowerPoint leaves it out of the drawing calls
 * altogether; comparing what was drawn to what was asked for has to too.
 */
function letters(text: string): string {
  return text.replace(/ /g, '');
}

interface CapsRow {
  readonly row: Row;
  readonly cap: string;
  readonly drawn: readonly { text: string; size: number }[];
  readonly statedSize: number;
}

const capsRows: CapsRow[] = rowsFor('shift')
  .filter((row) => row.probe.id.startsWith('shift-cap-'))
  .map((row) => {
    const first = row.texts[0];
    if (first === undefined) throw new Error(`${row.probe.id} drew no text`);
    return {
      row,
      cap: row.probe.id.replace('shift-cap-', ''),
      statedSize: pt(-(first.font?.height ?? 0) / UNITS_PER_POINT, 2),
      drawn: row.texts.map((text) => ({
        text: text.text,
        size: pt(-(text.font?.height ?? 0) / UNITS_PER_POINT, 2),
      })),
    };
  });

choose<CapsRow>('caps', 'what a:rPr/@cap draws, and at what size', capsRows, () => true, [
  {
    name: 'small-caps-are-four-fifths',
    says: 'cap="all" uppercases at the full size; cap="small" uppercases the lowercase at four fifths of it',
    predict: (row) => {
      const joined = letters(row.drawn.map((d) => d.text).join(''));
      const sizes = row.drawn.map((d) => d.size);
      if (row.cap === 'none')
        return (
          joined === letters('Hamburg fox') && sizes.every((s) => near(s, row.statedSize, 0.1))
        );
      if (row.cap === 'all') {
        return (
          joined === letters('HAMBURG FOX') && sizes.every((s) => near(s, row.statedSize, 0.1))
        );
      }
      return (
        joined === letters('HAMBURG FOX') &&
        sizes.some((s) => near(s, 0.8 * row.statedSize, 0.15)) &&
        sizes.some((s) => near(s, row.statedSize, 0.1))
      );
    },
  },
  {
    name: 'small-caps-are-three-quarters',
    says: 'the reduced capital is three quarters of the size',
    predict: (row) => {
      const joined = letters(row.drawn.map((d) => d.text).join(''));
      const sizes = row.drawn.map((d) => d.size);
      if (row.cap === 'none')
        return (
          joined === letters('Hamburg fox') && sizes.every((s) => near(s, row.statedSize, 0.1))
        );
      if (row.cap === 'all') {
        return (
          joined === letters('HAMBURG FOX') && sizes.every((s) => near(s, row.statedSize, 0.1))
        );
      }
      return (
        joined === letters('HAMBURG FOX') && sizes.some((s) => near(s, 0.75 * row.statedSize, 0.15))
      );
    },
  },
  {
    name: 'caps-change-only-the-glyphs',
    says: 'both values uppercase and neither changes the size',
    predict: (row) => {
      const joined = letters(row.drawn.map((d) => d.text).join(''));
      const sizes = row.drawn.map((d) => d.size);
      const upper = row.cap === 'none' ? letters('Hamburg fox') : letters('HAMBURG FOX');
      return joined === upper && sizes.every((s) => near(s, row.statedSize, 0.1));
    },
  },
]);

/* -------------------------------------------------------------------------- */
/* Q7 - what a run boundary does                                              */
/* -------------------------------------------------------------------------- */

interface RunRow {
  readonly id: string;
  readonly advances: readonly number[];
  readonly whole: readonly number[];
  readonly records: number;
}

const runRows: RunRow[] = (() => {
  const rows = rowsFor('runs');
  const wholeOf = new Map<string, readonly number[]>();
  for (const key of ['unset', 'off']) {
    const row = rows.find((r) => r.probe.id === `runs-whole-${key}`);
    const text = row?.texts[0];
    if (text === undefined) throw new Error(`no whole-string probe for kern ${key}`);
    wholeOf.set(key, text.dx);
  }
  return rows
    .filter((row) => row.probe.id !== 'runs-spc')
    .map((row) => {
      const key = row.probe.id.endsWith('-off') ? 'off' : 'unset';
      const whole = wholeOf.get(key);
      if (whole === undefined) throw new Error('unreachable');
      return {
        id: row.probe.id,
        advances: row.texts.flatMap((text) => text.dx),
        whole,
        records: row.texts.length,
      };
    });
})();

choose<RunRow>(
  'runs',
  'whether a run boundary changes the advances either side of it',
  runRows,
  (row) => row.advances.join(','),
  [
    {
      name: 'the-line-is-shaped-whole',
      says: 'a run boundary splits the drawing calls and nothing else: the kern crosses it',
      predict: (row) => row.whole.join(','),
    },
    {
      name: 'each-run-is-shaped-alone',
      says: 'shaping restarts at a run boundary, so the pair straddling it does not kern',
      predict: (row) => {
        if (row.records < 2) return row.whole.join(',');
        const unkerned = [...row.whole];
        // The pair straddling the boundary loses its kern, which shows on the
        // last character of the first run.
        const at = 1;
        const value = unkerned[at];
        const following = unkerned[at + 1];
        if (value === undefined || following === undefined) return row.whole.join(',');
        unkerned[at] = value + 25;
        return unkerned.join(',');
      },
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const spcRow = (() => {
  const rows = rowsFor('runs');
  const plain = rows.find((row) => row.probe.id === 'runs-whole-unset')?.texts[0];
  const spaced = rows.find((row) => row.probe.id === 'runs-spc')?.texts[0];
  if (plain === undefined || spaced === undefined) throw new Error('no @spc row');
  return {
    plain: plain.dx,
    spaced: spaced.dx,
    addedPerCharacter: spaced.dx.map((value, index) => value - (plain.dx[index] ?? 0)),
  };
})();

const fixture = {
  $comment:
    'Experiment T8: what a renderer must know to draw text where PowerPoint draws it. ' +
    'Regenerate with tools/ground-truth/render/text/analyse.ts. ' +
    'Lengths are points unless a field says otherwise.',
  experiment: 'T8',
  subPhase: '3.8',
  adr: 'docs/adr/phase-3-text/0034-text-in-both-renderers.md',
  units: {
    slidePoints: { width: 960, height: 540 },
    exportPixels: { width: 1920, height: 1080 },
    emfLogicalUnitsPerInch: 600,
  },
  measuredOn: {
    probes: inputs.probes.length,
    packages: readings.decks.length,
    everyPackageOpenedWithoutRepair: readings.decks.every(
      (deck) => deck.opened && deck.repaired === false,
    ),
  },
  questions: Object.fromEntries(
    scored.map((entry) => [
      entry.question,
      {
        asks: entry.asks,
        answer: entry.answer,
        says: entry.says,
        rows: entry.total,
        scores: entry.scores,
      },
    ]),
  ),
  refuted: scored
    .filter((entry) => entry.answer === 'none-of-these')
    .map((entry) => entry.question),
  powerpoint: {
    flip: {
      answer: flipAnswer.name,
      rows: flipRows.map((row) => ({
        probe: row.row.probe.id,
        rot: row.rot,
        flipH: row.flipH,
        flipV: row.flipV,
        extraTurn: row.extraTurn,
        mirrored: row.mirrored,
      })),
    },
    rotation: rotRows.map((row) => ({
      probe: row.row.probe.id,
      rot: row.rot,
      upright: row.upright,
      bodyRot: row.bodyRot,
      angle: row.angle,
    })),
    uprightBox: uprightRows.map((row) => ({
      probe: row.row.probe.id,
      rot: row.rot,
      boxLeft: row.boxLeft,
    })),
    baseline: {
      answer: baselineAnswer.name,
      rows: baseRows.map((row) => ({
        probe: row.row.probe.id,
        face: row.face,
        size: row.size,
        sizes: row.sizes,
        lineHeight: pt(row.lineHeight, 3),
        drop: row.drop,
        share: pt(row.drop / row.lineHeight, 6),
      })),
    },
    lineAdvance: advanceRows,
    align: alignRows.map((row) => ({
      probe: row.row.probe.id,
      algn: row.algn,
      line: row.lineIndex,
      lastLine: row.lastLine,
      offset: row.offset,
      width: pt(row.width, 3),
      naturalWidth: pt(row.naturalWidth, 3),
      anchorWidth: pt(row.anchorWidth, 3),
    })),
    trailingSpace: trailRows.map((row) => ({
      probe: row.row.probe.id,
      algn: row.algn,
      spaces: row.spaces,
      offset: row.offset,
      bareWidth: pt(row.bareWidth, 3),
      ownWidth: pt(row.ownWidth, 3),
    })),
    decorationFace: decFaceRows.map((row) => ({
      probe: row.probe,
      kind: row.kind,
      face: row.face,
      size: row.size,
      offsetEm: pt(row.offset / row.size, 6),
      thicknessEm: pt(row.thickness / row.size, 6),
      chromiumOffsetEm: pt(row.chromium.offsetEm, 6),
      chromiumThicknessEm: pt(row.chromium.thicknessEm, 6),
    })),
    decorationExtent: decTrailRows,
    decoration: decRows.map((row) => ({
      probe: row.row.probe.id,
      face: row.face,
      size: row.size,
      glyphRight: row.glyphRight,
      rules: row.rules.map((rule) => ({
        ...rule,
        offsetEm: pt(rule.offset / row.size, 5),
        thicknessEm: pt(rule.thickness / row.size, 5),
      })),
    })),
    shift: shiftRows.map((row) => ({
      probe: row.row.probe.id,
      percent: row.percent,
      statedSize: row.statedSize,
      drawnSize: row.drawnSize,
      rise: row.rise,
      sizeRatio: pt(row.drawnSize / row.statedSize, 5),
    })),
    caps: capsRows.map((row) => ({
      probe: row.row.probe.id,
      cap: row.cap,
      statedSize: row.statedSize,
      drawn: row.drawn,
    })),
    runs: runRows.map((row) => ({ probe: row.id, records: row.records, advances: row.advances })),
    letterSpacing: spcRow,
  },
  browser: {
    chromium: browser.chromium,
    ladderPx: LADDER_PX,
    faces: Object.fromEntries(
      FACES.map((face) => {
        const box = faceBox(face);
        return [
          face,
          {
            ascent: pt(box.ascent, 6),
            descent: pt(box.descent, 6),
            baselineShare: pt(box.ascent / (box.ascent + box.descent), 6),
          },
        ];
      }),
    ),
    sizes: SIZES,
  },
};

const config = await resolveConfig(fixturePath);
writeFileSync(
  fixturePath,
  await format(JSON.stringify(fixture), { ...config, filepath: fixturePath }),
);

for (const entry of scored) {
  const rivals = Object.entries(entry.scores)
    .filter(([name]) => name !== entry.answer)
    .map(([name, score]) => `${name} ${score}`)
    .join(', ');
  console.log(
    `${entry.question.padEnd(18)} ${entry.answer.padEnd(30)} ${String(entry.total)}/${String(entry.total)}   ${rivals}`,
  );
}
console.log(`\n${String(inputs.probes.length)} probe(s) -> ${fixturePath}`);

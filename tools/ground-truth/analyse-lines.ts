/**
 * Experiment C4, step 3 - read the bitmaps and say what PowerPoint did.
 *
 * ```
 * node tools/ground-truth/analyse-lines.ts <work-dir> [--fixture corpus/ground-truth/lines.json]
 * ```
 *
 * Each section states its hypothesis before it prints its evidence, so a reader
 * can disagree with the question and not only with the answer. The fixture it
 * writes is the committed measurement; everything printed here is derived from
 * it, and nothing in `@pptx-studio/paint` depends on this file having been run.
 *
 * ## What gets stored, and why it is not the raw pixels
 *
 * C3 committed whole sample strips because a gradient's answer *is* a curve. A
 * stroke's answer usually is not. Where a dash begins and ends is a set of
 * positions, and a position is recovered far more precisely than a pixel: the
 * edge of an antialiased run crosses half coverage somewhere *inside* one pixel,
 * and interpolating between the two straddling samples locates it to about a
 * tenth of one. So for anything whose answer is a position, the fixture stores
 * sub-pixel **crossings**, in points, and the raw row is not kept.
 *
 * Where the answer really is a shape - a blur's falloff, a soft edge's ramp, an
 * arrowhead's outline - the profile is stored, quantised to a byte per sample.
 * The rule is mechanical: a read region of 800 samples or fewer keeps its
 * coverage; a longer one keeps only its crossings. Every dash row is longer than
 * that and every profile row is shorter, which is not a coincidence - the probes
 * were sized that way.
 *
 * ## Coverage is `1 - min(r,g,b)/255`
 *
 * Not luminance, and not one channel. Every ink in C4 is either black, a
 * saturated primary, or a known grey, and for all of those the minimum channel
 * over a white slide is exactly the ink's alpha. Luminance would report a
 * saturated red as 70% covered; a single channel would report a saturated green
 * as 0.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { readBmp, type Bitmap } from './bmp.ts';
import { ARROW_SIZES, ARROW_TYPES, FOLKLORE_DASHES, PRESET_DASHES, type ReadSpec } from './lines.ts';

/* -------------------------------------------------------------------------- */
/* inputs                                                                     */
/* -------------------------------------------------------------------------- */

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse-lines.ts <work-dir> [--fixture <path>]');
// Narrowed once, here: `bitmap` below is a hoisted declaration and does not see
// the narrowing a throw established on a module-level `const`.
const dir: string = arg;
const fixtureFlag = process.argv.indexOf('--fixture');
const fixturePath = fixtureFlag === -1 ? null : process.argv[fixtureFlag + 1];

interface ProbeInput {
  id: string;
  deck: string;
  group: string;
  question: string;
  x: number;
  y: number;
  cx: number;
  cy: number;
  read: ReadSpec;
  line: string;
  effect: string;
  geom: string;
}

interface Inputs {
  slidePoints: { w: number; h: number };
  fineDecks: string[];
  decks: { deck: string; file: string; probes: number }[];
  probes: ProbeInput[];
}

interface ComShape {
  id: string;
  weight: number | null;
  dashStyle: number | null;
  lineStyle: number | null;
  insetPen: number | null;
  endType: number | null;
  endLen: number | null;
  endWidth: number | null;
  shadowType: number | null;
  shadowBlur: number | null;
  shadowOffX: number | null;
  shadowOffY: number | null;
  glowRadius: number | null;
  softEdgeRad: number | null;
}

interface DeckReadback {
  deck: string;
  opened: boolean;
  repaired: boolean | null;
  error: string | null;
  shapes: ComShape[];
  bitmaps: { file: string; slide: number; width: number; height: number }[];
}

const inputs = JSON.parse(readFileSync(join(dir, 'line-inputs.json'), 'utf8')) as Inputs;
const readback = JSON.parse(readFileSync(join(dir, 'com-readback-lines.json'), 'utf8')) as {
  decks: DeckReadback[];
};

const deckOf = new Map(readback.decks.map((d) => [d.deck, d]));
const bitmapCache = new Map<string, Bitmap>();

function bitmap(file: string): Bitmap {
  const cached = bitmapCache.get(file);
  if (cached) return cached;
  const bmp = readBmp(readFileSync(join(dir, file)));
  bitmapCache.set(file, bmp);
  return bmp;
}

/** The finest export available for a deck, and how many pixels one point is. */
function finest(deck: string): { file: string; scale: number } | null {
  const record = deckOf.get(deck);
  if (record === undefined || record.bitmaps.length === 0) return null;
  const best = [...record.bitmaps].sort((a, b) => b.width - a.width)[0]!;
  return { file: best.file, scale: best.width / inputs.slidePoints.w };
}

/** The 1920-wide export specifically, for anything stored as a profile. */
function coarse(deck: string): { file: string; scale: number } | null {
  const record = deckOf.get(deck);
  if (record === undefined) return null;
  const found = record.bitmaps.find((b) => b.width === 1920);
  return found === undefined ? null : { file: found.file, scale: found.width / inputs.slidePoints.w };
}

/* -------------------------------------------------------------------------- */
/* sampling                                                                   */
/* -------------------------------------------------------------------------- */

/** Ink coverage over a white slide: `1 - min(r,g,b)/255`. See the header. */
function ink(bmp: Bitmap, x: number, y: number): number {
  const [r, g, b] = bmp.pixel(x, y);
  return 1 - Math.min(r, g, b) / 255;
}

/**
 * The pixel row or column whose centre is nearest a point coordinate, clamped
 * into the bitmap.
 *
 * The clamp is deliberate rather than defensive. A read window that runs off
 * the slide is a probe laid out wrongly, and clamping turns that into a
 * plausible-looking measurement of the slide's edge - so this reports it
 * instead, and the caller decides.
 */
function pixelAt(points: number, scale: number, limit: number, what: string): number {
  const at = Math.floor(points * scale);
  if (at < 0 || at >= limit) {
    throw new Error(
      `${what}: ${String(points)}pt is pixel ${String(at)}, outside 0..${String(limit - 1)} - the probe is off the slide`,
    );
  }
  return at;
}

interface Sampled {
  /** Coverage per sample, in order. */
  readonly cover: number[];
  /** The point coordinate of sample 0's centre. */
  readonly start: number;
  /** Points per sample. */
  readonly step: number;
}

function sampleRow(bmp: Bitmap, scale: number, y: number, x0: number, x1: number, what: string): Sampled {
  const row = pixelAt(y, scale, bmp.height, what);
  const from = Math.max(0, Math.round(x0 * scale));
  const to = Math.min(bmp.width - 1, Math.round(x1 * scale));
  const cover: number[] = [];
  for (let x = from; x <= to; x++) cover.push(ink(bmp, x, row));
  return { cover, start: (from + 0.5) / scale, step: 1 / scale };
}

function sampleCol(bmp: Bitmap, scale: number, x: number, y0: number, y1: number, what: string): Sampled {
  const col = pixelAt(x, scale, bmp.width, what);
  const from = Math.max(0, Math.round(y0 * scale));
  const to = Math.min(bmp.height - 1, Math.round(y1 * scale));
  const cover: number[] = [];
  for (let y = from; y <= to; y++) cover.push(ink(bmp, col, y));
  return { cover, start: (from + 0.5) / scale, step: 1 / scale };
}

/**
 * Where coverage crosses one half, to sub-pixel precision.
 *
 * An antialiased edge does not jump from 0 to 1; it passes through the
 * intervening values in one or two pixels, and where it crosses 0.5 is where
 * the geometric edge is. Linear interpolation between the two straddling
 * samples locates that to about a tenth of a pixel, which is why every length
 * in this experiment is quoted from crossings and not from run lengths in whole
 * pixels.
 */
function crossings(s: Sampled): { at: number; rising: boolean }[] {
  const out: { at: number; rising: boolean }[] = [];
  for (let i = 0; i + 1 < s.cover.length; i++) {
    const a = s.cover[i]!;
    const b = s.cover[i + 1]!;
    if (a < 0.5 && b >= 0.5) {
      out.push({ at: s.start + (i + (0.5 - a) / (b - a)) * s.step, rising: true });
    } else if (a >= 0.5 && b < 0.5) {
      out.push({ at: s.start + (i + (a - 0.5) / (a - b)) * s.step, rising: false });
    }
  }
  return out;
}

/** Alternating ink/gap lengths from a crossing list, in points. */
function runsOf(cross: { at: number; rising: boolean }[]): { on: number[]; off: number[] } {
  const on: number[] = [];
  const off: number[] = [];
  for (let i = 0; i + 1 < cross.length; i++) {
    const a = cross[i]!;
    const b = cross[i + 1]!;
    if (a.rising && !b.rising) on.push(b.at - a.at);
    else if (!a.rising && b.rising) off.push(b.at - a.at);
  }
  return { on, off };
}

/** Coverage as one byte per sample. */
function packCover(cover: readonly number[]): string {
  return cover
    .map((c) => Math.round(Math.max(0, Math.min(1, c)) * 255).toString(16).padStart(2, '0'))
    .join('');
}

const round = (v: number, places = 3): number => {
  const f = 10 ** places;
  return Math.round(v * f) / f;
};

/* -------------------------------------------------------------------------- */
/* per-probe measurement                                                      */
/* -------------------------------------------------------------------------- */

/** Read regions no longer than this keep their coverage; longer ones keep only crossings. */
const PROFILE_LIMIT = 800;

/** A `cols` region taller than this keeps only its boxes, not its column profile. */
const COLUMN_HEIGHT_LIMIT = 60;

interface Box {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** Which of the three inks a pixel is, or null if it is background. */
function hueOf(r: number, g: number, b: number): 'red' | 'green' | 'dark' | null {
  if (r > 240 && g > 240 && b > 240) return null;
  if (r > g + 40 && r > b + 40) return 'red';
  if (g > r + 40 && g > b + 40) return 'green';
  return 'dark';
}

function boxesOf(
  bmp: Bitmap,
  scale: number,
  region: { x0: number; y0: number; x1: number; y1: number },
): { ink: Box | null; red: Box | null; green: Box | null } {
  const acc: Record<string, Box | null> = { ink: null, red: null, green: null };
  const add = (key: string, px: number, py: number): void => {
    const b = acc[key];
    const x0 = px / scale;
    const y0 = py / scale;
    const x1 = (px + 1) / scale;
    const y1 = (py + 1) / scale;
    acc[key] =
      b == null
        ? { x0, y0, x1, y1 }
        : {
            x0: Math.min(b.x0, x0),
            y0: Math.min(b.y0, y0),
            x1: Math.max(b.x1, x1),
            y1: Math.max(b.y1, y1),
          };
  };
  for (let px = Math.max(0, Math.round(region.x0 * scale)); px <= Math.min(bmp.width - 1, Math.round(region.x1 * scale)); px++) {
    for (let py = Math.max(0, Math.round(region.y0 * scale)); py <= Math.min(bmp.height - 1, Math.round(region.y1 * scale)); py++) {
      const [r, g, b] = bmp.pixel(px, py);
      const hue = hueOf(r, g, b);
      if (hue === null) continue;
      add('ink', px, py);
      if (hue === 'red' || hue === 'green') add(hue, px, py);
    }
  }
  const tidy = (b: Box | null): Box | null =>
    b === null ? null : { x0: round(b.x0, 2), y0: round(b.y0, 2), x1: round(b.x1, 2), y1: round(b.y1, 2) };
  return { ink: tidy(acc['ink'] ?? null), red: tidy(acc['red'] ?? null), green: tidy(acc['green'] ?? null) };
}

interface Measured {
  id: string;
  deck: string;
  group: string;
  question: string;
  opened: boolean;
  repaired: boolean | null;
  error: string | null;
  rect: { x: number; y: number; cx: number; cy: number };
  read: ReadSpec;
  export: {
    file: string;
    pxPerPoint: number;
    samples: number;
    /**
     * Where sample zero's centre actually is, in points, and how far apart the
     * samples are.
     *
     * Stored rather than recomputed from the read window. A window that runs off
     * the slide is clamped to the bitmap, so its first sample is not where it
     * asked to start - and every section that recovered the origin by
     * arithmetic on `read.x0` was reading one probe's soft edge ten points to
     * the left of where it is.
     */
    start: number;
    step: number;
  } | null;
  /** Sub-pixel positions, in points, where coverage crosses one half. */
  edges: { at: number; rising: boolean }[];
  /**
   * The coverage integral over the read region, in points.
   *
   * Computed for every probe whether or not the profile is stored. The first
   * version of this file derived it from the stored string, so the widest probe
   * in the experiment - a stroke at the 1584pt maximum, whose region is longer
   * than the storage limit - reported zero ink and looked like a shape that had
   * not been painted at all.
   */
  total: number;
  /** One byte per sample, only when the region is short enough. */
  cover: string | null;
  /** Per-column ink height and extent, for the arrowhead probes. */
  columns: { ink: string; span: string; x0: number; scale: number } | null;
  /** Ink bounding boxes over a `cols` region, in points: all ink, and by hue. */
  boxes: {
    ink: Box | null;
    red: Box | null;
    green: Box | null;
  } | null;
  /** What the object model says, which is how a clamp is told from a refusal. */
  com: Partial<ComShape> | null;
}

const measured: Measured[] = [];

for (const probe of inputs.probes) {
  const record = deckOf.get(probe.deck);
  const opened = record?.opened ?? false;
  const com = record?.shapes.find((s) => s.id === probe.id) ?? null;
  const base: Measured = {
    id: probe.id,
    deck: probe.deck,
    group: probe.group,
    question: probe.question,
    opened,
    repaired: record?.repaired ?? null,
    error: record?.error ?? null,
    rect: { x: probe.x, y: probe.y, cx: probe.cx, cy: probe.cy },
    read: probe.read,
    export: null,
    edges: [],
    total: 0,
    cover: null,
    columns: null,
    boxes: null,
    com:
      com === null
        ? null
        : {
            weight: com.weight,
            dashStyle: com.dashStyle,
            lineStyle: com.lineStyle,
            insetPen: com.insetPen,
            endType: com.endType,
            endLen: com.endLen,
            endWidth: com.endWidth,
            shadowType: com.shadowType,
            shadowBlur: com.shadowBlur,
            shadowOffX: com.shadowOffX,
            shadowOffY: com.shadowOffY,
            glowRadius: com.glowRadius,
            softEdgeRad: com.softEdgeRad,
          },
  };

  const fine = finest(probe.deck);
  if (!opened || fine === null) {
    measured.push(base);
    continue;
  }

  const bmp = bitmap(fine.file);

  if (probe.read.mode === 'cols') {
    // Two-dimensional regions. Every one of them gets its ink bounding boxes,
    // separated by hue so a red shadow can be told from the black shape casting
    // it. Only the short ones - the arrowhead windows - also keep a per-column
    // profile: a column's ink height is stored as a byte at half-pixel
    // resolution, which tops out at 64 points.
    const view = coarse(probe.deck) ?? fine;
    const cb = bitmap(view.file);
    base.export = {
      file: view.file,
      pxPerPoint: view.scale,
      samples: 0,
      start: round(Math.max(0, Math.round(probe.read.x0 * view.scale)) / view.scale, 4),
      step: round(1 / view.scale, 6),
    };
    base.boxes = boxesOf(cb, view.scale, probe.read);

    const x0 = Math.max(0, Math.round(probe.read.x0 * view.scale));
    const x1 = Math.min(cb.width - 1, Math.round(probe.read.x1 * view.scale));
    const y0 = Math.max(0, Math.round(probe.read.y0 * view.scale));
    const y1 = Math.min(cb.height - 1, Math.round(probe.read.y1 * view.scale));
    let sum = 0;
    const inkCol: number[] = [];
    const spanCol: number[] = [];
    for (let x = x0; x <= x1; x++) {
      let column = 0;
      let top = -1;
      let bot = -1;
      for (let y = y0; y <= y1; y++) {
        const c = ink(cb, x, y);
        column += c;
        if (c >= 0.5) {
          if (top === -1) top = y;
          bot = y;
        }
      }
      sum += column;
      inkCol.push(column);
      spanCol.push(top === -1 ? 0 : bot - top + 1);
    }
    base.total = round(sum / (view.scale * view.scale), 3);
    base.export.samples = inkCol.length;
    if (probe.read.y1 - probe.read.y0 <= COLUMN_HEIGHT_LIMIT) {
      const half = (v: number): string =>
        Math.min(255, Math.round(v * 2))
          .toString(16)
          .padStart(2, '0');
      base.columns = {
        ink: inkCol.map(half).join(''),
        span: spanCol.map(half).join(''),
        x0: round(x0 / view.scale),
        scale: view.scale,
      };
    }
    measured.push(base);
    continue;
  }

  const s =
    probe.read.mode === 'row'
      ? sampleRow(bmp, fine.scale, probe.read.y, probe.read.x0, probe.read.x1, probe.id)
      : sampleCol(bmp, fine.scale, probe.read.x, probe.read.y0, probe.read.y1, probe.id);

  base.export = {
    file: fine.file,
    pxPerPoint: fine.scale,
    samples: s.cover.length,
    start: round(s.start, 4),
    step: round(s.step, 6),
  };
  base.edges = crossings(s).map((c) => ({ at: round(c.at), rising: c.rising }));
  base.total = round(s.cover.reduce((a, b) => a + b, 0) * s.step, 4);
  if (s.cover.length <= PROFILE_LIMIT) base.cover = packCover(s.cover);
  measured.push(base);
}

const byId = new Map(measured.map((m) => [m.id, m]));

/* -------------------------------------------------------------------------- */
/* reporting                                                                  */
/* -------------------------------------------------------------------------- */

function section(title: string, hypothesis: string): void {
  console.log(`\n${'='.repeat(78)}\n${title}\n${'-'.repeat(78)}\n${hypothesis}\n`);
}

const findings: Record<string, unknown> = {};

/* ---- A. the dash arrays -------------------------------------------------- */

section(
  'A. the eleven preset dash arrays',
  'Hypothesis: each preset is a fixed sequence of on/off lengths expressed as\n' +
    'multiples of the stroke width, and the widely quoted arrays are right.',
);

const DASH_W = 12;

/** On/off runs in width multiples, from the middle of a row where the pattern has settled. */
function dashArray(id: string, width = DASH_W): number[] | null {
  const m = byId.get(id);
  if (m === undefined || m.edges.length < 6) return null;
  const { on, off } = runsOf(m.edges);
  // Drop the first and last of each: the run at either end of the line is
  // clipped by the endpoint and is not a whole segment.
  const trim = (v: number[]): number[] => v.slice(1, -1);
  const onT = trim(on);
  const offT = trim(off);
  if (onT.length === 0 || offT.length === 0) return null;
  // The pattern repeats, so cluster each list into the distinct lengths present
  // rather than averaging a dash and a dot together.
  const cluster = (v: number[]): number[] => {
    const sorted = [...v].sort((a, b) => a - b);
    const groups: number[][] = [];
    for (const value of sorted) {
      const last = groups[groups.length - 1];
      if (last !== undefined && value - last[0]! < 0.25 * width) last.push(value);
      else groups.push([value]);
    }
    return groups.map((g) => g.reduce((a, b) => a + b, 0) / g.length / width);
  };
  return [...cluster(onT), NaN, ...cluster(offT)];
}

const dashTable: Record<string, { on: number[]; off: number[]; sequence: number[] }> = {};
for (const name of PRESET_DASHES) {
  const m = byId.get(`dash-${name}`);
  if (m === undefined) continue;
  const { on, off } = runsOf(m.edges);
  const trimmed = { on: on.slice(1, -1), off: off.slice(1, -1) };
  if (name === 'solid') {
    dashTable[name] = { on: [], off: [], sequence: [] };
    console.log(`  ${name.padEnd(14)} ${String(m.edges.length)} crossing(s) - one unbroken run`);
    continue;
  }
  // Recover the repeating sequence directly: walk the crossings and emit
  // alternating on/off lengths, then find the shortest repeating period.
  const seq: number[] = [];
  for (let i = 1; i + 1 < m.edges.length - 1; i++) {
    seq.push(round((m.edges[i + 1]!.at - m.edges[i]!.at) / DASH_W, 3));
  }
  const period = (values: number[]): number[] => {
    for (let p = 2; p <= 8; p += 2) {
      if (values.length < p * 3) continue;
      let ok = true;
      for (let i = 0; i + p < values.length; i++) {
        if (Math.abs(values[i]! - values[i + p]!) > 0.12) {
          ok = false;
          break;
        }
      }
      if (ok) {
        const avg: number[] = [];
        for (let k = 0; k < p; k++) {
          const picks = values.filter((_, i) => i % p === k);
          avg.push(round(picks.reduce((a, b) => a + b, 0) / picks.length, 3));
        }
        return avg;
      }
    }
    return values.slice(0, 8);
  };
  const sequence = period(seq);
  dashTable[name] = { on: trimmed.on.map((v) => round(v / DASH_W, 3)), off: trimmed.off.map((v) => round(v / DASH_W, 3)), sequence };
  console.log(
    `  ${name.padEnd(14)} period ${JSON.stringify(sequence)}  (x line width, ${String(m.edges.length)} crossings)`,
  );
}
findings['dashArrays'] = dashTable;

/* ---- the folklore check -------------------------------------------------- */

section(
  'A2. do the arrays everyone quotes paint the same picture?',
  'Each preset was also written as an explicit a:custDash of the array people\n' +
    'cite for it. Identical crossings mean the folklore is right; a different\n' +
    'period means it is not, and the preset row above is the answer.',
);

const folkloreVerdicts: Record<string, { same: boolean; preset: number[]; folklore: number[] }> = {};
for (const name of Object.keys(FOLKLORE_DASHES)) {
  const preset = byId.get(`dash-${name}`);
  const folk = byId.get(`folk-${name}`);
  if (preset === undefined || folk === undefined) continue;
  const seqOf = (m: Measured): number[] => {
    const seq: number[] = [];
    for (let i = 1; i + 2 < m.edges.length; i++) {
      seq.push(round((m.edges[i + 1]!.at - m.edges[i]!.at) / DASH_W, 2));
    }
    return seq.slice(0, 12);
  };
  const a = seqOf(preset);
  const b = seqOf(folk);
  const same = a.length === b.length && a.every((v, i) => Math.abs(v - b[i]!) < 0.12);
  folkloreVerdicts[name] = { same, preset: a.slice(0, 6), folklore: b.slice(0, 6) };
  console.log(
    `  ${name.padEnd(14)} ${same ? 'MATCHES' : 'DIFFERS'}  preset ${JSON.stringify(a.slice(0, 6))} vs custDash ${JSON.stringify(b.slice(0, 6))}`,
  );
}
findings['folklore'] = folkloreVerdicts;

/* ---- B. proportionality to @w -------------------------------------------- */

section(
  'B. is the array a multiple of the stroke width?',
  'The same `dash` preset at 3, 6 and 24 points. If the array is in width\n' +
    'multiples the periods differ by exactly those factors; if it is an absolute\n' +
    'length they are all the same.',
);

const widthScaling: Record<string, number[]> = {};
for (const w of [3, 6, 24]) {
  const m = byId.get(`dashw-${String(w)}`);
  if (m === undefined) continue;
  const seq: number[] = [];
  for (let i = 1; i + 2 < m.edges.length; i++) seq.push(m.edges[i + 1]!.at - m.edges[i]!.at);
  const on = seq.filter((_, i) => i % 2 === 0);
  const off = seq.filter((_, i) => i % 2 === 1);
  const mean = (v: number[]): number => round(v.reduce((a, b) => a + b, 0) / v.length, 3);
  widthScaling[String(w)] = [mean(on), mean(off), round(mean(on) / w, 3), round(mean(off) / w, 3)];
  console.log(
    `  w=${String(w).padStart(2)}pt  on ${String(mean(on)).padStart(7)}pt  off ${String(mean(off)).padStart(7)}pt` +
      `   -> ${String(round(mean(on) / w, 2))} x w on, ${String(round(mean(off) / w, 2))} x w off`,
  );
}
findings['widthScaling'] = widthScaling;

/* ---- C. caps ------------------------------------------------------------- */

section(
  'C. what does @cap do, and what is the default?',
  'ECMA-376 gives ST_LineEndCap a default of "sq". Office is widely said to use\n' +
    'flat instead. A 200pt line at 24pt wide, drawn from x=300 to x=500: a flat\n' +
    'cap stops at 500, a square one reaches 512.',
);

const capExtents: Record<string, { left: number; right: number; overshoot: number }> = {};
for (const cap of ['absent', 'flat', 'sq', 'rnd']) {
  const m = byId.get(`cap-${cap}`);
  if (m === undefined || m.edges.length < 2) continue;
  const left = m.edges[0]!.at;
  const right = m.edges[m.edges.length - 1]!.at;
  const overshoot = round(right - 500, 3);
  capExtents[cap] = { left: round(left), right: round(right), overshoot };
  console.log(
    `  cap=${cap.padEnd(7)} ink spans ${String(round(left, 2))} .. ${String(round(right, 2))}pt` +
      `   overshoot ${String(overshoot)}pt = ${String(round(overshoot / 24, 3))} x w`,
  );
}
findings['capExtents'] = capExtents;

/* ---- C2. cap and dash together ------------------------------------------- */

section(
  'C2. does the cap lengthen each dash, or is the array compensated?',
  'PowerPoint distinguishes square-dot from round-dot by @cap alone, so the cap\n' +
    'must change what a dash looks like. Either each segment grows by one width\n' +
    '(half at each end) and each gap shrinks to match, or the array is adjusted\n' +
    'so the painted segment stays the same length.',
);

const capDash: Record<string, { on: number; off: number; period: number }> = {};
for (const name of ['dot', 'dash']) {
  for (const cap of ['flat', 'sq', 'rnd']) {
    const m = byId.get(`dashcap-${name}-${cap}`);
    if (m === undefined || m.edges.length < 6) continue;
    const { on, off } = runsOf(m.edges);
    const mean = (v: number[]): number =>
      v.length === 0 ? 0 : v.slice(1, -1).reduce((a, b) => a + b, 0) / Math.max(1, v.length - 2);
    const onW = round(mean(on) / DASH_W, 3);
    const offW = round(mean(off) / DASH_W, 3);
    capDash[`${name}-${cap}`] = { on: onW, off: offW, period: round(onW + offW, 3) };
    console.log(
      `  ${name.padEnd(5)} cap=${cap.padEnd(5)}  on ${String(onW).padStart(6)} x w   off ${String(offW).padStart(6)} x w   period ${String(round(onW + offW, 3))}`,
    );
  }
}
findings['capDash'] = capDash;

/* ---- D. widths ----------------------------------------------------------- */

section(
  'D. how thick is a stroke with no @w, or with w="0"?',
  'A column across each line. The stroke thickness is the distance between the\n' +
    'two crossings; a sub-pixel stroke has no crossings at all and is measured by\n' +
    'its total coverage instead.',
);

const widths: Record<string, { thickness: number; coverage: number; comWeight: number | null }> = {};
for (const m of measured.filter((p) => p.group === 'width')) {
  const thickness =
    m.edges.length >= 2 ? round(m.edges[m.edges.length - 1]!.at - m.edges[0]!.at, 3) : 0;
  widths[m.id] = { thickness, coverage: m.total, comWeight: m.com?.weight ?? null };
  console.log(
    `  ${m.id.padEnd(12)} between crossings ${String(thickness).padStart(8)}pt   total ink ${String(m.total).padStart(8)}pt` +
      `   COM Weight=${String(m.com?.weight ?? '-')}`,
  );
}
findings['widths'] = widths;

function unpack(hex: string): number[] {
  const out: number[] = [];
  for (let i = 0; i + 1 < hex.length; i += 2) out.push(parseInt(hex.slice(i, i + 2), 16) / 255);
  return out;
}

/* ---- E. alignment -------------------------------------------------------- */

section(
  'E. where does algn="in" put the stroke?',
  'A 20pt stroke on a rectangle whose left edge is at a known whole point.\n' +
    'Centred, the ink straddles the edge by 10pt either side. Inset, it should\n' +
    'sit wholly inside - which is what SVG cannot express and what the plan\n' +
    'proposes to emulate by clipping and double-stroking.',
);

const alignment: Record<string, { edgeAt: number; inkFrom: number; inkTo: number; outside: number }> = {};
for (const m of measured.filter((p) => p.group === 'algn')) {
  if (m.edges.length < 2) continue;
  const edge = m.rect.x;
  const from = m.edges[0]!.at;
  const to = m.edges[1]!.at;
  alignment[m.id] = {
    edgeAt: edge,
    inkFrom: round(from, 3),
    inkTo: round(to, 3),
    outside: round(edge - from, 3),
  };
  console.log(
    `  ${m.id.padEnd(16)} edge at ${String(edge)}pt, ink ${String(round(from, 2))}..${String(round(to, 2))}` +
      `   outside the edge: ${String(round(edge - from, 2))}pt   COM InsetPen=${String(m.com?.insetPen ?? '-')}`,
  );
}
findings['alignment'] = alignment;

/* ---- F. compound --------------------------------------------------------- */

section(
  'F. how does @cmpd divide the stroke?',
  'A 36pt stroke on a rectangle, read across its left edge. The crossings give\n' +
    'each rail and each gap; the question is what fractions of @w they are, and\n' +
    'whether the total still comes to @w.',
);

const CMPD_W = 36;
const compound: Record<
  string,
  { at: number[]; rails: number[]; gaps: number[]; ink: number; span: number; firstAt: number }
> = {};
for (const m of measured.filter((p) => p.group === 'cmpd')) {
  if (m.edges.length < 2) continue;
  // The rectangle's own left edge, which is where a centred stroke of width w
  // would run from -w/2 to +w/2. Everything below is quoted relative to it.
  const edge = m.rect.x;
  const at = m.edges.map((e) => round((e.at - edge) / CMPD_W, 4));
  const { on, off } = runsOf(m.edges);
  const span = m.edges[m.edges.length - 1]!.at - m.edges[0]!.at;
  compound[m.id] = {
    at,
    rails: on.map((v) => round(v / CMPD_W, 4)),
    gaps: off.map((v) => round(v / CMPD_W, 4)),
    ink: round(on.reduce((a, b) => a + b, 0) / CMPD_W, 4),
    span: round(span / CMPD_W, 4),
    firstAt: round((m.edges[0]!.at - edge) / CMPD_W, 4),
  };
  console.log(
    `  ${m.id.padEnd(16)} crossings at ${JSON.stringify(at)} x w from the edge` +
      `\n  ${' '.repeat(16)} rails ${JSON.stringify(on.map((v) => round(v / CMPD_W, 3)))} gaps ${JSON.stringify(off.map((v) => round(v / CMPD_W, 3)))}` +
      ` -> ${String(round(on.reduce((a, b) => a + b, 0) / CMPD_W, 3))} x w of ink across ${String(round(span / CMPD_W, 3))} x w`,
  );
}
findings['compound'] = compound;

/* ---- G. joins ------------------------------------------------------------ */

section(
  'G. which join does Office use when none is written, and where does a miter stop?',
  'A V with a shallow apex, stroked at 20pt, read as a column through the apex.\n' +
    'A miter spikes past the geometric corner; a bevel and a round join do not.',
);

const JOIN_W = 20;
const joins: Record<string, { reach: number; apexY: number; inWidths: number }> = {};
for (const m of measured.filter((p) => p.group === 'join' && !p.id.startsWith('miter-r'))) {
  if (m.edges.length === 0) continue;
  const deepest = m.edges[m.edges.length - 1]!.at;
  // The path's apex sits at the bottom of the shape rectangle.
  const apex = m.rect.y + m.rect.cy;
  joins[m.id] = {
    reach: round(deepest - apex, 3),
    apexY: apex,
    inWidths: round((deepest - apex) / JOIN_W, 3),
  };
  console.log(
    `  ${m.id.padEnd(16)} ink reaches ${String(round(deepest, 2))}pt, apex at ${String(apex)}pt` +
      `   -> ${String(round(deepest - apex, 2))}pt past the corner = ${String(round((deepest - apex) / JOIN_W, 2))} x w`,
  );
}
findings['joins'] = joins;

/* ---- G2. the default miter limit ----------------------------------------- */

section(
  'G2. how far will a miter with no @lim reach?',
  'The V above has an 80-degree apex, whose full miter is only 1.56 stroke\n' +
    'widths - enough to prove the default limit is at least that and no more.\n' +
    'These five corners have miter ratios of 2, 4, 8, 12 and 20. Full miters all\n' +
    'the way up means a very large default; a plateau is the default itself.',
);

const miterLimits: Record<string, { ratio: number; reach: number; full: number; clipped: boolean }> = {};
for (const m of measured
  .filter((p) => p.id.startsWith('miter-r'))
  .sort((a, b) => Number(a.id.replace('miter-r', '')) - Number(b.id.replace('miter-r', '')))) {
  const ratio = Number(m.id.replace('miter-r', ''));
  if (m.edges.length === 0) {
    console.log(`  ratio ${String(ratio).padStart(2)}  no ink in the window`);
    continue;
  }
  const apex = m.rect.y + m.rect.cy;
  const deepest = m.edges[m.edges.length - 1]!.at;
  const reach = round((deepest - apex) / JOIN_W, 3);
  // A full miter tip stands (w/2) * ratio past the corner; a bevel stands
  // (w/2) * cos(half-angle) past it, which for a sharp corner is nearly zero.
  const full = round(ratio / 2, 3);
  miterLimits[m.id] = { ratio, reach, full, clipped: reach < full - 0.15 };
  console.log(
    `  ratio ${String(ratio).padStart(2)}  reach ${String(reach).padStart(7)} x w` +
      `   a full miter would be ${String(full).padStart(6)} x w   ${reach < full - 0.15 ? 'CLIPPED' : 'full'}`,
  );
}
findings['miterLimits'] = miterLimits;

/* ---- G3. what unit is @lim in? ------------------------------------------- */

section(
  'G3. is a:miter/@lim a ratio of the stroke width, or of half of it?',
  'Four explicit limits on a corner whose miter ratio is 5. In whole widths the\n' +
    'change is between 400% and 600%; in half widths it is between 800% and 1200%.',
);

const miterUnits: Record<string, { limitPct: number; reach: number; clipped: boolean }> = {};
for (const pct of [200, 400, 600, 1000]) {
  const m = byId.get(`mlim-${String(pct)}`);
  if (m === undefined || m.edges.length === 0) {
    console.log(`  lim=${String(pct).padStart(4)}%  no ink in the window`);
    continue;
  }
  const apex = m.rect.y + m.rect.cy;
  const reach = round((m.edges[m.edges.length - 1]!.at - apex) / JOIN_W, 3);
  // A full miter on this corner stands 2.5 stroke widths past it.
  const clipped = reach < 2.35;
  miterUnits[m.id] = { limitPct: pct, reach, clipped };
  console.log(
    `  lim=${String(pct).padStart(4)}%  reach ${String(reach).padStart(6)} x w   ${clipped ? 'CLIPPED' : 'full (2.5 x w)'}`,
  );
}
findings['miterUnits'] = miterUnits;

/* ---- H. arrowheads ------------------------------------------------------- */

section(
  'H. the six arrowheads, at three widths and three lengths',
  'Each head is recovered from the bitmap as a per-column ink profile around the\n' +
    'line endpoint - which is the marker outline, sampled. That is also how this\n' +
    "sub-phase avoids transcribing LibreOffice's vertex tables, which are MPL-2.0\n" +
    'and would put file-level copyleft into an Apache-2.0 tree.',
);

const AH_W = 8;

/**
 * Candidate outlines, as the half-height a marker has at a fractional distance
 * `u` back from its tip. Each is normalised: `u` runs 0 at the tip to 1 at the
 * back, and the value returned is a fraction of the marker's half-width.
 *
 * Thresholding the profile is not enough on its own and the first run of this
 * section proved it. A triangle's span reaches the shaft's height almost at its
 * base, so a threshold finds the base; an ellipse's does not, so a threshold
 * finds a point well inside it and reports every oval about nine percent short.
 * Fitting the family recovers the real extent, and the residual says whether the
 * family was right.
 */
const OUTLINES: Readonly<Record<string, (u: number) => number>> = {
  triangle: (u) => u,
  ellipse: (u) => Math.sqrt(Math.max(0, 1 - (2 * u - 1) ** 2)),
  diamond: (u) => (u <= 0.5 ? 2 * u : 2 - 2 * u),
};

interface HeadFit {
  outline: string;
  length: number;
  width: number;
  rms: number;
  /**
   * Ink divided by span over the head.
   *
   * The span profile cannot tell a stealth from a triangle - the notch in a
   * stealth's back is interior, so its silhouette is a triangle's - and it
   * cannot tell an open arrow from a filled one either. How much of the
   * silhouette is actually inked can tell both.
   */
  fillRatio: number;
  tipAt: number;
}

/**
 * Fit each family to the measured span profile and keep the best.
 *
 * `length` and `width` come out in stroke widths, which is the only unit the
 * answer can be quoted in - the same head at a different `@w` is a different
 * number of points and the same number of widths.
 */
function fitHead(m: Measured): HeadFit | null {
  if (m.columns == null) return null;
  const inkP = unpackHalf(m.columns.ink);
  const spanP = unpackHalf(m.columns.span);
  const scale = m.columns.scale;
  const shaftPx = AH_W * scale;
  let tip = -1;
  let maxSpan = 0;
  for (let i = 0; i < spanP.length; i++) {
    if (inkP[i]! > 0.25) tip = i;
    maxSpan = Math.max(maxSpan, spanP[i]!);
  }
  if (tip === -1 || maxSpan <= shaftPx * 1.05) {
    return { outline: 'none', length: 0, width: 0, rms: 0, fillRatio: 1, tipAt: 0 };
  }
  // One fixed window for every candidate, from the tip back to the longest
  // length any candidate could have. The first version of this fit scored only
  // the columns inside the candidate's own length and divided by their count,
  // so a longer candidate swept in more of the flat shaft - which every model
  // fits perfectly - and drove its own error down. Every narrow head came back
  // about seventy percent too long.
  const window = Math.min(tip + 1, Math.ceil(shaftPx * 9));
  let best: HeadFit | null = null;
  for (const [name, shape] of Object.entries(OUTLINES)) {
    for (let lenPx = shaftPx * 0.5; lenPx <= shaftPx * 8; lenPx += 0.5) {
      let sum = 0;
      for (let k = 0; k < window; k++) {
        const i = tip - k;
        const back = k + 1; // pixels behind the tip
        const u = back / lenPx;
        // The painted span is the greater of the marker's own outline and the
        // shaft it sits on, because the two overlap. Behind the marker there is
        // only shaft.
        const model = Math.max(u > 1 ? 0 : shape(u) * maxSpan, shaftPx);
        const d = model - spanP[i]!;
        sum += d * d;
      }
      const rms = Math.sqrt(sum / window) / scale;
      if (best === null || rms < best.rms) {
        best = {
          outline: name,
          length: round(lenPx / scale / AH_W, 3),
          width: round(maxSpan / scale / AH_W, 3),
          rms: round(rms, 3),
          fillRatio: 1,
          tipAt: round(m.columns.x0 + (tip + 1) / scale, 2),
        };
      }
    }
  }
  if (best === null) return null;
  const from = Math.max(0, tip - Math.round(best.length * AH_W * scale));
  const inkSum = inkP.slice(from, tip + 1).reduce((a, b) => a + b, 0);
  const spanSum = spanP.slice(from, tip + 1).reduce((a, b) => a + b, 0);
  best.fillRatio = spanSum > 0 ? round(inkSum / spanSum, 3) : 1;
  return best;
}

const heads: Record<string, HeadFit> = {};
for (const type of ARROW_TYPES) {
  for (const len of ARROW_SIZES) {
    for (const wid of ARROW_SIZES) {
      const m = byId.get(`ah-${type}-l${len}-w${wid}`);
      if (m === undefined) continue;
      const fit = fitHead(m);
      if (fit !== null) heads[`${type}-l${len}-w${wid}`] = fit;
    }
  }
  const row = ARROW_SIZES.flatMap((len) =>
    ARROW_SIZES.map((wid) => {
      const h = heads[`${type}-l${len}-w${wid}`];
      return h === undefined ? '  -   ' : `${h.length.toFixed(1)}x${h.width.toFixed(1)}`;
    }),
  );
  const shapes = new Set(
    ARROW_SIZES.flatMap((len) =>
      ARROW_SIZES.map((wid) => heads[`${type}-l${len}-w${wid}`]?.outline ?? '?'),
    ),
  );
  const worst = Math.max(
    0,
    ...ARROW_SIZES.flatMap((len) =>
      ARROW_SIZES.map((wid) => heads[`${type}-l${len}-w${wid}`]?.rms ?? 0),
    ),
  );
  const ratios = ARROW_SIZES.flatMap((len) =>
    ARROW_SIZES.map((wid) => heads[`${type}-l${len}-w${wid}`]?.fillRatio ?? 1),
  );
  const meanRatio = ratios.reduce((a, b) => a + b, 0) / ratios.length;
  console.log(
    `  ${type.padEnd(9)} ${row.join(' ')}  | outline ${[...shapes].join('/')}, ink/silhouette ${meanRatio.toFixed(2)}, worst rms ${String(round(worst, 2))}pt`,
  );
}
findings['arrowheads'] = heads;

function unpackHalf(hex: string): number[] {
  const out: number[] = [];
  for (let i = 0; i + 1 < hex.length; i += 2) out.push(parseInt(hex.slice(i, i + 2), 16) / 2);
  return out;
}

{
  const shaft = byId.get('ah-shaft');
  if (shaft !== undefined && shaft.edges.length >= 2) {
    console.log(
      `\n  ah-shaft: ink from ${String(round(shaft.edges[0]!.at, 2))} to ${String(round(shaft.edges[shaft.edges.length - 1]!.at, 2))}pt` +
        ` against a line that ends at 700pt`,
    );
  }
  const both = byId.get('ah-both');
  if (both !== undefined && both.edges.length >= 2) {
    console.log(
      `  ah-both:  ink from ${String(round(both.edges[0]!.at, 2))} to ${String(round(both.edges[both.edges.length - 1]!.at, 2))}pt` +
        ` against a line from 300 to 700pt (oval head at the start, triangle at the end)`,
    );
  }
}

/* ---- I. shadow direction ------------------------------------------------- */

section(
  'I. which way is dir=0, and which way does it turn?',
  'A 100pt black square with a red shadow at dist=30pt and no blur, at eight\n' +
    'directions. The shadow rectangle is found in the bitmap and its offset from\n' +
    'the shape is the answer. If dir=0 is to the right and the angle increases\n' +
    'clockwise in screen space, dir=90 lands below.',
);

const shadowDirs: Record<
  string,
  { visible: Box | null; dx: number | null; dy: number | null; comX: number | null; comY: number | null }
> = {};
for (const deg of [0, 45, 90, 135, 180, 225, 270, 315]) {
  const m = byId.get(`shdw-dir${String(deg)}`);
  if (m?.boxes == null) continue;
  const red = m.boxes.red;
  // Only the part of the shadow that is not behind the opaque shape is red, so
  // the visible box is a *corner* of the shadow rather than the whole of it.
  // Which corner it is tells the direction, and the offset is then read from
  // whichever pair of sides sticks out.
  const dx =
    red === null
      ? null
      : red.x1 > m.rect.x + m.rect.cx
        ? round(red.x1 - (m.rect.x + m.rect.cx), 2)
        : red.x0 < m.rect.x
          ? round(red.x0 - m.rect.x, 2)
          : 0;
  const dy =
    red === null
      ? null
      : red.y1 > m.rect.y + m.rect.cy
        ? round(red.y1 - (m.rect.y + m.rect.cy), 2)
        : red.y0 < m.rect.y
          ? round(red.y0 - m.rect.y, 2)
          : 0;
  shadowDirs[String(deg)] = {
    visible: red,
    dx,
    dy,
    comX: m.com?.shadowOffX ?? null,
    comY: m.com?.shadowOffY ?? null,
  };
  console.log(
    `  dir=${String(deg).padStart(3)}deg  offset from the bitmap (${String(dx).padStart(7)}, ${String(dy).padStart(7)})pt` +
      `   COM (${String(round(m.com?.shadowOffX ?? 0, 2)).padStart(7)}, ${String(round(m.com?.shadowOffY ?? 0, 2)).padStart(7)})`,
  );
}
findings['shadowDirs'] = shadowDirs;

/* ---- J. the blur --------------------------------------------------------- */

section(
  'J. what is blurRad, in the units an SVG filter wants?',
  'feGaussianBlur takes a standard deviation. a:outerShdw takes a "blur radius"\n' +
    'and does not say what radius means. A step edge blurred by a Gaussian has an\n' +
    'erf profile, so fitting one to the measured edge gives sigma directly - and\n' +
    'the residual says whether it is a Gaussian at all.',
);

/** Abramowitz & Stegun 7.1.26. Good to 1.5e-7, which is far below a byte. */
function erf(x: number): number {
  const sign = x < 0 ? -1 : 1;
  const t = 1 / (1 + 0.3275911 * Math.abs(x));
  const y =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-x * x);
  return sign * y;
}

interface BlurFit {
  sigma: number;
  edge: number;
  amplitude: number;
  rms: number;
  worst: number;
}

/**
 * Least squares over sigma and the edge position; amplitude from the plateau.
 *
 * Only one edge at a time. A window that holds both sides of a blurred slab is
 * not a step and no single erf fits it - the first run of this section put a
 * whole 60pt shadow inside one column and the fit ran away to the top of the
 * sweep, which is what a runaway fit looks like when the model is wrong rather
 * than the data.
 */
function fitBlur(cover: readonly number[], start: number, step: number, falling: boolean): BlurFit {
  const amp = Math.max(...cover);
  let best: BlurFit = { sigma: 0, edge: 0, amplitude: amp, rms: Infinity, worst: Infinity };
  const span = (cover.length - 1) * step;
  for (let sigma = 0.02; sigma <= span / 2; sigma += 0.02) {
    for (let e = start; e <= start + span; e += step / 4) {
      let sum = 0;
      let worst = 0;
      for (let i = 0; i < cover.length; i++) {
        const x = start + i * step;
        const z = (x - e) / (sigma * Math.SQRT2);
        const model = amp * 0.5 * (falling ? 1 - erf(z) : 1 + erf(z));
        const d = model - cover[i]!;
        sum += d * d;
        worst = Math.max(worst, Math.abs(d));
      }
      const rms = Math.sqrt(sum / cover.length);
      if (rms < best.rms) best = { sigma: round(sigma, 4), edge: round(e, 3), amplitude: round(amp, 4), rms: round(rms, 5), worst: round(worst, 5) };
    }
  }
  return best;
}

const blurFits: Record<string, BlurFit & { blurRad: number; ratio: number }> = {};
for (const id of ['shdw-blur0', 'shdw-blur4', 'shdw-blur8', 'shdw-blur16', 'shdw-blur24']) {
  const m = byId.get(id);
  if (m?.cover == null || m.export === null) continue;
  const rad = Number(id.replace('shdw-blur', ''));
  const cover = unpack(m.cover);
  const read = m.read;
  if (read.mode !== 'row') continue;
  const start = m.export.start;
  const fit = fitBlur(cover, start, 1 / m.export.pxPerPoint, true);
  blurFits[id] = { ...fit, blurRad: rad, ratio: rad === 0 ? 0 : round(fit.sigma / rad, 4) };
  console.log(
    `  blurRad=${String(rad).padStart(2)}pt   sigma ${String(fit.sigma).padStart(7)}pt   sigma/blurRad ${String(rad === 0 ? '-' : round(fit.sigma / rad, 4)).padStart(7)}` +
      `   rms ${String(round(fit.rms * 255, 2)).padStart(5)}/255   worst ${String(round(fit.worst * 255, 2)).padStart(5)}/255`,
  );
}
for (const id of ['shdw-vblur0', 'shdw-vblur8', 'shdw-vblur24']) {
  const m = byId.get(id);
  if (m?.cover == null || m.export === null) continue;
  const rad = Number(id.replace('shdw-vblur', ''));
  const cover = unpack(m.cover);
  const read = m.read;
  if (read.mode !== 'col') continue;
  const start = m.export.start;
  const fit = fitBlur(cover, start, 1 / m.export.pxPerPoint, false);
  blurFits[id] = { ...fit, blurRad: rad, ratio: rad === 0 ? 0 : round(fit.sigma / rad, 4) };
  console.log(
    `  vertical blurRad=${String(rad).padStart(2)}pt   sigma ${String(fit.sigma).padStart(7)}pt   ratio ${String(rad === 0 ? '-' : round(fit.sigma / rad, 4)).padStart(7)}` +
      `   worst ${String(round(fit.worst * 255, 2)).padStart(5)}/255`,
  );
}
// Is the blur the same in both directions? Fitting two sigmas and comparing
// them tests the model as much as the data. Comparing the two ramps byte for
// byte does not: if a horizontal edge and a vertical one produce the same
// sequence of bytes, the blur is isotropic whatever its shape.
const isotropy: Record<string, { samples: number; worst: number }> = {};
for (const rad of [8, 24]) {
  const h = byId.get(`shdw-blur${String(rad)}`);
  const v = byId.get(`shdw-vblur${String(rad)}`);
  if (h?.cover == null || v?.cover == null) continue;
  const ramp = (cover: number[], reverse: boolean): number[] => {
    const seq = reverse ? [...cover].reverse() : cover;
    const from = seq.findIndex((c) => c > 0.01);
    const to = seq.findIndex((c) => c > 0.99);
    return from === -1 || to === -1 ? [] : seq.slice(from, to + 1);
  };
  const a = ramp(unpack(h.cover), true);
  const b = ramp(unpack(v.cover), false);
  const n = Math.min(a.length, b.length);
  let worst = 0;
  for (let i = 0; i < n; i++) worst = Math.max(worst, Math.abs(a[i]! - b[i]!) * 255);
  isotropy[String(rad)] = { samples: n, worst: round(worst, 2) };
  console.log(
    `  blurRad=${String(rad).padStart(2)}pt  the horizontal and vertical ramps agree to ${String(round(worst, 2))}/255 over ${String(n)} samples` +
      ` (${String(a.length)} vs ${String(b.length)} long)`,
  );
}
findings['blurIsotropy'] = isotropy;
findings['blurFits'] = blurFits;

/* ---- K. shadow affine ---------------------------------------------------- */

section(
  'K. sx, sy, kx and algn - the shadow as an affine transform',
  'Each of these is a red shadow of a black 120x100pt block. The measurement is\n' +
    'the red bounding box, which says what the transform did and about what\n' +
    'origin it did it.',
);

const affineBoxes: Record<
  string,
  { box: Box | null; dx0: number; dy0: number; dw: number; dh: number }
> = {};
{
  const controlOf = (deck: string): Measured | undefined =>
    measured.find((p) => p.deck === deck && p.id.startsWith('shdw-plain'));
  for (const m of measured.filter((p) => p.deck.startsWith('shaffine'))) {
    const control = controlOf(m.deck);
    const base = control?.boxes?.red ?? null;
    const red = m.boxes?.red ?? null;
    if (red === null || base === null || control === undefined) {
      console.log(`  ${m.id.padEnd(18)} no shadow visible`);
      continue;
    }
    // Everything is quoted against the untransformed control, which is the
    // shape's own box thrown 200pt right - so a zero row is a transform that
    // did nothing and every other row says what it moved and by how much.
    const rel = {
      dx0: round(red.x0 - base.x0 - (m.rect.x - control.rect.x), 2),
      dy0: round(red.y0 - base.y0 - (m.rect.y - control.rect.y), 2),
      dw: round(red.x1 - red.x0 - (base.x1 - base.x0), 2),
      dh: round(red.y1 - red.y0 - (base.y1 - base.y0), 2),
    };
    affineBoxes[m.id] = { box: red, ...rel };
    console.log(
      `  ${m.id.padEnd(18)} box ${String(round(red.x1 - red.x0, 1)).padStart(6)} x ${String(round(red.y1 - red.y0, 1)).padStart(6)}pt` +
        `   vs the control: top-left moved (${String(rel.dx0).padStart(7)}, ${String(rel.dy0).padStart(7)}), size changed (${String(rel.dw).padStart(6)}, ${String(rel.dh).padStart(6)})`,
    );
  }
}
findings['shadowAffine'] = affineBoxes;

/* ---- L. glow, soft edge, blur, inner shadow ------------------------------ */

section(
  'L. glow, soft edge, a:blur and the inner shadow',
  'Each read across an edge. A glow reaches outward from the shape; a soft edge\n' +
    'fades the shape itself; a:blur blurs everything; an inner shadow darkens\n' +
    'inward. The numbers are how far, and with what falloff.',
);

interface EffectReach {
  /** How far outside the shape's own edge any ink appears, in points. */
  outward: number;
  /** How far inside it the coverage is still below full, in points. */
  inward: number;
  /** Coverage exactly at the geometric edge. */
  atEdge: number;
  /**
   * The fitted step edge as an offset from the shape's geometry, in points.
   *
   * Positive means the effect's edge sits outside the shape - a glow that has
   * been dilated before it was blurred. This is the number that separates
   * "blurred by sigma" from "grown by d and then blurred by sigma", and the
   * two look identical in a reach measurement.
   */
  fitOffset: number | null;
  fit: BlurFit | null;
}

const reaches: Record<string, EffectReach> = {};
for (const m of measured.filter((p) =>
  ['glow', 'softEdge', 'blur', 'innerShdw'].includes(p.group),
)) {
  if (m.cover === null || m.export === null || m.read.mode !== 'row') continue;
  const cover = unpack(m.cover);
  const step = m.export.step;
  const start = m.export.start;
  const at = (points: number): number => {
    const i = Math.round((points - start) / step);
    return cover[Math.max(0, Math.min(cover.length - 1, i))] ?? 0;
  };
  // Outward reach: the first sample with any ink at all. Inward reach: how far
  // past the edge coverage is still climbing, which is the only thing a soft
  // edge or an inner shadow does and which an outward-only metric reports as a
  // negative number, as the first run of this section did.
  let first = -1;
  for (let i = 0; i < cover.length; i++) {
    if (cover[i]! > 0.02) {
      first = i;
      break;
    }
  }
  const peak = Math.max(...cover);
  let settled = -1;
  for (let i = 0; i < cover.length; i++) {
    const x = start + i * step;
    if (x >= m.rect.x && cover[i]! >= peak - 0.02) {
      settled = i;
      break;
    }
  }
  const outward = first === -1 ? 0 : m.rect.x - (start + first * step);
  const inward = settled === -1 ? 0 : start + settled * step - m.rect.x;
  // Every one of these is a step edge with something done to it, so every one
  // of them gets the same erf fit. An inner shadow is the exception - it is
  // darkest at the edge and fades inward, which is a ridge and not a step.
  const fit = m.group === 'innerShdw' ? null : fitBlur(cover, start, step, false);
  reaches[m.id] = {
    outward: round(outward, 3),
    inward: round(inward, 3),
    atEdge: round(at(m.rect.x), 4),
    fitOffset: fit === null ? null : round(m.rect.x - fit.edge, 3),
    fit,
  };
  console.log(
    `  ${m.id.padEnd(12)} ink from ${String(round(outward, 2)).padStart(6)}pt outside;` +
      ` full ${String(round(inward, 2)).padStart(6)}pt inside; at the edge ${String(round(at(m.rect.x), 3)).padStart(5)}` +
      (fit === null
        ? ''
        : `; fitted edge ${String(round(m.rect.x - fit.edge, 2)).padStart(6)}pt outside, sigma ${String(fit.sigma).padStart(5)}pt, worst ${String(round(fit.worst * 255, 1))}/255`),
  );
}
findings['effectReach'] = reaches;

// An inner shadow is a ridge, not a step: darkest at the edge, fading inward.
// The measurement is how far in it fades and how dark it gets, which the reach
// metric above reports as a negative outward distance and nothing else.
const innerProfiles: Record<string, { peak: number; peakAt: number; extent: number }> = {};
for (const m of measured.filter((p) => p.group === 'innerShdw')) {
  if (m.cover === null || m.export === null || m.read.mode !== 'row') continue;
  const cover = unpack(m.cover);
  const step = m.export.step;
  const start = m.export.start;
  let peak = 0;
  let peakAt = 0;
  for (let i = 0; i < cover.length; i++) {
    if (cover[i]! > peak) {
      peak = cover[i]!;
      peakAt = start + i * step;
    }
  }
  // From the peak inward, where does it fall below a fortieth of full?
  let extent = 0;
  for (let i = Math.round((peakAt - start) / step); i < cover.length; i++) {
    if (cover[i]! < peak * 0.025) {
      extent = start + i * step - peakAt;
      break;
    }
  }
  innerProfiles[m.id] = {
    peak: round(peak, 4),
    peakAt: round(peakAt - m.rect.x, 3),
    extent: round(extent, 3),
  };
  console.log(
    `  ${m.id.padEnd(12)} darkest ${String(round(peak, 3))} at ${String(round(peakAt - m.rect.x, 2))}pt inside the edge,` +
      ` faded out ${String(round(extent, 2))}pt further in`,
  );
}
findings['innerShadow'] = innerProfiles;

// Which effect is painted on top? Both are present, so the question is settled
// only where they overlap: a 16pt green glow reaches 16pt out on every side, and
// a red shadow thrown 30pt right passes through that band on the shape's right.
{
  const stack = byId.get('stack-glow-shadow');
  const view = stack === undefined ? null : coarse(stack.deck);
  const r = stack?.read;
  if (stack !== undefined && view !== null && r !== undefined && r.mode === 'cols') {
    const bmp = bitmap(view.file);
    // A band just outside the shape's right edge, vertically in its middle:
    // inside the glow's reach, and inside the shadow's.
    const y = stack.rect.y + stack.rect.cy / 2;
    const row: string[] = [];
    for (let px = Math.round((stack.rect.x + stack.rect.cx) * view.scale); px <= Math.round((stack.rect.x + stack.rect.cx + 24) * view.scale); px++) {
      const [rr, gg, bb] = bmp.pixel(px, Math.floor(y * view.scale));
      row.push(hueOf(rr, gg, bb) ?? 'bg');
    }
    const counts: Record<string, number> = {};
    for (const hue of row) counts[hue] = (counts[hue] ?? 0) + 1;
    console.log(
      `\n  stack-glow-shadow: across the 24pt just right of the shape, where a 16pt glow and a 30pt shadow both fall:` +
        `\n    ${Object.entries(counts).map(([k, v]) => `${k} ${String(v)}`).join(', ')}` +
        `\n    first 12 samples: ${row.slice(0, 12).join(' ')}`,
    );
    findings['stacking'] = { counts, sequence: row.slice(0, 24) };
  }
}

/* ---- L2. the twenty preset shadows --------------------------------------- */

section(
  'L2. a:prstShdw, which nothing documents and PowerPoint still writes',
  'Twenty presets, each a red shadow of a black 70x60pt block thrown 60pt right.\n' +
    'The measurement is the shadow box against where an a:outerShdw with the same\n' +
    'dist and dir would have put it. Rows that match are a plain offset shadow\n' +
    'under a different element name; rows that do not are something else.',
);

const presetShadows: Record<
  string,
  { box: Box | null; dx: number | null; dy: number | null; w: number | null; h: number | null }
> = {};
for (let n = 1; n <= 20; n++) {
  const m = byId.get(`prst-shdw${String(n)}`);
  const red = m?.boxes?.red ?? null;
  if (m === undefined) continue;
  if (red === null) {
    presetShadows[m.id] = { box: null, dx: null, dy: null, w: null, h: null };
    console.log(`  shdw${String(n).padStart(2)}  nothing painted`);
    continue;
  }
  // Only the part not hidden behind the opaque shape is visible, so the right
  // and bottom edges are the reliable ones for a shadow thrown right.
  presetShadows[m.id] = {
    box: red,
    dx: round(red.x1 - (m.rect.x + m.rect.cx), 2),
    dy: round(red.y0 - m.rect.y, 2),
    w: round(red.x1 - red.x0, 2),
    h: round(red.y1 - red.y0, 2),
  };
  console.log(
    `  shdw${String(n).padStart(2)}  visible box ${String(round(red.x1 - red.x0, 1)).padStart(6)} x ${String(round(red.y1 - red.y0, 1)).padStart(6)}pt` +
      `   right edge ${String(round(red.x1 - (m.rect.x + m.rect.cx), 1)).padStart(6)}pt past the shape,` +
      ` top ${String(round(red.y0 - m.rect.y, 1)).padStart(6)}pt from its top`,
  );
}
findings['presetShadows'] = presetShadows;

/* ---- L3. does the shadow turn with the shape? ---------------------------- */

section(
  'L3. rotWithShape on a rotated shape',
  'A 45-degree shape with a shadow thrown at dir=0. If rotWithShape is honoured,\n' +
    'one of the two lands along the rotated axis and the other stays horizontal.',
);

const rotWith: Record<string, { box: Box | null; centre: [number, number] | null }> = {};
for (const rot of [0, 1]) {
  const m = byId.get(`shdw-rotwith${String(rot)}`);
  const red = m?.boxes?.red ?? null;
  if (m === undefined) continue;
  rotWith[m.id] = {
    box: red,
    centre: red === null ? null : [round((red.x0 + red.x1) / 2, 2), round((red.y0 + red.y1) / 2, 2)],
  };
  const shapeCentre: [number, number] = [m.rect.x + m.rect.cx / 2, m.rect.y + m.rect.cy / 2];
  console.log(
    `  rotWithShape=${String(rot)}  shadow box ${red === null ? 'none' : `${String(red.x0)},${String(red.y0)}..${String(red.x1)},${String(red.y1)}`}` +
      (red === null
        ? ''
        : `   centre offset from the shape's (${String(round((red.x0 + red.x1) / 2 - shapeCentre[0], 1))}, ${String(round((red.y0 + red.y1) / 2 - shapeCentre[1], 1))})pt`),
  );
}
findings['rotWithShape'] = rotWith;

/* ---- H2. does an arrowhead scale with the stroke? ------------------------ */

section(
  'H2. is a head really a fixed number of stroke widths?',
  'Every length in section H is quoted in stroke widths, which is only the right\n' +
    'unit if the same head at a different @w is the same number of them. A\n' +
    'triangle at med/med measured 3 x 3 at 8pt; here it is measured at 4 and 16.',
);

const headScale: Record<string, { w: number; length: number; width: number }> = {};
for (const w of [4, 16]) {
  const m = byId.get(`ah-scale${String(w)}`);
  if (m?.columns == null) continue;
  const spanP = unpackHalf(m.columns.span);
  const inkP = unpackHalf(m.columns.ink);
  const scale = m.columns.scale;
  const shaftPx = w * scale;
  let tip = -1;
  let back = -1;
  let maxSpan = 0;
  for (let i = 0; i < spanP.length; i++) {
    if (inkP[i]! > 0.25) tip = i;
    if (spanP[i]! > shaftPx * 1.15 && back === -1) back = i;
    maxSpan = Math.max(maxSpan, spanP[i]!);
  }
  const length = tip === -1 || back === -1 ? 0 : (tip - back + 1) / scale / w;
  headScale[m.id] = { w, length: round(length, 3), width: round(maxSpan / scale / w, 3) };
  console.log(
    `  w=${String(w).padStart(2)}pt   thresholded length ${String(round(length, 2)).padStart(6)} x w   width ${String(round(maxSpan / scale / w, 2)).padStart(6)} x w`,
  );
}
findings['headScale'] = headScale;

/* ---- M. the hostile probes ----------------------------------------------- */

section(
  'M. what does PowerPoint refuse?',
  'Fifteen probes, each alone in its own package so a repair names its own\n' +
    'cause. "REFUSED" means the file would not open at all; "REPAIRED" means it\n' +
    'opened only when repair was allowed, which is the same verdict on the\n' +
    'markup and a different one on the user experience.',
);

const refusals: Record<string, { opened: boolean; repaired: boolean | null; error: string | null }> = {};
for (const m of measured.filter((p) => p.group === 'hostile')) {
  refusals[m.id] = { opened: m.opened, repaired: m.repaired, error: m.error };
  const verdict = !m.opened ? 'REFUSED' : m.repaired ? 'REPAIRED' : 'accepted';
  console.log(`  ${m.id.padEnd(14)} ${verdict.padEnd(9)} ${m.question}`);
}
const accepted = Object.entries(refusals).filter(([, v]) => v.opened && v.repaired === false);
console.log(
  `\n  ${String(Object.keys(refusals).length - accepted.length)} of ${String(Object.keys(refusals).length)} refused; accepted: ${accepted.map(([k]) => k).join(', ') || 'none'}`,
);
findings['refusals'] = refusals;

/* ---- N. the four blurs are one blur --------------------------------------- */

section(
  'N. every blurred edge in DrawingML, against one rule',
  'a:outerShdw/@blurRad, a:blur/@rad, a:glow/@rad and a:softEdge/@rad are four\n' +
    'attributes on four elements. If they are one operator with one constant, a\n' +
    'single pair of rules fits all of them: a Gaussian of sigma r/3, and an edge\n' +
    'moved 0.9r - outward for a glow, inward for a soft edge, and not at all for\n' +
    'a shadow or a blur. The r a glow uses is half the radius it declares.',
);

interface BlurRule {
  element: string;
  declared: number;
  r: number;
  sigma: number;
  sigmaOverR: number;
  offset: number;
  offsetOverR: number;
}

const blurRules: BlurRule[] = [];
const pushRule = (element: string, id: string, declared: number, r: number): void => {
  const fit =
    element === 'outerShdw'
      ? (blurFits[id] ?? null)
      : (findings['effectReach'] as Record<string, EffectReach> | undefined)?.[id]?.fit ?? null;
  const offset =
    element === 'outerShdw'
      ? 0
      : ((findings['effectReach'] as Record<string, EffectReach>)[id]?.fitOffset ?? 0);
  if (fit === null || r === 0) return;
  blurRules.push({
    element,
    declared,
    r,
    sigma: fit.sigma,
    sigmaOverR: round(fit.sigma / r, 4),
    offset: round(offset, 3),
    offsetOverR: round(offset / r, 4),
  });
};
for (const rad of [4, 8, 16, 24]) pushRule('outerShdw', `shdw-blur${String(rad)}`, rad, rad);
for (const rad of [4, 16]) pushRule('blur', `blur-${String(rad)}`, rad, rad);
for (const rad of [4, 8, 16]) pushRule('glow', `glow-${String(rad)}`, rad, rad / 2);
for (const rad of [2, 4, 8, 16, 32]) pushRule('softEdge', `soft-${String(rad)}`, rad, rad);

for (const rule of blurRules) {
  console.log(
    `  ${rule.element.padEnd(10)} declared ${String(rule.declared).padStart(3)}pt -> r=${String(rule.r).padStart(4)}pt` +
      `   sigma/r ${rule.sigmaOverR.toFixed(4)}   edge offset/r ${rule.offsetOverR.toFixed(4)}`,
  );
}
{
  const sigmas = blurRules.map((r) => r.sigmaOverR);
  const shifted = blurRules.filter((r) => Math.abs(r.offsetOverR) > 0.2).map((r) => Math.abs(r.offsetOverR));
  const spread = (v: number[]): string =>
    v.length === 0 ? '-' : `${Math.min(...v).toFixed(4)}..${Math.max(...v).toFixed(4)}`;
  console.log(
    `\n  sigma/r over all ${String(sigmas.length)} measurements: ${spread(sigmas)}   (one third is ${(1 / 3).toFixed(4)})`,
  );
  console.log(
    `  |edge offset|/r where there is one, ${String(shifted.length)} measurements: ${spread(shifted)}`,
  );
  findings['blurRule'] = {
    rules: blurRules,
    sigmaOverR: { min: Math.min(...sigmas), max: Math.max(...sigmas) },
    offsetOverR: shifted.length === 0 ? null : { min: Math.min(...shifted), max: Math.max(...shifted) },
  };
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

if (fixturePath != null) {
  const fixture = {
    generatedBy: 'tools/ground-truth/analyse-lines.ts',
    experiment: 'C4 - strokes and effects',
    sourceNote:
      'Measured against Microsoft PowerPoint 365 by exporting probe decks this repository ' +
      'generated and reading the bitmaps. Positions are sub-pixel crossings of half coverage, ' +
      'in points on a 960x540pt slide. Coverage is 1 - min(r,g,b)/255. Run ' +
      'tools/ground-truth/build-line-deck.ts, read-lines.ps1 and analyse-lines.ts to regenerate, ' +
      'then prettier --write this file before committing.',
    slide: inputs.slidePoints,
    strokeWidths: { dash: DASH_W, cap: 24, algn: 20, cmpd: 36, join: 20, arrowhead: AH_W },
    findings,
    probes: measured,
  };
  writeFileSync(fixturePath, JSON.stringify(fixture, null, 2));
  console.log(`\nwrote ${fixturePath}`);
}

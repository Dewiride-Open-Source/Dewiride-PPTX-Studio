/**
 * Experiment F3, step 3 - read the bitmaps and say where each edge landed.
 *
 * `node tools/ground-truth/render/snap/analyse.ts <work-dir> [--capture] [--fixture <path>]`. Every
 * probe's coverage profile is read on both exports and every candidate's profile is scored against
 * every case. The rule is read from `GRID_WIDTHS`, and the run throws unless exactly one reading
 * per family fits all of them; every other width is scored against that rule and characterised.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp, type Bitmap } from '../../lib/bmp.ts';
import { fixtureJson } from '../../lib/json.ts';
import { ink } from '../../lib/profile.ts';
import { repoPath } from '../../../repo/root.ts';

import {
  axisScale,
  EXPORT_WIDTHS,
  exportShape,
  GRID_WIDTHS,
  heightOf,
  MODELS_BY_FAMILY,
  predictedProfile,
  profileOf,
  SLIDE,
  snapProbes,
  TOLERANCE,
  windowOf,
  type Family,
  type Model,
  type Probe,
} from './probes.ts';

const dirArg = process.argv[2];
if (dirArg === undefined)
  throw new Error('usage: analyse.ts <work-dir> [--capture] [--fixture <path>]');
const dir = dirArg;
const capture = process.argv.includes('--capture');
const fixtureAt = process.argv.indexOf('--fixture');
const fixtureArg = fixtureAt === -1 ? undefined : process.argv[fixtureAt + 1];
if (fixtureAt !== -1 && fixtureArg === undefined) throw new Error('--fixture needs a path');
const fixturePath = fixtureArg ?? repoPath('corpus/ground-truth/snap.json');

/** How many lines along an edge are averaged. */
const SAMPLES_ALONG = 40;

/** How many of a losing model's misses the fixture names; `fits` and `of` carry the count. */
const MISSES_KEPT = 6;

/** How many of the rule's misses at one width the fixture names; past that the width is another regime. */
const WIDTH_MISSES_NAMED = 24;

if (capture) {
  const script = repoPath('tools/ground-truth/render/snap/read.ps1');
  const run = spawnSync(
    'powershell',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-Dir', dir],
    { stdio: 'inherit' },
  );
  if (run.status !== 0) throw new Error(`read.ps1 exited ${String(run.status)}`);
}

interface Export {
  readonly slide: number;
  readonly width: number;
  readonly height: number;
  readonly pass: 'a' | 'b';
  readonly file: string;
}

const exported = JSON.parse(readFileSync(join(dir, 'snap-export.json'), 'utf8')) as {
  readonly powerpoint: string;
  readonly opened: boolean;
  readonly repaired: boolean;
  readonly exports: readonly Export[];
  readonly refused: readonly { slide: number; width: number; pass: string; error: string }[];
};
if (!exported.opened || exported.repaired) {
  throw new Error(`PowerPoint did not open the deck clean: ${JSON.stringify(exported.refused)}`);
}

const probes = snapProbes();

const bitmaps = new Map<string, Bitmap>();
function bitmap(file: string, width: number): Bitmap {
  let bmp = bitmaps.get(file);
  if (bmp === undefined) {
    bmp = readBmp(new Uint8Array(readFileSync(join(dir, file))));
    if (bmp.width !== width || bmp.height !== heightOf(width)) {
      throw new Error(
        `${file} is ${String(bmp.width)}x${String(bmp.height)}, not ${String(width)}x${String(heightOf(width))}`,
      );
    }
    bitmaps.set(file, bmp);
  }
  return bmp;
}

function exportOf(slide: number, width: number, pass: 'a' | 'b'): Export | null {
  return (
    exported.exports.find((e) => e.slide === slide && e.width === width && e.pass === pass) ?? null
  );
}

/* -------------------------------------------------------------------------- */
/* measurement                                                                */
/* -------------------------------------------------------------------------- */

/** The coverage profile across one edge: one number per device row (or column) of the window. */
interface Measure {
  /** The first row `cover` describes, in device pixels; every row of the window outside it is clear. */
  readonly from: number;
  readonly cover: readonly number[];
}

/** Rows neither clear nor full: an antialiased edge has one, a snapped one none. */
const partialOf = (cover: readonly number[]): number =>
  cover.filter((v) => v > 0.1 && v < 0.9).length;

function measure(probe: Probe, bmp: Bitmap, scale: number): Measure {
  const { from, to } = windowOf(probe, scale);
  const limit = probe.read.axis === 'y' ? bmp.height : bmp.width;
  if (from < 0 || to >= limit) {
    throw new Error(
      `${probe.id}@${String(bmp.width)}: window ${String(from)}..${String(to)} is off the slide`,
    );
  }
  const alongScale = probe.read.axis === 'y' ? bmp.width / SLIDE.w : bmp.height / SLIDE.h;
  const along0 = Math.round(probe.read.along0 * alongScale);
  const along1 = Math.round(probe.read.along1 * alongScale);
  const step = Math.max(1, Math.floor((along1 - along0) / SAMPLES_ALONG));
  const sums = new Array<number>(to - from + 1).fill(0);
  let lines = 0;
  for (let at = along0; at <= along1; at += step) {
    for (let r = from; r <= to; r++) {
      sums[r - from]! += probe.read.axis === 'y' ? ink(bmp, at, r) : ink(bmp, r, at);
    }
    lines += 1;
  }
  let cover = sums.map((v) => v / lines);
  // A line end is read along one row of the stroke, whose own coverage is the stroke family's question.
  if (probe.family === 'end') {
    const plateau = Math.max(...cover);
    if (plateau > 0) cover = cover.map((v) => v / plateau);
  }
  return { from, cover };
}

/* -------------------------------------------------------------------------- */
/* the two exports                                                            */
/* -------------------------------------------------------------------------- */

interface Case {
  readonly probe: Probe;
  readonly width: number;
  /** Device pixels per point along the read axis: the rounded height's for rows. */
  readonly scale: number;
  readonly a: Measure;
  readonly b: Measure;
}

const cases: Case[] = [];
const missing: string[] = [];
for (const probe of probes) {
  for (const width of EXPORT_WIDTHS) {
    const a = exportOf(probe.slide, width, 'a');
    const b = exportOf(probe.slide, width, 'b');
    if (a === null || b === null) {
      missing.push(`${probe.id}@${String(width)}`);
      continue;
    }
    const bmpA = bitmap(a.file, width);
    const scale = probe.read.axis === 'y' ? bmpA.height / SLIDE.h : bmpA.width / SLIDE.w;
    cases.push({
      probe,
      width,
      scale,
      a: measure(probe, bmpA, scale),
      b: measure(probe, bitmap(b.file, width), scale),
    });
  }
}

const worstOf = (x: readonly number[], y: readonly number[]): number =>
  Math.max(...x.map((v, i) => Math.abs(v - (y[i] ?? 0))));

const disagreeing = cases.filter((c) => worstOf(c.a.cover, c.b.cover) > TOLERANCE);
const excluded = new Set(disagreeing.map((c) => `${c.probe.id}@${String(c.width)}`));
const usable = (c: Case): boolean => !excluded.has(`${c.probe.id}@${String(c.width)}`);

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Miss {
  readonly probe: string;
  readonly width: number;
  readonly worst: number;
}

interface Score {
  readonly model: string;
  readonly description: string;
  readonly fits: number;
  readonly of: number;
  readonly worstError: number;
  /** The first `MISSES_KEPT` cases the model missed. */
  readonly misses: readonly Miss[];
}

const errorOf = (model: Model, c: Case): number =>
  worstOf(c.a.cover, predictedProfile(model, c.probe, c.scale));

function scoreModels(
  family: Family,
  models: readonly Model[],
  applies: (c: Case) => boolean,
): Score[] {
  const pool = cases.filter((c) => c.probe.family === family && usable(c) && applies(c));
  return models.map((model) => {
    const misses: Miss[] = [];
    let worst = 0;
    for (const c of pool) {
      const error = errorOf(model, c);
      worst = Math.max(worst, error);
      if (error > TOLERANCE) {
        misses.push({ probe: c.probe.id, width: c.width, worst: Math.round(error * 1000) / 1000 });
      }
    }
    return {
      model: model.name,
      description: model.description,
      fits: pool.length - misses.length,
      of: pool.length,
      worstError: Math.round(worst * 1000) / 1000,
      misses: misses.slice(0, MISSES_KEPT),
    };
  });
}

function winner(family: string, scores: readonly Score[]): Score {
  const perfect = scores.filter((s) => s.fits === s.of && s.of > 0);
  if (perfect.length !== 1) {
    const table = scores.map((s) => `${s.model} ${String(s.fits)}/${String(s.of)}`).join(', ');
    const most = Math.max(...scores.map((s) => s.fits));
    const misses = scores
      .filter((s) => s.fits === most)
      .flatMap((s) =>
        s.misses
          .slice(0, 6)
          .map((m) => `${s.model} off by ${String(m.worst)} at ${m.probe}@${String(m.width)}`),
      )
      .join('; ');
    throw new Error(
      `no single ${family} model fits every case (${table}) - the experiment is inconclusive; the closest misses: ${misses}`,
    );
  }
  return perfect[0]!;
}

const families = Object.keys(MODELS_BY_FAMILY) as Family[];
const onGrid = (c: Case): boolean => GRID_WIDTHS.includes(c.width);
const candidates: Record<string, Score[]> = {};
for (const family of families) {
  candidates[family] = scoreModels(family, MODELS_BY_FAMILY[family], onGrid);
  console.log(
    `${family.padEnd(10)} ${candidates[family].map((s) => `${s.model} ${String(s.fits)}/${String(s.of)}`).join('  ')}`,
  );
}

const inconclusive: string[] = [];
const findings: Record<string, string> = {};
for (const family of families) {
  try {
    const won = winner(family, candidates[family]!);
    findings[family] = won.model;
    findings[`${family}Description`] = won.description;
  } catch (error) {
    inconclusive.push(error instanceof Error ? error.message : String(error));
  }
}
if (inconclusive.length > 0) throw new Error(inconclusive.join('\n'));

/** Rows the export left partial on the grid widths, per family: a snapped edge leaves none. */
const partialRows: Record<string, number> = {};
for (const family of families) {
  partialRows[family] = cases
    .filter((c) => c.probe.family === family && onGrid(c))
    .reduce((sum, c) => sum + partialOf(c.a.cover), 0);
}

/* -------------------------------------------------------------------------- */
/* every width on its own                                                     */
/* -------------------------------------------------------------------------- */

/** How close a coverage is to a multiple of a quarter, as a 2x2 supersampler would leave it. */
const offQuarter = (v: number): number => Math.abs(v * 4 - Math.round(v * 4)) / 4;

/** The dots per inch a width is, on a 13⅓-inch slide. */
const dpiOf = (width: number): number => (width * 72) / SLIDE.w;

/**
 * How a width off the grid could come from a render on it: the rule at some other scale, and
 * that render's band carried to this width's pixels.
 */
interface Mapping {
  readonly name: string;
  readonly description: string;
  /** The scale the rule is applied at, and the map from that scale's pixels to this width's. */
  readonly through: (
    width: number,
    axis: 'x' | 'y',
  ) => { readonly scale: number; readonly to: (v: number) => number } | null;
}

/** The stretch of a rounded height over the true one, as a fraction: zero across columns. */
const stretchOf = (width: number, axis: 'x' | 'y'): number =>
  axis === 'x' ? 0 : heightOf(width) / ((width * SLIDE.h) / SLIDE.w) - 1;

const uniform = (width: number): number => width / SLIDE.w;

const resampled = (round: (dpi: number) => number): Mapping['through'] => {
  return (width, axis) => {
    const dpi = round(dpiOf(width));
    if (dpi === dpiOf(width)) return null;
    const scale = dpi / 72;
    const own = axisScale(width, axis);
    return { scale, to: (v) => (v * own) / scale };
  };
};

const MAPPINGS: readonly Mapping[] = [
  {
    name: 'uniform',
    description:
      'the rule at the width over the slide width, rows stretched to the rounded height exactly',
    through: (width, axis) => {
      const stretch = stretchOf(width, axis);
      return stretch === 0 ? null : { scale: uniform(width), to: (v) => v * (1 + stretch) };
    },
  },
  {
    name: 'quarterStep',
    description:
      'the rule at the width over the slide width, rows moved by the stretch rounded up to a quarter pixel',
    through: (width, axis) => {
      const stretch = stretchOf(width, axis);
      return stretch === 0
        ? null
        : { scale: uniform(width), to: (v) => v + Math.ceil(v * stretch * 4) / 4 };
    },
  },
  {
    name: 'dpiDown',
    description: 'the rule at the whole dots per inch below, resampled to the width',
    through: resampled(Math.floor),
  },
  {
    name: 'dpiUp',
    description: 'the rule at the whole dots per inch above, resampled to the width',
    through: resampled(Math.ceil),
  },
];

/** The family's winning model under a mapping, or `null` where the mapping has nothing to say. */
function errorUnder(mapping: Mapping, c: Case): number | null {
  const route = mapping.through(c.width, c.probe.read.axis);
  if (route === null) return null;
  const model = MODELS_BY_FAMILY[c.probe.family].find((m) => m.name === findings[c.probe.family])!;
  const { c: centre, w } = windowOf(c.probe, route.scale);
  const band = model.predict(centre, w, c.probe, route.scale);
  const carry = (v: number): number => (Number.isFinite(v) ? route.to(v) : v);
  const { from, to } = windowOf(c.probe, c.scale);
  return worstOf(
    c.a.cover,
    profileOf({ top: carry(band.top), bottom: carry(band.bottom) }, from, to),
  );
}

/** One width: what it is, how the rule did on it, and what its partial rows look like. */
interface AtWidth {
  readonly height: number;
  readonly dpi: number;
  readonly eighth: boolean;
  readonly wholeHeight: boolean;
  readonly wholeDpi: boolean;
  readonly cases: number;
  /** The rule's fits and cases per family: `[fits, of]`. */
  readonly rule: Readonly<Record<string, readonly [number, number]>>;
  /** The cases the rule missed, by name, where there are few enough to name. */
  readonly ruleMisses: readonly string[];
  /** The rule's fits per family under each mapping that applies here, and the cases it was tried on. */
  readonly mappings: Readonly<Record<string, Readonly<Record<string, readonly [number, number]>>>>;
  /** Rows neither clear nor full, over every case: a snapped width has none but the curves' and borders'. */
  readonly partialRows: number;
  /** Of the partial rows, how many are within 0.02 of a quarter: all of them under a 2x2 supersampler. */
  readonly quarterRows: number;
}

const atWidth: Record<string, AtWidth> = {};
for (const width of EXPORT_WIDTHS) {
  const here = cases.filter((c) => c.width === width && usable(c));
  const rule: Record<string, [number, number]> = {};
  const ruleMisses: string[] = [];
  const mappings: Record<string, Record<string, [number, number]>> = {};
  for (const family of families) {
    const pool = here.filter((c) => c.probe.family === family);
    const won = MODELS_BY_FAMILY[family].find((m) => m.name === findings[family])!;
    const missed = pool.filter((c) => errorOf(won, c) > TOLERANCE);
    rule[family] = [pool.length - missed.length, pool.length];
    ruleMisses.push(...missed.map((c) => c.probe.id));
    for (const mapping of MAPPINGS) {
      const errors = pool.map((c) => errorUnder(mapping, c)).filter((e) => e !== null);
      if (errors.length === 0) continue;
      (mappings[mapping.name] ??= {})[family] = [
        errors.filter((e) => e <= TOLERANCE).length,
        errors.length,
      ];
    }
  }
  const partials = here.flatMap((c) => c.a.cover.filter((v) => v > 0.1 && v < 0.9));
  const shape = exportShape(width);
  atWidth[String(width)] = {
    height: shape.height,
    dpi: dpiOf(width),
    eighth: shape.eighth,
    wholeHeight: shape.wholeHeight,
    wholeDpi: shape.wholeDpi,
    cases: here.length,
    rule,
    ruleMisses: ruleMisses.length > WIDTH_MISSES_NAMED ? [] : ruleMisses,
    mappings,
    partialRows: partials.length,
    quarterRows: partials.filter((v) => offQuarter(v) <= 0.02).length,
  };
  for (const [name, byFamily] of Object.entries(mappings)) {
    const fits = Object.values(byFamily).reduce((sum, [f]) => sum + f, 0);
    const of = Object.values(byFamily).reduce((sum, [, n]) => sum + n, 0);
    console.log(
      `${String(width).padStart(5)} px under ${name.padEnd(12)} ${String(fits)}/${String(of)}`,
    );
  }
  const held = Object.values(rule).every(([fits, of]) => fits === of);
  console.log(
    `${String(width).padStart(5)} px ${held ? 'holds' : 'FAILS'}  ` +
      Object.entries(rule)
        .filter(([, [fits, of]]) => fits !== of)
        .map(([family, [fits, of]]) => `${family} ${String(fits)}/${String(of)}`)
        .join('  ') +
      `  partial ${String(partials.length)}, quarters ${String(atWidth[String(width)]!.quarterRows)}`,
  );
}

const holdsAt = EXPORT_WIDTHS.filter((width) =>
  Object.values(atWidth[String(width)]!.rule).every(([fits, of]) => fits === of),
);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

/** The inked rows and one clear row each side, so the file carries the edge and not the window. */
function trimmed(measure: Measure): Measure {
  const cover = measure.cover.map((v) => Math.round(v * 1000) / 1000);
  let first = cover.findIndex((v) => v > 0);
  if (first === -1) return { from: measure.from, cover: [] };
  let last = cover.length - 1;
  while (last > first && cover[last] === 0) last -= 1;
  first = Math.max(0, first - 1);
  last = Math.min(cover.length - 1, last + 1);
  return { from: measure.from + first, cover: cover.slice(first, last + 1) };
}

const byProbe: Record<string, Record<string, Measure>> = {};
for (const c of cases) {
  (byProbe[c.probe.id] ??= {})[String(c.width)] = trimmed(c.a);
}

const fixture = {
  $comment:
    'GENERATED by tools/ground-truth/render/snap/analyse.ts from PowerPoint exports of the F3 probe deck. Do not edit by hand.',
  experiment: 'F3',
  powerpoint: exported.powerpoint,
  slide: SLIDE,
  widths: EXPORT_WIDTHS,
  gridWidths: GRID_WIDTHS,
  tolerance: TOLERANCE,
  cases: cases.length,
  missing,
  excluded: [...excluded],
  findings: {
    ...findings,
    partialRows,
    holdsAt,
    failsAt: EXPORT_WIDTHS.filter((width) => !holdsAt.includes(width)),
    exportCeiling: exported.refused.length === 0 ? null : exported.refused,
  },
  candidates,
  atWidth,
  probes: probes.map(({ markup: _markup, ...rest }) => rest),
  measured: byProbe,
};

writeFileSync(fixturePath, fixtureJson(fixture));

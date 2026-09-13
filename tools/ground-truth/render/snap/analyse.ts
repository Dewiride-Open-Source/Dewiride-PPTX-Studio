/**
 * Experiment F3, step 3 - read the bitmaps and say where each edge landed.
 *
 * `node tools/ground-truth/render/snap/analyse.ts <work-dir> [--capture] [--fixture <path>]`. Every
 * probe's coverage profile is read on both exports, every candidate's profile is scored against
 * every case from `SNAP_FROM` up, and the run throws unless exactly one reading per family fits
 * all of them: an imperfect fit records a guess.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp, type Bitmap } from '../../lib/bmp.ts';
import { ink } from '../../lib/profile.ts';
import { repoPath } from '../../../repo/root.ts';

import {
  EXPORT_WIDTHS,
  MODELS_BY_FAMILY,
  predictedProfile,
  SLIDE,
  SNAP_FROM,
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
const MISSES_KEPT = 12;

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
function bitmap(file: string): Bitmap {
  let bmp = bitmaps.get(file);
  if (bmp === undefined) {
    bmp = readBmp(new Uint8Array(readFileSync(join(dir, file))));
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
  readonly peak: number;
  /** Rows neither clear nor full: an antialiased edge has one, a snapped one none. */
  readonly partial: number;
}

function measure(probe: Probe, bmp: Bitmap, width: number): Measure {
  const s = width / SLIDE.w;
  const { from, to } = windowOf(probe, s);
  const limit = probe.read.axis === 'y' ? bmp.height : bmp.width;
  if (from < 0 || to >= limit) {
    throw new Error(
      `${probe.id}@${String(width)}: window ${String(from)}..${String(to)} is off the slide`,
    );
  }
  const along0 = Math.round(probe.read.along0 * s);
  const along1 = Math.round(probe.read.along1 * s);
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
  return {
    from,
    cover,
    peak: Math.max(...cover),
    partial: cover.filter((v) => v > 0.1 && v < 0.9).length,
  };
}

/* -------------------------------------------------------------------------- */
/* the two exports                                                            */
/* -------------------------------------------------------------------------- */

interface Case {
  readonly probe: Probe;
  readonly width: number;
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
    cases.push({
      probe,
      width,
      scale: width / SLIDE.w,
      a: measure(probe, bitmap(a.file), width),
      b: measure(probe, bitmap(b.file), width),
    });
  }
}

const worstOf = (x: readonly number[], y: readonly number[]): number =>
  Math.max(...x.map((v, i) => Math.abs(v - (y[i] ?? 0))));

const disagreeing = cases.filter((c) => worstOf(c.a.cover, c.b.cover) > TOLERANCE);
const excluded = new Set(disagreeing.map((c) => `${c.probe.id}@${String(c.width)}`));

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

function scoreModels(
  family: Family,
  models: readonly Model[],
  applies: (c: Case) => boolean,
): Score[] {
  const pool = cases.filter(
    (c) =>
      c.probe.family === family && !excluded.has(`${c.probe.id}@${String(c.width)}`) && applies(c),
  );
  return models.map((model) => {
    const misses: Miss[] = [];
    let worst = 0;
    for (const c of pool) {
      const error = worstOf(c.a.cover, predictedProfile(model, c.probe, c.scale));
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
const candidates: Record<string, Score[]> = {};
const below: Record<string, Score[]> = {};
for (const family of families) {
  candidates[family] = scoreModels(family, MODELS_BY_FAMILY[family], (c) => c.width >= SNAP_FROM);
  below[family] = scoreModels(family, MODELS_BY_FAMILY[family], (c) => c.width < SNAP_FROM);
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

/** Rows the export left partial from `SNAP_FROM` up, per family: a snapped edge leaves none. */
const partialRows: Record<string, number> = {};
for (const family of families) {
  partialRows[family] = cases
    .filter((c) => c.probe.family === family && c.width >= SNAP_FROM)
    .reduce((sum, c) => sum + c.a.partial, 0);
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

/** The inked rows and one clear row each side, so the file carries the edge and not the window. */
function trimmed(measure: Measure): Measure {
  const cover = measure.cover.map((v) => Math.round(v * 1000) / 1000);
  let first = cover.findIndex((v) => v > 0);
  if (first === -1) return { ...measure, cover: [] };
  let last = cover.length - 1;
  while (last > first && cover[last] === 0) last -= 1;
  first = Math.max(0, first - 1);
  last = Math.min(cover.length - 1, last + 1);
  return {
    from: measure.from + first,
    cover: cover.slice(first, last + 1),
    peak: Math.round(measure.peak * 1000) / 1000,
    partial: measure.partial,
  };
}

const byProbe: Record<string, Record<string, Measure>> = {};
for (const c of cases) {
  (byProbe[c.probe.id] ??= {})[String(c.width)] = trimmed(c.a);
}

/** Under the scored width no model is expected to fit; the file keeps the counts and not the cases. */
const belowSummary = Object.fromEntries(
  Object.entries(below).map(([family, scores]) => [
    family,
    scores.map(({ misses: _misses, ...score }) => score),
  ]),
);

const fixture = {
  $comment:
    'GENERATED by tools/ground-truth/render/snap/analyse.ts from PowerPoint exports of the F3 probe deck. Do not edit by hand.',
  experiment: 'F3',
  powerpoint: exported.powerpoint,
  slide: SLIDE,
  widths: EXPORT_WIDTHS,
  snapFrom: SNAP_FROM,
  tolerance: TOLERANCE,
  cases: cases.length,
  missing,
  excluded: [...excluded],
  findings: {
    ...findings,
    partialRows,
    exportCeiling: exported.refused.length === 0 ? null : exported.refused,
  },
  candidates,
  below: belowSummary,
  probes: probes.map(({ markup: _markup, ...rest }) => rest),
  measured: byProbe,
};

writeFileSync(fixturePath, `${JSON.stringify(fixture, null, 2)}\n`);

console.log(
  `F3: ${String(cases.length)} cases, ${String(excluded.size)} excluded, PowerPoint ${exported.powerpoint}`,
);
for (const family of families) {
  const won = candidates[family]!.find((s) => s.model === findings[family])!;
  console.log(
    `  ${family.padEnd(10)} ${won.model}  ${won.description}  (${String(won.fits)}/${String(won.of)}, partial rows ${String(partialRows[family])})`,
  );
}
console.log(`wrote ${fixturePath}`);

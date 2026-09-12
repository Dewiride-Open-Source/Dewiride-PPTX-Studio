/**
 * Experiment F2, step 3 - read the bitmaps and say what zoom did.
 *
 * `node tools/ground-truth/render/zoom/analyse.ts <work-dir> [--capture] [--fixture <path>]`. Every
 * probe is measured on both exports, every candidate is scored against every case, and the run
 * throws unless exactly one reading per family fits all of them: an imperfect fit records a guess.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp, type Bitmap } from '../../lib/bmp.ts';
import { crossings, ink, runsOf, sampleCol, sampleRow, type Sampled } from '../../lib/profile.ts';
import { repoPath } from '../../../repo/root.ts';

import {
  BORDER_EDGE_MODELS,
  EXPORT_WIDTHS,
  MODELS_BY_FAMILY,
  SLIDE,
  zoomProbes,
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
const fixturePath =
  fixtureAt === -1 ? repoPath('corpus/ground-truth/zoom.json') : process.argv[fixtureAt + 1]!;

/** Ink agreement between the two exports, and between a model and a case, in pixels. */
const TOLERANCE_PX = 0.15;

/** A silhouette is a whole-pixel box, so a head agrees to a pixel. */
const BOX_TOLERANCE_PX = 1;

if (capture) {
  const script = repoPath('tools/ground-truth/render/zoom/read.ps1');
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

const exported = JSON.parse(readFileSync(join(dir, 'zoom-export.json'), 'utf8')) as {
  readonly powerpoint: string;
  readonly opened: boolean;
  readonly repaired: boolean;
  readonly exports: readonly Export[];
  readonly refused: readonly { slide: number; width: number; pass: string; error: string }[];
};
if (!exported.opened || exported.repaired) {
  throw new Error(`PowerPoint did not open the deck clean: ${JSON.stringify(exported.refused)}`);
}

const probes = zoomProbes();

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

/** What one read of one probe at one width came to, all in pixels of that export. */
interface Measure {
  /** Ink across the stroke, averaged over the columns or rows read. */
  readonly inkPx?: number;
  /** The highest coverage seen in one sample. */
  readonly peak?: number;
  /** Width between the half-coverage crossings, where the peak reaches one half. */
  readonly thicknessPx?: number | null;
  /** The falling crossing's distance inside the window's centre line, in pixels. */
  readonly innerEdgePx?: number | null;
  /** A dash period, in pixels; null where the row has no gaps. */
  readonly periodPx?: number | null;
  /** Ink along the row as a share of its length. */
  readonly coverage?: number;
  /** The ink box inside the region. */
  readonly boxW?: number;
  readonly boxH?: number;
  /** Total ink in the region. */
  readonly inkSum?: number;
  /** Luma 0..255 at ten normalised positions along a ramp. */
  readonly ramp?: readonly number[];
  /** Distinct luma levels along the ramp row. */
  readonly levels?: number;
  /** Mean coverage down a pattern column. */
  readonly mean?: number;
  /** Rows at the top and bottom of the bitmap that are not the background. */
  readonly bandRows?: number;
  readonly bitmapHeight?: number;
}

function acrossStroke(
  bmp: Bitmap,
  s: number,
  probe: Probe,
  read: { kind: 'across-h' | 'across-v'; x0: number; x1: number; y0: number; y1: number },
): Measure {
  const along = read.kind === 'across-h' ? [read.x0, read.x1] : [read.y0, read.y1];
  const from = Math.round(along[0]! * s);
  const to = Math.round(along[1]! * s);
  const step = Math.max(1, Math.floor((to - from) / 40));
  const centre = read.kind === 'across-h' ? (read.y0 + read.y1) / 2 : (read.x0 + read.x1) / 2;

  let inkSum = 0;
  let peak = 0;
  let count = 0;
  const thickness: number[] = [];
  const inner: number[] = [];
  for (let at = from; at <= to; at += step) {
    const pos = (at + 0.5) / s;
    const sampled: Sampled =
      read.kind === 'across-h'
        ? sampleCol(bmp, s, pos, read.y0, read.y1, probe.id)
        : sampleRow(bmp, s, pos, read.x0, read.x1, probe.id);
    const total = sampled.cover.reduce((a, b) => a + b, 0);
    inkSum += total;
    peak = Math.max(peak, ...sampled.cover);
    count += 1;
    const cross = crossings(sampled);
    const rise = cross.find((c) => c.rising);
    const fall = cross.find((c) => !c.rising);
    if (rise !== undefined && fall !== undefined && fall.at > rise.at) {
      thickness.push((fall.at - rise.at) * s);
      inner.push((fall.at - centre) * s);
    }
  }
  const mean = (values: readonly number[]): number | null =>
    values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
  const across = { inkPx: inkSum / count, peak, thicknessPx: mean(thickness) };
  return probe.family === 'border' ? { ...across, innerEdgePx: mean(inner) } : across;
}

/** The row nearest `y` with the most ink, so a hairline on a pixel boundary is still read whole. */
function inkedRow(bmp: Bitmap, s: number, y: number, x0: number, x1: number, what: string): number {
  let best = y;
  let most = -1;
  const half = 8;
  const rows = Math.max(1, Math.round(2 * half * s));
  for (let i = 0; i <= rows; i++) {
    const at = y - half + (2 * half * i) / rows;
    const sampled = sampleRow(bmp, s, at, x0, x1, what);
    const total = sampled.cover.reduce((a, b) => a + b, 0);
    if (total > most) {
      most = total;
      best = at;
    }
  }
  return best;
}

function alongStroke(
  bmp: Bitmap,
  s: number,
  probe: Probe,
  read: { kind: 'along-h'; y: number; x0: number; x1: number },
): Measure {
  const y = inkedRow(bmp, s, read.y, read.x0, read.x1, probe.id);
  const sampled = sampleRow(bmp, s, y, read.x0, read.x1, probe.id);
  const runs = runsOf(crossings(sampled));
  const coverage = sampled.cover.reduce((a, b) => a + b, 0) / Math.max(1, sampled.cover.length);
  const mean = (values: readonly number[]): number =>
    values.reduce((a, b) => a + b, 0) / values.length;
  const periodPx =
    runs.on.length >= 3 && runs.off.length >= 3 ? (mean(runs.on) + mean(runs.off)) * s : null;
  return { periodPx, coverage, peak: Math.max(...sampled.cover) };
}

function box(
  bmp: Bitmap,
  s: number,
  read: { kind: 'box'; x0: number; y0: number; x1: number; y1: number },
  threshold: number,
): Measure {
  const x0 = Math.max(0, Math.round(read.x0 * s));
  const y0 = Math.max(0, Math.round(read.y0 * s));
  const x1 = Math.min(bmp.width - 1, Math.round(read.x1 * s));
  const y1 = Math.min(bmp.height - 1, Math.round(read.y1 * s));
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  let inkSum = 0;
  for (let y = y0; y <= y1; y++) {
    for (let x = x0; x <= x1; x++) {
      const value = ink(bmp, x, y);
      inkSum += value;
      if (value > threshold) {
        minX = Math.min(minX, x);
        maxX = Math.max(maxX, x);
        minY = Math.min(minY, y);
        maxY = Math.max(maxY, y);
      }
    }
  }
  const empty = minX === Infinity;
  return { boxW: empty ? 0 : maxX - minX + 1, boxH: empty ? 0 : maxY - minY + 1, inkSum };
}

/** The spacing of the ink runs' centroids down one column, in pixels; null under three runs. */
function columnPeriod(sampled: Sampled, s: number): number | null {
  const cover = sampled.cover;
  const centroids: number[] = [];
  let runStart = -1;
  for (let i = 0; i <= cover.length; i++) {
    const inked = i < cover.length && cover[i]! > 0.3;
    if (inked && runStart === -1) runStart = i;
    if (!inked && runStart !== -1) {
      let weight = 0;
      let moment = 0;
      for (let k = runStart; k < i; k++) {
        weight += cover[k]!;
        moment += cover[k]! * k;
      }
      centroids.push(sampled.start + (moment / weight) * sampled.step);
      runStart = -1;
    }
  }
  const first = centroids[0];
  const last = centroids[centroids.length - 1];
  return centroids.length >= 3 && first !== undefined && last !== undefined
    ? ((last - first) / (centroids.length - 1)) * s
    : null;
}

function period(
  bmp: Bitmap,
  s: number,
  probe: Probe,
  read: { kind: 'period-v'; x: number; y0: number; y1: number },
): Measure {
  const sampled = sampleCol(bmp, s, read.x, read.y0, read.y1, probe.id);
  const mean = sampled.cover.reduce((a, b) => a + b, 0) / sampled.cover.length;
  // Run centroids see a line snapped to one half-intensity row; the median of five columns
  // steps past a checkerboard's blended column.
  const periods: number[] = [];
  for (let dx = -2; dx <= 2; dx++) {
    const column = sampleCol(bmp, s, read.x + dx / s, read.y0, read.y1, probe.id);
    const found = columnPeriod(column, s);
    if (found !== null) periods.push(found);
  }
  periods.sort((a, b) => a - b);
  const periodPx = periods.length === 0 ? null : periods[Math.floor(periods.length / 2)]!;
  return { periodPx, mean };
}

function ramp(
  bmp: Bitmap,
  s: number,
  probe: Probe,
  read: { kind: 'ramp-h'; y: number; x0: number; x1: number },
): Measure {
  const sampled = sampleRow(bmp, s, read.y, read.x0, read.x1, probe.id);
  const n = sampled.cover.length;
  const at = (fraction: number): number =>
    Math.round((1 - sampled.cover[Math.min(n - 1, Math.floor(fraction * n))]!) * 255);
  const positions = [0.05, 0.15, 0.25, 0.35, 0.45, 0.55, 0.65, 0.75, 0.85, 0.95];
  const levels = new Set(sampled.cover.map((c) => Math.round((1 - c) * 255))).size;
  return { ramp: positions.map(at), levels };
}

function frame(bmp: Bitmap): Measure {
  const background = ink(bmp, Math.floor(bmp.width / 2), Math.floor(bmp.height / 2));
  const isBand = (y: number): boolean => {
    let differing = 0;
    for (let x = 0; x < bmp.width; x++)
      if (Math.abs(ink(bmp, x, y) - background) > 0.1) differing++;
    return differing > bmp.width / 2;
  };
  let bandRows = 0;
  for (let y = 0; y < bmp.height && isBand(y); y++) bandRows++;
  for (let y = bmp.height - 1; y >= 0 && isBand(y); y--) bandRows++;
  return { bandRows, bitmapHeight: bmp.height };
}

function measure(probe: Probe, bmp: Bitmap, width: number): Measure {
  const s = width / SLIDE.w;
  const read = probe.read;
  switch (read.kind) {
    case 'across-h':
    case 'across-v':
      return acrossStroke(bmp, s, probe, read);
    case 'along-h':
      return alongStroke(bmp, s, probe, read);
    case 'box':
      // A head is a solid silhouette, so its antialiased fringe is not the head; faint glyphs are text.
      return box(bmp, s, read, probe.family === 'marker' ? 0.5 : 0.25);
    case 'period-v':
      return period(bmp, s, probe, read);
    case 'ramp-h':
      return ramp(bmp, s, probe, read);
    case 'frame':
      return frame(bmp);
  }
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

/** The one number a family's models predict, from a measure. */
function observed(family: Family, m: Measure): number | null | undefined {
  switch (family) {
    case 'hairline':
    case 'thin':
    case 'border':
      return m.inkPx;
    case 'dash':
      return m.periodPx;
    case 'marker':
      return m.boxH;
    case 'pattern':
      return m.periodPx;
    case 'text':
    case 'gradient':
    case 'frame':
      return undefined;
  }
}

const near = (x: number | null | undefined, y: number | null | undefined, tol: number): boolean =>
  x === null || y === null || x === undefined || y === undefined ? x === y : Math.abs(x - y) <= tol;

const disagreeing = cases.filter((c) => {
  const x = observed(c.probe.family, c.a);
  const y = observed(c.probe.family, c.b);
  if (x === undefined) return false;
  return !near(x, y, c.probe.family === 'marker' ? BOX_TOLERANCE_PX : TOLERANCE_PX);
});
const excluded = new Set(disagreeing.map((c) => `${c.probe.id}@${String(c.width)}`));

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Miss {
  readonly probe: string;
  readonly width: number;
  readonly observed: number | null;
  readonly predicted: number | null;
}

interface Score {
  readonly model: string;
  readonly description: string;
  readonly fits: number;
  readonly of: number;
  readonly worstErrorPx: number;
  readonly misses: readonly Miss[];
}

function scoreModels(
  family: Family,
  models: readonly Model[],
  value: (c: Case) => number | null | undefined,
  applies: (c: Case) => boolean = () => true,
  tolerance = TOLERANCE_PX,
): Score[] {
  const pool = cases.filter(
    (c) =>
      c.probe.family === family && !excluded.has(`${c.probe.id}@${String(c.width)}`) && applies(c),
  );
  return models.map((model) => {
    const misses: Miss[] = [];
    let worst = 0;
    let of = 0;
    for (const c of pool) {
      const seen = value(c);
      if (seen === undefined) continue;
      const predicted = model.predict(c.scale, c.probe);
      of += 1;
      if (seen !== null && predicted !== null) {
        worst = Math.max(worst, Math.abs(seen - predicted));
      }
      if (!near(seen, predicted, tolerance)) {
        misses.push({ probe: c.probe.id, width: c.width, observed: seen, predicted });
      }
    }
    return {
      model: model.name,
      description: model.description,
      fits: of - misses.length,
      of,
      worstErrorPx: Math.round(worst * 1000) / 1000,
      misses,
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
        s.misses.map(
          (m) =>
            `${m.probe}@${String(m.width)} saw ${m.observed === null ? 'nothing' : m.observed.toFixed(3)}, ${s.model} said ${m.predicted === null ? 'nothing' : m.predicted.toFixed(3)}`,
        ),
      )
      .join('; ');
    throw new Error(
      `no single ${family} model fits every case (${table}) - the experiment is inconclusive; the closest misses: ${misses}`,
    );
  }
  return perfect[0]!;
}

const hairlineScores = scoreModels('hairline', MODELS_BY_FAMILY.hairline!, (c) => c.a.inkPx);
const thinScores = scoreModels('thin', MODELS_BY_FAMILY.thin!, (c) => c.a.inkPx);
const borderScores = scoreModels('border', MODELS_BY_FAMILY.border!, (c) => c.a.inkPx);
const dashScores = scoreModels('dash', MODELS_BY_FAMILY.dash!, (c) => c.a.periodPx);
const markerScores = scoreModels(
  'marker',
  MODELS_BY_FAMILY.marker!,
  (c) => c.a.boxH,
  () => true,
  BOX_TOLERANCE_PX,
);
// Under 960 the tile is not drawn as a scaled tile, and a period under two pixels is not drawn at all.
const PATTERN_SCALED_FROM = 960;
const patternScores = scoreModels(
  'pattern',
  MODELS_BY_FAMILY.pattern!,
  (c) => c.a.periodPx,
  (c) => c.width >= PATTERN_SCALED_FROM && (c.probe.periodPt ?? 6) * c.scale >= 2,
);
// The inner edge is only located where the stroke is at least a pixel thick.
const borderEdgeScores = scoreModels(
  'border',
  BORDER_EDGE_MODELS,
  (c) => c.a.innerEdgePx,
  (c) => (c.probe.widthPt ?? 1) * c.scale >= 1,
);

const families: readonly [string, readonly Score[]][] = [
  ['hairline', hairlineScores],
  ['thin', thinScores],
  ['border', borderScores],
  ['dash', dashScores],
  ['marker', markerScores],
  ['pattern', patternScores],
  ['border edge', borderEdgeScores],
];
for (const [family, scores] of families) {
  console.log(
    `${family.padEnd(12)} ${scores.map((s) => `${s.model} ${String(s.fits)}/${String(s.of)}`).join('  ')}`,
  );
}
const inconclusive: string[] = [];
const pick = (family: string, scores: readonly Score[]): Score => {
  try {
    return winner(family, scores);
  } catch (error) {
    inconclusive.push(error instanceof Error ? error.message : String(error));
    return scores[0]!;
  }
};
const winners = {
  hairline: pick('hairline', hairlineScores),
  thin: pick('thin', thinScores),
  border: pick('border', borderScores),
  dash: pick('dash', dashScores),
  marker: pick('marker', markerScores),
  pattern: pick('pattern', patternScores),
  borderEdge: pick('border edge', borderEdgeScores),
};
if (inconclusive.length > 0) throw new Error(inconclusive.join('; '));

/* ------------------------------------------------------------ text, ramps */

/** Text: does an extent scale with the export? Compared against the 960 export. */
interface TextRow {
  readonly probe: string;
  readonly width: number;
  readonly boxW: number;
  readonly boxH: number;
  readonly inkSum: number;
  readonly linearW: number;
  readonly linearH: number;
  readonly withinPx: number;
}
const textRows: TextRow[] = [];
for (const probe of probes.filter((p) => p.family === 'text')) {
  const reference = cases.find((c) => c.probe === probe && c.width === 960);
  if (reference === undefined) continue;
  for (const c of cases.filter((k) => k.probe === probe)) {
    const linearW = (reference.a.boxW ?? 0) * c.scale;
    const linearH = (reference.a.boxH ?? 0) * c.scale;
    textRows.push({
      probe: probe.id,
      width: c.width,
      boxW: c.a.boxW ?? 0,
      boxH: c.a.boxH ?? 0,
      inkSum: Math.round((c.a.inkSum ?? 0) * 100) / 100,
      linearW: Math.round(linearW * 100) / 100,
      linearH: Math.round(linearH * 100) / 100,
      withinPx:
        Math.round(
          Math.max(Math.abs((c.a.boxW ?? 0) - linearW), Math.abs((c.a.boxH ?? 0) - linearH)) * 100,
        ) / 100,
    });
  }
}
const textTolerance = (row: TextRow): number => 1 + 0.02 * Math.max(row.linearW, row.linearH);
const textLinear = textRows.every((row) => row.withinPx <= textTolerance(row));
const textDroppedBelowPx = textRows
  .filter((row) => row.boxH === 0 && row.linearH >= 1)
  .reduce((worst, row) => Math.max(worst, row.linearH), 0);
const textWorstPx = textRows.reduce((worst, row) => Math.max(worst, row.withinPx), 0);

/** The ramp: the same curve at every width, against the 960 export. */
const rampReference = cases.find((c) => c.probe.id === 'gradient-lin' && c.width === 960);
const rampRows = cases
  .filter((c) => c.probe.id === 'gradient-lin')
  .map((c) => ({
    width: c.width,
    ramp: c.a.ramp ?? [],
    levels: c.a.levels ?? 0,
    worstLevels: Math.max(
      ...(c.a.ramp ?? []).map((v, i) => Math.abs(v - (rampReference?.a.ramp?.[i] ?? v))),
    ),
  }));
/** The smallest width from which every larger export agrees with 960 to one level. */
const gradientIdenticalFrom = [...rampRows]
  .sort((a, b) => a.width - b.width)
  .reduce<number | null>((from, row) => (row.worstLevels <= 1 ? (from ?? row.width) : null), null);

/** The pattern where its period is under two pixels: coverage against the 960 export. */
const patternMeanRows = cases
  .filter((c) => c.probe.family === 'pattern')
  .map((c) => {
    const reference = cases.find((k) => k.probe === c.probe && k.width === 960);
    return {
      probe: c.probe.id,
      width: c.width,
      periodPx: c.a.periodPx ?? null,
      mean: Math.round((c.a.mean ?? 0) * 1000) / 1000,
      referenceMean: Math.round((reference?.a.mean ?? 0) * 1000) / 1000,
    };
  });
/** A tile row in eight is nominal coverage 0.125; the smallest width from which every larger one is within 0.01 of it. */
const patternCoverageNominalFrom = patternMeanRows
  .filter((row) => row.probe === 'pattern-horz')
  .sort((a, b) => a.width - b.width)
  .reduce<number | null>(
    (from, row) => (Math.abs(row.mean - 0.125) <= 0.01 ? (from ?? row.width) : null),
    null,
  );
const patternBelow = patternMeanRows.filter((row) => row.width < PATTERN_SCALED_FROM);

const frameRows = cases
  .filter((c) => c.probe.family === 'frame')
  .map((c) => ({ width: c.width, bandRows: c.a.bandRows ?? 0, height: c.a.bitmapHeight ?? 0 }));
const frameStretched = frameRows.every((row) => row.bandRows === 0);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const byProbe: Record<string, Record<string, Measure>> = {};
for (const c of cases) {
  const rounded: Measure = Object.fromEntries(
    Object.entries(c.a).map(([key, value]) => [
      key,
      typeof value === 'number' ? Math.round(value * 1000) / 1000 : value,
    ]),
  );
  (byProbe[c.probe.id] ??= {})[String(c.width)] = rounded;
}

const fixture = {
  $comment:
    'GENERATED by tools/ground-truth/render/zoom/analyse.ts from PowerPoint exports of the F2 probe deck. Do not edit by hand.',
  experiment: 'F2',
  powerpoint: exported.powerpoint,
  slide: SLIDE,
  widths: EXPORT_WIDTHS,
  tolerancePx: TOLERANCE_PX,
  boxTolerancePx: BOX_TOLERANCE_PX,
  cases: cases.length,
  missing,
  excluded: [...excluded],
  findings: {
    hairline: winners.hairline.model,
    hairlineDescription: winners.hairline.description,
    thin: winners.thin.model,
    thinDescription: winners.thin.description,
    border: winners.border.model,
    borderEdge: winners.borderEdge.model,
    borderEdgeDescription: winners.borderEdge.description,
    dashOnHairline: winners.dash.model,
    dashOnHairlineDescription: winners.dash.description,
    markerOnHairline: winners.marker.model,
    markerOnHairlineDescription: winners.marker.description,
    pattern: winners.pattern.model,
    patternScaledFrom: PATTERN_SCALED_FROM,
    patternCoverageNominalFrom,
    patternBelow,
    textLinear,
    textTolerance: '1 px + 2 % of the linear extent',
    textWorstPx: Math.round(textWorstPx * 100) / 100,
    textDroppedBelowPx,
    gradientIdenticalFrom,
    gradientWorstLevels: Object.fromEntries(
      rampRows.map((row) => [String(row.width), row.worstLevels]),
    ),
    frameStretched,
    exportCeiling: exported.refused.length === 0 ? null : exported.refused,
  },
  candidates: {
    hairline: hairlineScores,
    thin: thinScores,
    border: borderScores,
    borderEdge: borderEdgeScores,
    dash: dashScores,
    marker: markerScores,
    pattern: patternScores,
  },
  text: textRows,
  gradient: rampRows,
  patternCoverage: patternMeanRows,
  frame: frameRows,
  probes: probes.map(({ markup: _markup, ...rest }) => rest),
  measured: byProbe,
};

writeFileSync(fixturePath, `${JSON.stringify(fixture, null, 2)}\n`);

console.log(
  `F2: ${String(cases.length)} cases, ${String(excluded.size)} excluded, PowerPoint ${exported.powerpoint}`,
);
for (const [family, score] of Object.entries(winners)) {
  console.log(
    `  ${family.padEnd(11)} ${score.model}  ${score.description}  (${String(score.fits)}/${String(score.of)})`,
  );
}
console.log(
  `  text        linear ${String(textLinear)} (worst ${textWorstPx.toFixed(2)} px, dropped below ${String(textDroppedBelowPx)} px)`,
);
console.log(
  `  gradient    the same curve to one level from ${String(gradientIdenticalFrom)} px wide`,
);
console.log(`  pattern     nominal coverage from ${String(patternCoverageNominalFrom)} px wide`);
console.log(`  frame       stretched ${String(frameStretched)}`);
console.log(`wrote ${fixturePath}`);

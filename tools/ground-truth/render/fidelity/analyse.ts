/**
 * Experiment F1, step 2 - turn PowerPoint's PNGs into the committed oracle.
 *
 * ```
 * pnpm build
 * node tools/ground-truth/render/fidelity/analyse.ts <work-dir> --capture [--only <id>[,<id>]]
 * ```
 *
 * `--capture` runs `read.ps1` first; without it the work directory is expected
 * to hold a capture already. `--only` captures the named decks and merges them
 * into the oracle the others already have, because a full capture re-rolls the
 * slides PowerPoint draws differently twice and every date field. What is
 * committed is the reduced grid, not the PNG: a quarter of the bytes, no image
 * decoder in CI, and nothing that has been through a compressor whose output is
 * not in any specification. ADR 0035, ADR 0054.
 */

import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { FidelityError } from '../../../fidelity/errors.ts';
import { encodeGridSet, shardGridSets } from '../../../fidelity/metric/grid.ts';
import { differenceOf, scoreOf } from '../../../fidelity/metric/score.ts';
import type { Grid } from '../../../fidelity/metric/reduce.ts';
import type { Oracle, OracleRecord, ZoomRecord } from '../../../fidelity/oracle.ts';
import { openHarness } from '../../../fidelity/raster/browser.ts';
import {
  CELL,
  geometryOf,
  injectHarness,
  oracleGrid,
  RASTER_WIDTH,
} from '../../../fidelity/raster/render.ts';
import { REPO_ROOT, repoPath } from '../../../repo/root.ts';

import { claimFixtures, FIXTURE_PREFIX, type FixtureEntry } from '../../../fidelity/fixtures.ts';

import { fidelityProbes, slideKey } from './probes.ts';

const dirArg = process.argv[2];
if (dirArg === undefined) {
  throw new Error('usage: analyse.ts <work-dir> [--capture] [--only <id>[,<id>]]');
}
const dir = dirArg;
const capture = process.argv.includes('--capture');
const onlyAt = process.argv.indexOf('--only');
const only =
  onlyAt === -1 ? null : new Set((process.argv[onlyAt + 1] ?? '').split(',').filter(Boolean));
if (only !== null && only.size === 0) throw new Error('--only names no deck');

/** Where the committed oracle lives. */
const FIXTURES = repoPath('corpus/ground-truth/render/fidelity');

/** The sub-phase that wrote a deck's oracle: Gate 3's deck, or the original capture. */
function addedIn(deck: string): string {
  return deck === 'a46-hundred-slides' ? '3.11' : '3.9';
}

interface ExportedSlide {
  readonly slide: number;
  readonly width: number;
  readonly height: number;
  readonly file: string;
  /** The second export of the same slide, kept so the two can be compared. */
  readonly again: string;
  readonly bytes: number;
  readonly sameBytes: boolean;
}

interface ExportedDeck {
  readonly id: string;
  readonly path: string;
  readonly slides: readonly ExportedSlide[];
  readonly error: string | null;
  readonly pointWidth?: number;
  readonly pointHeight?: number;
}

const allProbes = fidelityProbes();
const probes = only === null ? allProbes : allProbes.filter((p) => only.has(p.id));
if (only !== null && probes.length !== only.size) {
  const known = new Set(probes.map((p) => p.id));
  throw new FidelityError(
    'FID_ORACLE_MISSING',
    `--only names a deck that is not a probe: ${[...only].filter((id) => !known.has(id)).join(', ')}`,
  );
}

if (capture) {
  mkdirSync(dir, { recursive: true });
  writeFileSync(
    join(dir, 'oracle-inputs.json'),
    `${JSON.stringify(
      {
        repo: REPO_ROOT,
        rasterWidth: RASTER_WIDTH,
        decks: probes.map((p) => ({ id: p.id, path: p.path, zoomWidths: p.zoomWidths })),
      },
      null,
      2,
    )}\n`,
  );
  const script = repoPath('tools/ground-truth/render/fidelity/read.ps1');
  const run = spawnSync(
    'powershell',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-Dir', dir],
    {
      stdio: 'inherit',
    },
  );
  if (run.status !== 0) throw new Error(`read.ps1 exited ${String(run.status)}`);
}

const exported = JSON.parse(readFileSync(join(dir, 'oracle-export.json'), 'utf8')) as {
  readonly powerpoint: string;
  readonly decks: readonly ExportedDeck[];
};

/* ------------------------------------------------------- the PNGs themselves */

/**
 * The chunk types a PNG carries, headers only.
 *
 * Never decoded here: the point is to refuse a colour-managed input before
 * anything looks at its pixels, because Chromium converts a profiled PNG and
 * the conversion is a property of the build rather than of the file.
 */
function chunkTypes(bytes: Uint8Array, subject: string): readonly string[] {
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (signature.some((byte, at) => bytes[at] !== byte)) {
    throw new FidelityError('FID_ORACLE_COLOR_MANAGED', `${subject} is not a PNG`, subject);
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const types: string[] = [];
  let at = 8;
  while (at + 8 <= bytes.length) {
    const length = view.getUint32(at, false);
    const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    types.push(type);
    if (type === 'IEND') break;
    at += 12 + length;
  }
  return types;
}

const colourManaged: string[] = [];
const failures: string[] = [];

for (const deck of exported.decks) {
  if (deck.error !== null) {
    failures.push(`${deck.id}: ${deck.error}`);
    continue;
  }
  for (const slide of deck.slides) {
    const types = chunkTypes(new Uint8Array(readFileSync(join(dir, slide.file))), slide.file);
    if (types.includes('iCCP')) colourManaged.push(`${slide.file} carries iCCP`);
  }
}

if (failures.length > 0) {
  throw new FidelityError('FID_ORACLE_MISSING', `PowerPoint refused: ${failures.join('; ')}`);
}
if (colourManaged.length > 0) {
  throw new FidelityError('FID_ORACLE_COLOR_MANAGED', colourManaged.join('; '));
}

/* ------------------------------------------------------------- the reduction */

const harness = await openHarness();
await injectHarness(harness.page);

mkdirSync(join(FIXTURES, 'grids'), { recursive: true });
mkdirSync(join(FIXTURES, 'zoom'), { recursive: true });
const captured = new Set(exported.decks.map((deck) => deck.id));
for (const sub of ['grids', 'zoom']) {
  for (const stale of readdirSync(join(FIXTURES, sub))) {
    // A full capture starts clean; a partial one replaces only the decks it captured.
    if (only === null || [...captured].some((id) => stale.startsWith(`${id}.`))) {
      rmSync(join(FIXTURES, sub, stale));
    }
  }
}

const records: OracleRecord[] = [];
const zoom: ZoomRecord[] = [];
/** How far PowerPoint disagrees with itself, which bounds what the score means. */
const jitter: { key: string; meanBp: number; maxD: number; cells: number }[] = [];

/** A PNG on disk as a URL the page can fetch without widening its served roots. */
function dataUrl(file: string): string {
  return `data:image/png;base64,${readFileSync(join(dir, file)).toString('base64')}`;
}

const files: FixtureEntry[] = [];

/** Write one deck's grids at one width in as many sets as the cap needs, and name each. */
function writeSets(
  deck: string,
  width: number,
  grids: readonly { key: string; grid: Grid }[],
): string[] {
  const shards = shardGridSets(grids);
  const sub = width === RASTER_WIDTH ? 'grids' : 'zoom';
  const stem = width === RASTER_WIDTH ? deck : `${deck}.${String(width)}`;
  return shards.map((shard, index) => {
    const name =
      shards.length === 1 ? `${stem}.ppt.grids` : `${stem}.${String(index + 1)}.ppt.grids`;
    writeFileSync(join(FIXTURES, sub, name), encodeGridSet(shard));
    const path = `${FIXTURE_PREFIX}${sub}/${name}`;
    const first = shard[0]!.key.slice(-2);
    const last = shard[shard.length - 1]!.key.slice(-2);
    const which = shards.length === 1 ? 'every slide' : `slides ${first} to ${last}`;
    const at = width === RASTER_WIDTH ? '' : ` at ${String(width)} pixels wide`;
    files.push({
      id: `fidelity-${stem}${shards.length === 1 ? '' : `-${String(index + 1)}`}`,
      path,
      tags: ['fidelity', 'oracle', 'render', ...(width === RASTER_WIDTH ? [] : ['zoom'])],
      description:
        `PowerPoint's own rendering of ${which} of ${deck}${at}, as luma and decimated chroma ` +
        `cell means. ` +
        (width === RASTER_WIDTH
          ? "The reference half of sub-phase 3.9's score."
          : "The reference half of Gate 3's score at that zoom."),
      recipe: {
        tool: 'tools/ground-truth/render/fidelity/analyse.ts',
        args: ['<dir>', '--capture', '--only', deck],
      },
      addedIn: addedIn(deck),
    });
    return path;
  });
}

for (const deck of exported.decks) {
  const probe = probes.find((p) => p.id === deck.id);
  if (probe === undefined) {
    throw new FidelityError('FID_ORACLE_MISSING', `${deck.id} is not a probe`);
  }
  const widths = [RASTER_WIDTH, ...probe.zoomWidths];
  for (const width of widths) {
    const cell = (CELL * width) / RASTER_WIDTH;
    if (!Number.isInteger(cell)) {
      throw new FidelityError(
        'FID_RASTER_GEOMETRY',
        `${String(width)} px wide is not a whole number of ${String(CELL)} px cells per 960`,
      );
    }
    const geometry = geometryOf(deck.pointWidth ?? 1, deck.pointHeight ?? 1, { width, cell });
    const grids: { key: string; grid: Grid }[] = [];
    const slides = deck.slides.filter((slide) => slide.width === width);

    for (const slide of slides) {
      const key = slideKey(deck.id, slide.slide);
      const first = await oracleGrid(harness.page, dataUrl(slide.file), geometry);
      const second = await oracleGrid(harness.page, dataUrl(slide.again), geometry);
      if (
        first.natural.width !== geometry.drawWidth ||
        first.natural.height !== geometry.drawHeight
      ) {
        throw new FidelityError(
          'FID_RASTER_GEOMETRY',
          `${slide.file} is ${String(first.natural.width)}x${String(first.natural.height)}, ` +
            `not the ${String(geometry.drawWidth)}x${String(geometry.drawHeight)} asked for`,
          slide.file,
        );
      }
      // PowerPoint does not rasterise a slide identically twice - measured, 28 of
      // 149 - so this records how far it disagrees with itself rather than
      // asserting it does not. That disagreement is the noise floor of every
      // number scored against it, and a score is not reported without it.
      if (first.rasterSha256 !== second.rasterSha256 && width === RASTER_WIDTH) {
        const score = scoreOf(differenceOf(first.grid, second.grid));
        jitter.push({
          key,
          meanBp: score.meanBp,
          maxD: score.maxD,
          cells: score.cells - (score.hist[0] ?? 0),
        });
      }
      grids.push({ key, grid: first.grid });
    }

    const paths = writeSets(deck.id, width, grids);
    shardGridSets(grids).forEach((shard, index) => {
      for (const { key, grid } of shard) {
        const row = {
          key,
          deck: deck.id,
          slide: Number(key.slice(key.lastIndexOf('-') + 1)),
          cells: grid.width * grid.height,
          file: paths[index]!,
        };
        if (width === RASTER_WIDTH) records.push(row);
        else zoom.push({ ...row, width, cell });
      }
    });
  }
}

await harness.close();

/* ----------------------------------------------------------------- the index */

/** The decks a partial capture did not touch keep their rows; a row from before sharding is in its deck's one set. */
function mergeOracle(existing: Oracle | null): Oracle {
  if (existing !== null && existing.powerpoint !== exported.powerpoint) {
    throw new FidelityError(
      'FID_ORACLE_MISSING',
      `the oracle was captured on PowerPoint ${existing.powerpoint} and this run on ` +
        `${exported.powerpoint}; two builds are not one oracle`,
    );
  }
  const kept = (row: { deck: string }): boolean => !captured.has(row.deck);
  const keptKeys = new Set((existing?.records ?? []).filter(kept).map((row) => row.key));
  const mergedRecords = [
    ...(existing?.records ?? []).filter(kept).map((row) => ({
      key: row.key,
      deck: row.deck,
      slide: row.slide,
      cells: row.cells,
      file: `${FIXTURE_PREFIX}grids/${row.deck}.ppt.grids`,
    })),
    ...records,
  ];
  const mergedJitter = [
    ...(existing?.jitter ?? []).filter((row) => keptKeys.has(row.key)),
    ...jitter,
  ];
  const mergedZoom = [...(existing?.zoom ?? []).filter(kept), ...zoom];
  return {
    powerpoint: exported.powerpoint,
    rasterWidth: RASTER_WIDTH,
    slides: mergedRecords.length,
    noiseFloor: {
      slidesAffected: mergedJitter.length,
      worstMeanBpDrop: mergedJitter.reduce((worst, row) => Math.max(worst, 10000 - row.meanBp), 0),
      worstMaxD: mergedJitter.reduce((worst, row) => Math.max(worst, row.maxD), 0),
      worstCells: mergedJitter.reduce((worst, row) => Math.max(worst, row.cells), 0),
    },
    jitter: mergedJitter,
    records: mergedRecords,
    zoom: mergedZoom,
  };
}

const existing =
  only === null
    ? null
    : (JSON.parse(readFileSync(join(FIXTURES, 'oracle.json'), 'utf8')) as Oracle);
const oracle = mergeOracle(existing);
const totalBytes = [...new Set([...oracle.records, ...oracle.zoom].map((row) => row.file))].reduce(
  (sum, file) => sum + statSync(repoPath('corpus/ground-truth', file)).size,
  0,
);

writeFileSync(
  join(FIXTURES, 'oracle.json'),
  `${JSON.stringify(
    {
      $comment: 'GENERATED by tools/ground-truth/render/fidelity/analyse.ts. Do not edit by hand.',
      ...oracle,
      bytes: totalBytes,
    },
    null,
    2,
  )}\n`,
);

/* -------------------------------------------------------------- the manifest */

claimFixtures(
  [
    ...files,
    {
      id: 'fidelity-oracle',
      path: 'render/fidelity/oracle.json',
      tags: ['fidelity', 'oracle', 'render'],
      description:
        'The index of the fidelity oracle: which file holds each grid at each width, and how far ' +
        'PowerPoint disagreed with its own second rendering of the same slide, which is the noise ' +
        'floor of every number scored against it.',
      recipe: {
        tool: 'tools/ground-truth/render/fidelity/analyse.ts',
        args: ['<dir>', '--capture'],
      },
      addedIn: '3.9',
    },
  ],
  only === null ? FIXTURE_PREFIX + 'grids/' : undefined,
);

console.log(
  `oracle: ${String(oracle.records.length)} slide(s) at 960, ${String(oracle.zoom.length)} at a zoom, ` +
    `${(totalBytes / 1024 / 1024).toFixed(2)} MiB of grids, PowerPoint ${exported.powerpoint}` +
    (only === null ? '' : ` (captured ${[...captured].join(', ')})`),
);
console.log(
  `noise floor: ${String(oracle.noiseFloor.slidesAffected)} slide(s) PowerPoint drew differently twice, ` +
    `worst ${String(oracle.noiseFloor.worstMeanBpDrop)} bp / maxD ${String(oracle.noiseFloor.worstMaxD)}`,
);

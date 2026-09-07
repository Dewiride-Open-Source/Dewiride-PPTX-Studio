/**
 * Experiment F1, step 2 - turn PowerPoint's PNGs into the committed oracle.
 *
 * ```
 * pnpm build
 * node tools/ground-truth/render/fidelity/analyse.ts <work-dir> --capture
 * ```
 *
 * `--capture` runs `read.ps1` first; without it the work directory is expected
 * to hold a capture already. What is committed is the reduced grid, not the
 * PNG: a quarter of the bytes, no image decoder in CI, and nothing that has been
 * through a compressor whose output is not in any specification.
 *
 * The reduction is `tools/fidelity/metric/reduce.ts`, run in the browser on both
 * sides. There is deliberately no second implementation here - the one thing
 * that must not drift is the thing both sides are measured with. ADR 0035.
 */

import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { FidelityError } from '../../../fidelity/errors.ts';
import { encodeGridSet } from '../../../fidelity/metric/grid.ts';
import { differenceOf, scoreOf } from '../../../fidelity/metric/score.ts';
import type { Grid } from '../../../fidelity/metric/reduce.ts';
import { openHarness } from '../../../fidelity/raster/browser.ts';
import {
  geometryOf,
  injectReduce,
  oracleGrid,
  RASTER_WIDTH,
} from '../../../fidelity/raster/render.ts';
import { REPO_ROOT, repoPath } from '../../../repo/root.ts';

import { claimFixtures, FIXTURE_PREFIX } from '../../../fidelity/fixtures.ts';

import { fidelityProbes, slideKey } from './probes.ts';

const dirArg = process.argv[2];
if (dirArg === undefined) throw new Error('usage: analyse.ts <work-dir> [--capture]');
const dir = dirArg;
const capture = process.argv.includes('--capture');

/** Where the committed oracle lives. */
const FIXTURES = repoPath('corpus/ground-truth/render/fidelity');

interface ExportedSlide {
  readonly slide: number;
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
  readonly pixelHeight?: number;
}

const probes = fidelityProbes();

if (capture) {
  mkdirSync(dir, { recursive: true });
  writeFileSync(
    join(dir, 'oracle-inputs.json'),
    `${JSON.stringify(
      {
        repo: REPO_ROOT,
        rasterWidth: RASTER_WIDTH,
        decks: probes.map((p) => ({ id: p.id, path: p.path })),
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
let differingBytes = 0;

for (const deck of exported.decks) {
  if (deck.error !== null) {
    failures.push(`${deck.id}: ${deck.error}`);
    continue;
  }
  for (const slide of deck.slides) {
    if (!slide.sameBytes) differingBytes += 1;
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
await injectReduce(harness.page);

mkdirSync(join(FIXTURES, 'grids'), { recursive: true });
for (const stale of readdirSync(join(FIXTURES, 'grids'))) {
  rmSync(join(FIXTURES, 'grids', stale));
}

interface OracleRecord {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  readonly cells: number;
  readonly bytes: number;
}

const records: OracleRecord[] = [];
let totalBytes = 0;
/** How far PowerPoint disagrees with itself, which bounds what the score means. */
const jitter: { key: string; meanBp: number; maxD: number; cells: number }[] = [];

/** A PNG on disk as a URL the page can fetch without widening its served roots. */
function dataUrl(file: string): string {
  return `data:image/png;base64,${readFileSync(join(dir, file)).toString('base64')}`;
}

const files: { deck: string; path: string; bytes: Uint8Array }[] = [];

for (const deck of exported.decks) {
  const probe = probes.find((p) => p.id === deck.id);
  if (probe === undefined)
    throw new FidelityError('FID_ORACLE_MISSING', `${deck.id} is not a probe`);
  const geometry = geometryOf(deck.pointWidth ?? 1, deck.pointHeight ?? 1);
  const grids: { key: string; grid: Grid }[] = [];

  for (const slide of deck.slides) {
    const key = slideKey(deck.id, slide.slide);
    const first = await oracleGrid(harness.page, dataUrl(slide.file), geometry);
    const second = await oracleGrid(harness.page, dataUrl(slide.again), geometry);
    // PowerPoint does not rasterise a slide identically twice - measured, 28 of
    // 149 - so this records how far it disagrees with itself rather than
    // asserting it does not. That disagreement is the noise floor of every
    // number scored against it, and a score is not reported without it.
    if (first.rasterSha256 !== second.rasterSha256) {
      const score = scoreOf(differenceOf(first.grid, second.grid));
      jitter.push({
        key,
        meanBp: score.meanBp,
        maxD: score.maxD,
        cells: score.cells - (score.hist[0] ?? 0),
      });
    }
    grids.push({ key, grid: first.grid });
    records.push({
      key,
      deck: deck.id,
      slide: slide.slide,
      cells: first.grid.width * first.grid.height,
      bytes: 0,
    });
  }

  const packed = encodeGridSet(grids);
  const path = `render/fidelity/grids/${deck.id}.ppt.grids`;
  writeFileSync(join(FIXTURES, 'grids', `${deck.id}.ppt.grids`), packed);
  files.push({ deck: deck.id, path, bytes: packed });
  totalBytes += packed.length;
}

await harness.close();

/** The worst self-disagreement over the corpus, which is what the floor is. */
const noiseFloor = {
  slidesAffected: jitter.length,
  worstMeanBpDrop: jitter.reduce((worst, row) => Math.max(worst, 10000 - row.meanBp), 0),
  worstMaxD: jitter.reduce((worst, row) => Math.max(worst, row.maxD), 0),
  worstCells: jitter.reduce((worst, row) => Math.max(worst, row.cells), 0),
};

writeFileSync(
  join(FIXTURES, 'oracle.json'),
  `${JSON.stringify(
    {
      $comment: 'GENERATED by tools/ground-truth/render/fidelity/analyse.ts. Do not edit by hand.',
      powerpoint: exported.powerpoint,
      rasterWidth: RASTER_WIDTH,
      slides: records.length,
      bytes: totalBytes,
      bytesDifferedOnSecondExport: differingBytes,
      noiseFloor,
      jitter,
      records,
    },
    null,
    2,
  )}\n`,
);

/* -------------------------------------------------------------- the manifest */

claimFixtures(
  [
    ...files.map((file) => ({
      id: `fidelity-${file.deck}`,
      path: file.path,
      tags: ['fidelity', 'oracle', 'render'],
      description:
        `PowerPoint's own rendering of every slide of ${file.deck}, as luma and decimated chroma ` +
        `cell means at ${String(RASTER_WIDTH)} pixels wide. The reference half of sub-phase 3.9's score.`,
      recipe: {
        tool: 'tools/ground-truth/render/fidelity/analyse.ts',
        args: ['<dir>', '--capture'],
      },
    })),
    {
      id: 'fidelity-oracle',
      path: 'render/fidelity/oracle.json',
      tags: ['fidelity', 'oracle', 'render'],
      description:
        'The index of the fidelity oracle: which slide each grid belongs to, and how far PowerPoint ' +
        'disagreed with its own second rendering of the same slide, which is the noise floor of ' +
        'every number scored against it.',
      recipe: {
        tool: 'tools/ground-truth/render/fidelity/analyse.ts',
        args: ['<dir>', '--capture'],
      },
    },
  ],
  FIXTURE_PREFIX + 'grids/',
);

console.log(
  `oracle: ${String(records.length)} slide(s) in ${String(files.length)} file(s), ` +
    `${(totalBytes / 1024 / 1024).toFixed(2)} MiB of grids, PowerPoint ${exported.powerpoint}`,
);
console.log(
  `noise floor: ${String(noiseFloor.slidesAffected)} slide(s) PowerPoint drew differently twice, ` +
    `worst ${String(noiseFloor.worstMeanBpDrop)} bp / maxD ${String(noiseFloor.worstMaxD)}`,
);

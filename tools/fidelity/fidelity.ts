/**
 * The fidelity harness.
 *
 * ```
 * pnpm fidelity            score the corpus and check the gate
 * pnpm fidelity --record   write this machine's expected digests and font lock
 * ```
 *
 * Two questions, kept apart because only one of them has an exact answer.
 * **Accuracy** is how far our raster is from PowerPoint's; it is reported and
 * never gated, because the two differ everywhere and any pass/fail on it would
 * need a tolerance. **Regression** is whether our own raster changed; that is
 * gated, on a digest, with no tolerance to loosen. ADR 0035.
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { fidelityProbes, slideKey } from '../ground-truth/render/fidelity/probes.ts';
import { repoPath } from '../repo/root.ts';

import { FidelityError } from './errors.ts';
import { claimFixtures } from './fixtures.ts';
import { decodeGridSet } from './metric/grid.ts';
import type { Grid } from './metric/reduce.ts';
import { differenceOf, scoreOf } from './metric/score.ts';
import { openHarness } from './raster/browser.ts';
import {
  assertSameEnvironment,
  familiesDrawn,
  probeEnvironment,
  type Environment,
} from './raster/fonts.ts';
import { injectReduce, renderSlide, type SlideRaster } from './raster/render.ts';
import {
  corpusMeanBp,
  reportMarkdown,
  scoresJson,
  type RunReport,
  type SlideResult,
} from './report.ts';

const FIXTURES = repoPath('corpus/ground-truth/render/fidelity');
const OUT = repoPath('fidelity');

/** Which recorded lock and digest set this machine is compared against. */
const ENV_ID = `${process.platform}-${process.arch}`;

const record = process.argv.includes('--record');

interface Expected {
  readonly environment: Environment;
  readonly slides: Readonly<Record<string, { rasterSha256: string; svgSha256: string }>>;
}

interface Oracle {
  readonly noiseFloor: { worstMeanBpDrop: number; worstMaxD: number };
  readonly records: readonly { key: string }[];
}

const oracle = JSON.parse(readFileSync(join(FIXTURES, 'oracle.json'), 'utf8')) as Oracle;
const oracleKeys = new Set(oracle.records.map((row) => row.key));

/* --------------------------------------------------------------- the render */

const harness = await openHarness();
await injectReduce(harness.page);

const rendered = new Map<string, SlideRaster>();
const unsupported: { key: string; reason: string }[] = [];

for (const probe of fidelityProbes()) {
  for (let slide = 1; slide <= probe.slides; slide++) {
    const key = slideKey(probe.id, slide);
    if (!oracleKeys.has(key)) {
      throw new FidelityError('FID_ORACLE_MISSING', `no committed oracle grid for ${key}`, key);
    }
    try {
      rendered.set(key, await renderSlide(harness.page, `/${probe.path}`, slide - 1));
    } catch (error) {
      // A direction the renderer refuses is a recorded gap, not a score of
      // zero: scoring a slide we did not draw would report a number for
      // something that never happened.
      const message = error instanceof Error ? error.message : String(error);
      unsupported.push({ key, reason: (message.split('\n')[0] ?? message).trim() });
    }
  }
}

const families = familiesDrawn([...rendered.values()]);
const environment = await probeEnvironment(harness.page, families, ENV_ID);
await harness.close();

/* ---------------------------------------------------------- the environment */

const expectedPath = join(FIXTURES, `expected.${ENV_ID}.json`);
let expected: Expected | null = null;
try {
  expected = JSON.parse(readFileSync(expectedPath, 'utf8')) as Expected;
} catch {
  if (!record) {
    throw new FidelityError(
      'FID_ENV_UNKNOWN',
      `no expected.${ENV_ID}.json; this machine has never recorded a baseline`,
      ENV_ID,
    );
  }
}

if (expected !== null) assertSameEnvironment(environment, expected.environment);

/* ------------------------------------------------------------- the scoring */

const oracleGrids = new Map<string, ReturnType<typeof decodeGridSet>>();
function oracleGridFor(deck: string, key: string): Grid {
  let set = oracleGrids.get(deck);
  if (set === undefined) {
    const file = `${deck}.ppt.grids`;
    set = decodeGridSet(new Uint8Array(readFileSync(join(FIXTURES, 'grids', file))), file);
    oracleGrids.set(deck, set);
  }
  const grid = set.get(key);
  if (grid === undefined)
    throw new FidelityError('FID_ORACLE_MISSING', `no oracle grid for ${key}`, key);
  return grid;
}

const results: SlideResult[] = [];
for (const [key, raster] of rendered) {
  const deck = key.slice(0, key.lastIndexOf('-'));
  const score = scoreOf(differenceOf(raster.grid, oracleGridFor(deck, key)));
  const was = expected?.slides[key];
  results.push({
    key,
    deck,
    slide: Number(key.slice(key.lastIndexOf('-') + 1)),
    score,
    rasterSha256: raster.rasterSha256,
    svgSha256: raster.svgSha256,
    changed: was !== undefined && was.rasterSha256 !== raster.rasterSha256,
  });
}

const report: RunReport = {
  envId: ENV_ID,
  slides: results,
  noiseFloorBp: oracle.noiseFloor.worstMeanBpDrop,
  noiseFloorMaxD: oracle.noiseFloor.worstMaxD,
  substituted: environment.faces.filter((face) => !face.available).map((face) => face.family),
};

mkdirSync(OUT, { recursive: true });
writeFileSync(join(OUT, 'scores.json'), scoresJson(report));
writeFileSync(join(OUT, 'report.md'), reportMarkdown(report));

if (record) {
  writeFileSync(
    expectedPath,
    `${JSON.stringify(
      {
        $comment: 'GENERATED by tools/fidelity/fidelity.ts --record. Do not edit by hand.',
        environment,
        slides: Object.fromEntries(
          [...results]
            .sort((a, b) => (a.key < b.key ? -1 : 1))
            .map((r) => [r.key, { rasterSha256: r.rasterSha256, svgSha256: r.svgSha256 }]),
        ),
      },
      null,
      2,
    )}\n`,
  );
  claimFixtures([
    {
      id: `fidelity-expected-${ENV_ID}`,
      path: `render/fidelity/expected.${ENV_ID}.json`,
      tags: ['fidelity', 'baseline', 'render'],
      description:
        `The digest of our own raster for every slide on ${ENV_ID}, and the font environment it ` +
        'was recorded in. This is the gate: a slide that rasterises differently here has changed, ' +
        'and no tolerance is involved in saying so.',
      recipe: { tool: 'tools/fidelity/fidelity.ts', args: ['--record'] },
    },
  ]);
}

/* -------------------------------------------------------------- the verdict */

for (const gap of unsupported) console.log(`  - ${gap.key}: ${gap.reason}`);
console.log(
  `fidelity: ${String(results.length)} slide(s) scored, ` +
    `${String(unsupported.length)} not drawn, ` +
    `corpus mean ${String(corpusMeanBp(results))} bp ` +
    `(oracle noise floor ${String(report.noiseFloorBp)} bp)`,
);

const changed = results.filter((slide) => slide.changed);
if (changed.length > 0) {
  throw new FidelityError(
    'FID_RENDER_CHANGED',
    `${String(changed.length)} slide(s) rasterise differently than recorded: ` +
      changed.map((slide) => slide.key).join(', '),
  );
}

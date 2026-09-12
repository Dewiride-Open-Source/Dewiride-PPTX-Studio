/**
 * The fidelity harness.
 *
 * ```
 * pnpm fidelity          score the corpus and check the gate
 * pnpm fidelity --record  write this machine's first baseline and font lock
 * pnpm fidelity --record --why '<reason>' --expect <n>   rewrite one that exists
 * ```
 *
 * Two questions, kept apart because only one of them has an exact answer.
 * **Accuracy** is how far our raster is from PowerPoint's; it is reported and
 * never gated, because the two differ everywhere and any pass/fail on it would
 * need a tolerance. **Regression** is whether our own raster changed; that is
 * gated, on a digest, with no tolerance to loosen. ADR 0035.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { fidelityProbes, slideKey } from '../ground-truth/render/fidelity/probes.ts';
import { repoPath } from '../repo/root.ts';

import { blameOf } from './blame.ts';
import { FidelityError } from './errors.ts';
import { claimFixtures } from './fixtures.ts';
import { differenceOf, scoreOf } from './metric/score.ts';
import { openOracle } from './oracle.ts';
import { openHarness, servedUrl } from './raster/browser.ts';
import {
  assertSameEnvironment,
  currentEnvId,
  familiesDrawn,
  probeEnvironment,
} from './raster/fonts.ts';
import { injectHarness, renderSlide, type SlideRaster } from './raster/render.ts';
import {
  assertBaselineCovers,
  assertExpectedCount,
  assertMayRecord,
  baselineGaps,
  baselineJson,
  parseRecordArgs,
  readBaseline,
} from './record.ts';
import {
  corpusMeanBp,
  reportMarkdown,
  scoresJson,
  type RunReport,
  type SlideResult,
} from './report.ts';

const FIXTURES = repoPath('corpus/ground-truth/render/fidelity');
const OUT = repoPath('fidelity');

const ENV_ID = currentEnvId();

const args = parseRecordArgs(process.argv.slice(2));
const record = args.record;

const oracle = openOracle(FIXTURES);

/* ------------------------------------------ the baseline, and leave to write it */

const expectedPath = join(FIXTURES, `expected.${ENV_ID}.json`);
const expected = readBaseline(expectedPath);
if (expected === null && !record) {
  throw new FidelityError(
    'FID_ENV_UNKNOWN',
    `no expected.${ENV_ID}.json; this machine has never recorded a baseline`,
    ENV_ID,
  );
}

assertMayRecord(args, { inCi: process.env['CI'] !== undefined, hasBaseline: expected !== null });

/* --------------------------------------------------------------- the render */

const harness = await openHarness();
await injectHarness(harness.page);

const rendered = new Map<string, SlideRaster>();
const unsupported: { key: string; reason: string }[] = [];

for (const probe of fidelityProbes()) {
  for (let slide = 1; slide <= probe.slides; slide++) {
    const key = slideKey(probe.id, slide);
    if (!oracle.keys.has(key)) {
      throw new FidelityError('FID_ORACLE_MISSING', `no committed oracle grid for ${key}`, key);
    }
    try {
      const raster = await renderSlide(harness.page, servedUrl(probe.path), slide - 1);
      // A baseline is only worth recording if what it records is stable, so the
      // digest is taken twice on the one run that writes it down.
      if (record) {
        const again = await renderSlide(harness.page, servedUrl(probe.path), slide - 1);
        if (again.rasterSha256 !== raster.rasterSha256) {
          throw new FidelityError(
            'FID_RASTER_NONDETERMINISTIC',
            `${key} rasterised twice in one run and did not agree with itself`,
            key,
          );
        }
      }
      rendered.set(key, raster);
    } catch (error) {
      if (error instanceof FidelityError) throw error;
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

const gaps = baselineGaps(Object.keys(expected?.slides ?? {}), rendered.keys());
if (expected !== null) {
  assertSameEnvironment(environment, expected.environment);
  if (!record) assertBaselineCovers(gaps);
}

/* ------------------------------------------------------------- the scoring */

const results: SlideResult[] = [];
for (const [key, raster] of rendered) {
  const deck = key.slice(0, key.lastIndexOf('-'));
  const difference = differenceOf(raster.grid, oracle.gridFor(key));
  const score = scoreOf(difference);
  const was = expected?.slides[key];
  results.push({
    key,
    deck,
    slide: Number(key.slice(key.lastIndexOf('-') + 1)),
    score,
    rasterSha256: raster.rasterSha256,
    svgSha256: raster.svgSha256,
    changed: was !== undefined && was.rasterSha256 !== raster.rasterSha256,
    blame: blameOf(
      difference,
      raster.grid.width,
      raster.grid.height,
      raster.grid.cell,
      raster.shapes,
    ),
  });
}

const changed = results.filter((slide) => slide.changed);

const report: RunReport = {
  envId: ENV_ID,
  slides: results,
  noiseFloorBp: oracle.oracle.noiseFloor.worstMeanBpDrop,
  noiseFloorMaxD: oracle.oracle.noiseFloor.worstMaxD,
  substituted: environment.faces.filter((face) => !face.available).map((face) => face.family),
};

mkdirSync(OUT, { recursive: true });
writeFileSync(join(OUT, 'scores.json'), scoresJson(report));
writeFileSync(join(OUT, 'report.md'), reportMarkdown(report));

if (record) {
  assertExpectedCount(args, changed.length + gaps.unrecorded.length + gaps.vanished.length);
  writeFileSync(
    expectedPath,
    baselineJson(
      {
        environment,
        slides: Object.fromEntries(
          results.map((r) => [r.key, { rasterSha256: r.rasterSha256, svgSha256: r.svgSha256 }]),
        ),
      },
      'tools/fidelity/fidelity.ts',
    ),
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
      // The flags that decide the output, not the justification for it.
      recipe: {
        tool: 'tools/fidelity/fidelity.ts',
        args: args.bootstrap ? ['--record', '--bootstrap'] : ['--record'],
      },
      addedIn: '3.9',
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

if (record) {
  const rewritten = [...changed.map((slide) => slide.key), ...gaps.unrecorded, ...gaps.vanished];
  console.log(
    `fidelity: recorded ${expectedPath}` +
      (rewritten.length === 0
        ? ''
        : `, moving ${String(rewritten.length)}: ${rewritten.join(', ')}`),
  );
} else if (gaps.vanished.length > 0) {
  throw new FidelityError(
    'FID_RENDER_CHANGED',
    `${String(gaps.vanished.length)} recorded slide(s) are no longer drawn: ` +
      gaps.vanished.join(', '),
  );
} else if (changed.length > 0) {
  throw new FidelityError(
    'FID_RENDER_CHANGED',
    `${String(changed.length)} slide(s) rasterise differently than recorded: ` +
      changed.map((slide) => slide.key).join(', '),
  );
}

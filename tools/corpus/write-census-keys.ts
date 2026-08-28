#!/usr/bin/env node
/**
 * Regenerate `census-keys.gen.ts` from the census's own feature table.
 *
 * ```
 * pnpm build && node tools/corpus/write-census-keys.ts
 * ```
 *
 * Needs the build, which is why the output is committed rather than computed:
 * `pnpm corpus` has to run on a bare clone, before anything is built, so that
 * a contributor who drops an unlicensed deck into `corpus/` finds out on the
 * cheapest possible gate rather than four minutes into one.
 *
 * `census-drift.test.ts` fails if this was not re-run when it should have been.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const DIST = resolve(ROOT, 'packages/census/dist/index.js');
const OUT = resolve(ROOT, 'tools/corpus/census-keys.gen.ts');

interface Rule {
  readonly key: string;
}

const census = (await import(DIST)) as {
  FEATURE_RULES: readonly Rule[];
  PART_FEATURE_RULES: readonly Rule[];
};

const keys = [...census.FEATURE_RULES, ...census.PART_FEATURE_RULES].map((rule) => rule.key).sort();

const text =
  `/* eslint-disable */
// Generated from @pptx-studio/census. Do not edit.
//
// The Node-side corpus checker runs on a bare clone with no build, so it cannot
// import the census package - nothing links workspace packages into tools/, and
// the resolution that lets tools/bench/bench.test.ts do it comes from Vitest's
// alias table, not from Node. So the key list is committed, and
// census-drift.test.ts asserts it still matches the census's own table.
//
// Regenerate: pnpm build && node tools/corpus/write-census-keys.ts

export const CENSUS_FEATURE_KEYS: readonly string[] = [
` +
  keys.map((key) => `  '${key}',`).join('\n') +
  '\n];\n';

const before = (() => {
  try {
    return readFileSync(OUT, 'utf8');
  } catch {
    return '';
  }
})();

writeFileSync(OUT, text);
console.log(
  'census-keys.gen.ts: ' + String(keys.length) + ' keys' + (text === before ? ' (unchanged)' : ''),
);

#!/usr/bin/env node
/**
 * Generate `packages/geometry/src/presets/*.gen.ts` from Apache POI's
 * `presetShapeDefinitions.xml`.
 *
 * ## Getting the input
 *
 * The file is not in this repository and is not redistributed by it. It is
 * Apache-2.0 and lives in POI's source tree:
 *
 * ```
 * curl -LO https://raw.githubusercontent.com/apache/poi/trunk/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml
 * node tools/geometry-codegen/generate.ts --presets presetShapeDefinitions.xml --write
 * ```
 *
 * There is no formatting step after it: `*.gen.ts` is in `.prettierignore` and
 * in eslint's ignores, as `schema-order.gen.ts` is.
 *
 * Only the generated files are committed, and they record the SHA-256 of the
 * input, so "which POI is this geometry from" has an answer that does not
 * depend on anyone's memory. The `NOTICE` entry is the other half of the
 * obligation and is already in place.
 *
 * Run without `--write` for the report alone, which is what the ADR quotes.
 *
 * ## Why POI rather than the specification
 *
 * ECMA-376 describes the preset geometries in prose across several hundred
 * pages and ships no machine-readable form of them. Every open implementation
 * that draws presets - POI, LibreOffice, ONLYOFFICE - carries its own
 * transcription. POI's is Apache-2.0, which is the licence this project
 * publishes under, and it is a single self-describing file rather than a
 * generator embedded in a rendering stack.
 *
 * ## What this refuses to do
 *
 * It does not repair the input. Two things in the file are outside ECMA's
 * grammar - eight `+-` formulas with a fourth operand, and six formulas
 * containing a double space - and both are carried through and reported rather
 * than tidied. A transcoder that silently corrected its source would be making
 * semantic decisions that belong to the evaluator in 2.2, and would hide the
 * fact that the correction was ever needed.
 */

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';

import { REPO_ROOT as ROOT } from '../repo/root.ts';
import type { PresetBucketName, PresetShape } from '../../packages/geometry/src/types.ts';
import { encodeBucket } from './encode.ts';
import { readPresets, ROOT_ELEMENT, type PresetAnomaly } from './read-presets.ts';
import { BUCKET_NAMES, bucketOf, ST_SHAPE_TYPE } from './shape-types.ts';

const OUT = 'packages/geometry/src/presets';

interface Options {
  readonly presets: string;
  readonly out: string;
  readonly write: boolean;
}

function parseArgs(argv: readonly string[]): Options {
  let presets = '';
  let out = OUT;
  let write = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--presets') presets = argv[++i] ?? '';
    else if (arg === '--out') out = argv[++i] ?? OUT;
    else if (arg === '--write') write = true;
    else throw new Error(`unknown argument ${String(arg)}`);
  }
  if (presets === '') {
    throw new Error(
      '--presets <file> is required: POI’s presetShapeDefinitions.xml. See the header.',
    );
  }
  return { presets, out, write };
}

const options = parseArgs(process.argv.slice(2));
const bytes = readFileSync(resolve(ROOT, options.presets));
const sha256 = createHash('sha256').update(bytes).digest('hex');
const source = readPresets(bytes.toString('utf8'));

// ---------------------------------------------------------------- the two sets

const found = source.shapes.map((shape) => shape.name);
const expected = new Set(ST_SHAPE_TYPE);
const actual = new Set(found);
const missing = [...expected].filter((name) => !actual.has(name)).sort();
const extra = [...actual].filter((name) => !expected.has(name)).sort();

// ----------------------------------------------------------------- the buckets

const buckets = new Map<PresetBucketName, PresetShape[]>(BUCKET_NAMES.map((name) => [name, []]));
for (const shape of source.shapes) {
  const bucket = buckets.get(bucketOf(shape.name));
  if (bucket === undefined) throw new Error(`no bucket for ${shape.name}`);
  bucket.push(shape);
}

const assigned = [...buckets.values()].reduce((total, list) => total + list.length, 0);
if (assigned !== source.shapes.length) {
  throw new Error(`bucketing lost shapes: ${String(assigned)} of ${String(source.shapes.length)}`);
}

// ------------------------------------------------------------------ the totals

function count<T>(pick: (shape: PresetShape) => readonly T[]): number {
  return source.shapes.reduce((total, shape) => total + pick(shape).length, 0);
}

const totals = {
  shapes: source.shapes.length,
  avGuides: count((s) => s.avLst),
  gdGuides: count((s) => s.gdLst),
  handles: count((s) => s.ahLst),
  sites: count((s) => s.cxnLst),
  textRects: source.shapes.filter((s) => s.rect !== null).length,
  paths: count((s) => s.pathLst),
  commands: source.shapes.reduce(
    (total, shape) =>
      total + shape.pathLst.reduce((inner, path) => inner + path.commands.length, 0),
    0,
  ),
};

// ------------------------------------------------------------------- emitting

const HEADER = (what: string): string =>
  [
    '/* eslint-disable */',
    `// Generated by tools/geometry-codegen/generate.ts. Do not edit.`,
    '//',
    `// ${what}`,
    '//',
    "// Source: Apache POI's presetShapeDefinitions.xml, Apache-2.0, Copyright The",
    '// Apache Software Foundation. The file itself is not redistributed here; these',
    '// are the shape definitions transcoded out of it. See NOTICE.',
    '//',
    `//   ${basename(options.presets)}   ${String(bytes.byteLength)} bytes`,
    `//   sha256 ${sha256}`,
    '//',
    '// Encoding: see packages/geometry/src/decode.ts, which is the only reader.',
    '',
  ].join('\n');

function bucketModule(name: PresetBucketName, shapes: readonly PresetShape[]): string {
  const data = encodeBucket(shapes);
  const constant = name.toUpperCase();
  return [
    HEADER(`${String(shapes.length)} presets, ${String(data.length)} encoded characters.`),
    "import { PresetBucket } from '../decode.js';",
    '',
    `/** ${shapes.map((s) => s.name).join(', ')} */`,
    `export const ${constant}_DATA: string =`,
    `  ${JSON.stringify(data)};`,
    '',
    `export const ${constant}: PresetBucket = new PresetBucket(${constant}_DATA);`,
    '',
  ].join('\n');
}

function anomalyList(kind: PresetAnomaly['kind']): string {
  const items = source.anomalies.filter((a) => a.kind === kind);
  return JSON.stringify(
    items.map((a) => ({ shape: a.shape, guide: a.guide, fmla: a.fmla })),
    null,
    2,
  );
}

function manifestModule(): string {
  const members = Object.fromEntries(
    [...buckets].map(([name, shapes]) => [name, shapes.map((s) => s.name)]),
  );
  return [
    HEADER('Provenance, totals, and the two things in the source that ECMA does not describe.'),
    'export const PRESET_SOURCE = {',
    `  file: ${JSON.stringify(basename(options.presets))},`,
    `  sha256: ${JSON.stringify(sha256)},`,
    `  bytes: ${String(bytes.byteLength)},`,
    `  root: ${JSON.stringify(ROOT_ELEMENT)},`,
    '} as const;',
    '',
    '/** Counted on the generator side, from the parse. The suite checks the decode against these. */',
    `export const PRESET_TOTALS = ${JSON.stringify(totals, null, 2)} as const;`,
    '',
    `export const BUCKET_MEMBERS = ${JSON.stringify(members, null, 2)} as const;`,
    '',
    '/**',
    " * `+-` takes three operands in ECMA. Eight formulas in POI's file carry four,",
    ' * every one of them ending in a spare `0`, and every one of them in the',
    ' * circular-arrow family. Carried through verbatim: the evaluator decides what',
    ' * to do with a fourth operand, not the transcoder.',
    ' */',
    `export const OVER_LONG_FORMULAS = ${anomalyList('over-long-formula')} as const;`,
    '',
    '/**',
    ' * Formulas written with a double space. Split on runs of whitespace, so the',
    ' * operands are right; listed because a naive single-space split gives each of',
    ' * these an empty operand and an arity one too high.',
    ' */',
    `export const COLLAPSED_WHITESPACE = ${anomalyList('collapsed-whitespace')} as const;`,
    '',
  ].join('\n');
}

// -------------------------------------------------------------------- report

const lines: string[] = [];
lines.push('geometry-codegen');
lines.push(`  input     ${basename(options.presets)}  ${String(bytes.byteLength)} bytes`);
lines.push(`  sha256    ${sha256}`);
lines.push(`  root      <${ROOT_ELEMENT}>`);
lines.push('');
lines.push(`  shapes    ${String(totals.shapes)}`);
lines.push(
  `  guides    ${String(totals.avGuides)} adjust + ${String(totals.gdGuides)} computed = ${String(totals.avGuides + totals.gdGuides)}`,
);
lines.push(`  handles   ${String(totals.handles)}`);
lines.push(`  sites     ${String(totals.sites)}`);
lines.push(
  `  textRect  ${String(totals.textRects)} of ${String(totals.shapes)} (${source.shapes
    .filter((s) => s.rect === null)
    .map((s) => s.name)
    .join(', ')} have none)`,
);
lines.push(`  paths     ${String(totals.paths)}, ${String(totals.commands)} commands`);
lines.push('');
lines.push('  ST_ShapeType cross-check, against the list in shape-types.ts:');
lines.push(`    expected  ${String(expected.size)}`);
lines.push(`    found     ${String(actual.size)}`);
lines.push(`    missing   ${missing.length === 0 ? 'none' : missing.join(', ')}`);
lines.push(`    extra     ${extra.length === 0 ? 'none' : extra.join(', ')}`);
lines.push('');
lines.push('  buckets:');
for (const [name, shapes] of buckets) {
  const data = encodeBucket(shapes);
  lines.push(
    `    ${name.padEnd(10)} ${String(shapes.length).padStart(3)} presets  ${String(data.length).padStart(6)} chars`,
  );
}
lines.push('');
lines.push('  outside the grammar, carried through and not repaired:');
for (const anomaly of source.anomalies) {
  lines.push(`    ${anomaly.kind.padEnd(21)} ${anomaly.shape}/${anomaly.guide}  "${anomaly.fmla}"`);
}
console.log(lines.join('\n'));

if (missing.length > 0 || extra.length > 0) {
  console.error(
    '\ngeometry-codegen: the file and shape-types.ts disagree about ST_ShapeType.\n' +
      'That is a finding. Work out which source is wrong before touching either list -\n' +
      'pasting one over the other turns the check back into a tautology.',
  );
  process.exit(1);
}

if (!options.write) {
  console.log('\n(report only; pass --write to emit)');
  process.exit(0);
}

const outDir = resolve(ROOT, options.out);
mkdirSync(outDir, { recursive: true });
for (const [name, shapes] of buckets) {
  writeFileSync(join(outDir, `${name}.gen.ts`), bucketModule(name, shapes));
}
writeFileSync(join(outDir, 'manifest.gen.ts'), manifestModule());
// No formatting step to follow: `*.gen.ts` is in .prettierignore and in
// eslint.config.mjs's ignores, exactly as `schema-order.gen.ts` is. Generated
// files are checked by the suite that decodes them, not by a style pass.
console.log(`\nwrote ${String(buckets.size + 1)} files to ${options.out}`);

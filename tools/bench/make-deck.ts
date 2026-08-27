#!/usr/bin/env node
/**
 * Build a benchmark deck.
 *
 * ```
 * node tools/bench/make-deck.ts media-200mb  out/deck.pptx
 * node tools/bench/make-deck.ts xml-heavy    out/xml.pptx
 * node tools/bench/make-deck.ts small        out/small.pptx --slides 12
 * ```
 *
 * Deterministic. The same recipe writes the same bytes, which is what lets
 * `corpus/bench/manifest.json` pin a SHA-256 against a file far too large to
 * commit: anyone can rebuild it and check they have the same deck.
 */

import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { RECIPES, writeDeck, type DeckRecipe } from './deck.ts';
import { ZipStream } from './zip-stream.ts';

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');

/**
 * The CC0 probe font sub-phase 0.7 authored and committed.
 *
 * Reused rather than regenerated: it is already in the repository, it is
 * already licensed clearly, and embedding a font somebody else owns in a file
 * we hand around would be a licensing question we have no need to ask.
 */
const PROBE_FONT = resolve(ROOT, 'corpus/ground-truth/fonts/probe-v2_2.eot');

/** `DeckRecipe` is readonly by design; the flags need somewhere to write. */
type RecipeOverrides = { -readonly [K in keyof DeckRecipe]?: DeckRecipe[K] };

function parse(argv: readonly string[]): {
  recipe: DeckRecipe;
  out: string;
  overrides: RecipeOverrides;
} {
  const [name, out, ...rest] = argv;
  if (name === undefined || out === undefined) {
    throw new Error(
      'usage: make-deck.ts <' +
        Object.keys(RECIPES).join('|') +
        '> <out.pptx> [--slides N] [--no-font]',
    );
  }
  const recipe = RECIPES[name];
  if (recipe === undefined) {
    throw new Error('unknown recipe ' + name + '; try one of ' + Object.keys(RECIPES).join(', '));
  }

  const overrides: RecipeOverrides = {};
  for (let i = 0; i < rest.length; i++) {
    const flag = rest[i];
    if (flag === '--slides') overrides.slides = Number.parseInt(rest[++i] ?? '', 10);
    else if (flag === '--image-bytes') overrides.imageBytes = Number.parseInt(rest[++i] ?? '', 10);
    else if (flag === '--shapes') overrides.shapesPerSlide = Number.parseInt(rest[++i] ?? '', 10);
    else if (flag === '--no-font') overrides.embedFont = false;
    else if (flag === '--no-media') overrides.imageBytes = 0;
    else throw new Error('unknown flag ' + String(flag));
  }
  return { recipe, out, overrides };
}

const { recipe, out, overrides } = parse(process.argv.slice(2));
const effective: DeckRecipe = { ...recipe, ...overrides };

mkdirSync(dirname(resolve(out)), { recursive: true });

const startedAt = Date.now();
const zip = new ZipStream(resolve(out));
const font = effective.embedFont && existsSync(PROBE_FONT) ? PROBE_FONT : null;
if (effective.embedFont && font === null) {
  console.warn('probe font not found at ' + PROBE_FONT + '; the deck will have no embedded font');
}
const summary = writeDeck(zip, effective, font);
const elapsed = Date.now() - startedAt;

const hash = createHash('sha256')
  .update(readFileSync(resolve(out)))
  .digest('hex');
const size = statSync(resolve(out)).size;

console.log('recipe        ' + effective.name);
console.log('slides        ' + String(effective.slides));
console.log('entries       ' + String(summary.entries));
console.log('bytes         ' + String(size) + '  (' + (size / 1024 / 1024).toFixed(1) + ' MiB)');
if (summary.imagePixels > 0) {
  console.log(
    'images        ' +
      String(summary.imagePixels) +
      'x' +
      String(summary.imagePixels) +
      ' noise PNG',
  );
}
console.log(
  'shapes        ' +
    String(summary.slideShapes) +
    ' across ' +
    String(effective.slides) +
    ' slides',
);
console.log('sha256        ' + hash);
console.log('built in      ' + (elapsed / 1000).toFixed(1) + ' s');
console.log('');
console.log('features the generator put in:');
for (const [key, count] of Object.entries(summary.expected)) {
  if (count > 0) console.log('  ' + key.padEnd(20) + String(count));
}

/**
 * Experiment C, step 1 - write the swatch deck and the manifest that describes
 * it.
 *
 * ```
 * node tools/ground-truth/paint/colour/transforms/build-deck.ts <out-dir>
 * ```
 *
 * Produces `swatch-deck.pptx` and `swatch-inputs.json`. Each swatch is a
 * full-bleed rectangle in a fixed grid whose shape name is the swatch id, so
 * the readback in step 2 joins on a name rather than on a position it had to
 * guess.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, rect, SLIDE_HEIGHT, SLIDE_WIDTH, CLR_SCHEME } from '../../../lib/pptx.ts';
import { colorXml, describe, swatches } from './probes.ts';

const COLS = 10;
const ROWS = 8;
const PER_SLIDE = COLS * ROWS;

const CELL_W = Math.floor(SLIDE_WIDTH / COLS);
const CELL_H = Math.floor(SLIDE_HEIGHT / ROWS);

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/paint/colour/transforms/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const all = swatches();
const slides: string[] = [];

for (let start = 0; start < all.length; start += PER_SLIDE) {
  const page = all.slice(start, start + PER_SLIDE);
  const shapes = page.map((swatch, i) => {
    const col = i % COLS;
    const row = Math.floor(i / COLS);
    // The last column and row are stretched to the slide edge so no background
    // shows through at the margin; every swatch is still sampled at its centre.
    const x = col * CELL_W;
    const y = row * CELL_H;
    const cx = col === COLS - 1 ? SLIDE_WIDTH - x : CELL_W;
    const cy = row === ROWS - 1 ? SLIDE_HEIGHT - y : CELL_H;
    return rect(i + 2, swatch.id, x, y, cx, cy, colorXml(swatch));
  });
  slides.push(shapes.join(''));
}

const pptx = buildPptx({ slides });
const deckPath = join(outDir, 'swatch-deck.pptx');
writeFileSync(deckPath, pptx);

const manifest = {
  clrScheme: CLR_SCHEME,
  grid: { cols: COLS, rows: ROWS, perSlide: PER_SLIDE, cellW: CELL_W, cellH: CELL_H },
  slideSize: { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT },
  swatches: all.map((s, i) => ({
    id: s.id,
    slide: Math.floor(i / PER_SLIDE) + 1,
    index: i % PER_SLIDE,
    group: s.group,
    description: describe(s),
    base: s.base,
    transforms: s.transforms,
    xml: colorXml(s),
  })),
};
const manifestPath = join(outDir, 'swatch-inputs.json');
writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));

console.log(`${String(all.length)} swatches across ${String(slides.length)} slides`);
console.log(`  ${deckPath}  (${String(pptx.length)} bytes)`);
console.log(`  ${manifestPath}`);
const groups = new Map<string, number>();
for (const s of all) groups.set(s.group, (groups.get(s.group) ?? 0) + 1);
for (const [group, n] of [...groups].sort()) console.log(`  ${group.padEnd(20)} ${String(n)}`);

/**
 * Experiment C2, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/build-swatch-deck2.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `swatches2.ts`, plus `swatch2-inputs.json`. The
 * split into several decks is the whole point of the design: three of the blocks
 * ask questions PowerPoint may answer by refusing the package, and a refusal
 * that takes the other five blocks with it measures nothing.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, rect, CLR_SCHEME, SLIDE_HEIGHT, SLIDE_WIDTH } from './pptx.ts';
import { colorXml2, CROSSED_CLR_MAP, describe2, swatches2, type Swatch2 } from './swatches2.ts';

const COLS = 10;
const ROWS = 8;
const PER_SLIDE = COLS * ROWS;

const CELL_W = Math.floor(SLIDE_WIDTH / COLS);
const CELL_H = Math.floor(SLIDE_HEIGHT / ROWS);

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-swatch-deck2.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const all = swatches2();

const byDeck = new Map<string, Swatch2[]>();
for (const swatch of all) {
  const bucket = byDeck.get(swatch.deck);
  if (bucket) bucket.push(swatch);
  else byDeck.set(swatch.deck, [swatch]);
}

interface ManifestEntry {
  id: string;
  deck: string;
  slide: number;
  index: number;
  group: string;
  description: string;
  base: Swatch2['base'];
  transforms: Swatch2['transforms'];
  xml: string;
}

const manifest: ManifestEntry[] = [];
const decks: { deck: string; file: string; slides: number; swatches: number }[] = [];

for (const [deck, page] of byDeck) {
  const slides: string[] = [];
  for (let start = 0; start < page.length; start += PER_SLIDE) {
    const chunk = page.slice(start, start + PER_SLIDE);
    const shapes = chunk.map((swatch, i) => {
      const col = i % COLS;
      const row = Math.floor(i / COLS);
      const x = col * CELL_W;
      const y = row * CELL_H;
      const cx = col === COLS - 1 ? SLIDE_WIDTH - x : CELL_W;
      const cy = row === ROWS - 1 ? SLIDE_HEIGHT - y : CELL_H;
      manifest.push({
        id: swatch.id,
        deck,
        slide: slides.length + 1,
        index: i,
        group: swatch.group,
        description: describe2(swatch),
        base: swatch.base,
        transforms: swatch.transforms,
        xml: colorXml2(swatch),
      });
      return rect(i + 2, swatch.id, x, y, cx, cy, colorXml2(swatch));
    });
    slides.push(shapes.join(''));
  }

  const bytes = buildPptx(deck === 'clrmap' ? { slides, clrMap: CROSSED_CLR_MAP } : { slides });
  const file = `swatch2-${deck}.pptx`;
  writeFileSync(join(outDir, file), bytes);
  decks.push({ deck, file, slides: slides.length, swatches: page.length });
}

writeFileSync(
  join(outDir, 'swatch2-inputs.json'),
  JSON.stringify(
    {
      clrScheme: CLR_SCHEME,
      crossedClrMap: CROSSED_CLR_MAP,
      grid: { cols: COLS, rows: ROWS, perSlide: PER_SLIDE, cellW: CELL_W, cellH: CELL_H },
      slideSize: { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT },
      decks,
      swatches: manifest,
    },
    null,
    2,
  ),
);

console.log(`${String(all.length)} swatches across ${String(decks.length)} decks`);
for (const d of decks) {
  console.log(
    `  ${d.file.padEnd(24)} ${String(d.swatches).padStart(4)} swatches, ${String(d.slides)} slide(s)`,
  );
}
const groups = new Map<string, number>();
for (const s of all) groups.set(s.group, (groups.get(s.group) ?? 0) + 1);
for (const [group, n] of [...groups].sort()) console.log(`  ${group.padEnd(20)} ${String(n)}`);

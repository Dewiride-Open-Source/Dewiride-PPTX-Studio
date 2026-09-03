/**
 * Experiment C3, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/build-fill-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `fills.ts`, plus `fill-inputs.json` carrying the
 * exact rectangle of every probe so the analysis never has to guess where a
 * shape landed.
 *
 * ## Everything is aligned to a whole point
 *
 * The slide is 12192000 x 6858000 EMU, which is 960 x 540 pt, and one point is
 * 12700 EMU. Every rectangle this file emits is a whole number of points, so at
 * an export width of 960, 1920 or 3840 pixels each shape edge falls exactly on a
 * pixel boundary. Without that, a shape edge lands mid-pixel, the edge pixel is
 * a blend of the shape and the slide behind it, and for a pattern probe that is
 * a bit of the tile we cannot read.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, shapeXml, CLR_SCHEME, SLIDE_HEIGHT, SLIDE_WIDTH } from './pptx.ts';
import { DECK_GRID, fillProbes, hostileProbes, type Probe } from './fills.ts';

/** One point. Every emitted coordinate is a multiple of this. */
const PT = 12700;

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-fill-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const probes = [...fillProbes(), ...hostileProbes()];

const byDeck = new Map<string, Probe[]>();
for (const probe of probes) {
  const bucket = byDeck.get(probe.deck);
  if (bucket) bucket.push(probe);
  else byDeck.set(probe.deck, [probe]);
}

/** Round down to a whole point. Offsets may legitimately be zero. */
function pt(emu: number): number {
  return Math.floor(emu / PT) * PT;
}

/** The same, for an extent, which may not be zero. */
function ptSize(emu: number): number {
  return Math.max(PT, pt(emu));
}

interface Rect {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

/**
 * Where a probe's shape goes inside its cell.
 *
 * With no `aspect` the shape *is* the cell - which is what the ramp strips want,
 * because a strip spanning the full slide width means pixel column i is exactly
 * position (i + 0.5) / 1920 along the gradient, with no arithmetic in between.
 */
function place(probe: Probe, cell: Rect): Rect {
  if (probe.aspect === undefined) return cell;
  const scale = probe.scale ?? 0.88;
  const availW = cell.cx * scale;
  const availH = cell.cy * scale;
  const cx = ptSize(Math.min(availW, availH * probe.aspect));
  const cy = ptSize(cx / probe.aspect);
  return {
    x: pt(cell.x + (cell.cx - cx) / 2),
    y: pt(cell.y + (cell.cy - cy) / 2),
    cx,
    cy,
  };
}

interface ManifestEntry extends Rect {
  id: string;
  deck: string;
  slide: number;
  index: number;
  group: string;
  question: string;
  sample: Probe['sample'];
  fill: string;
  prst: string;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  hasBackdrop: boolean;
}

const manifest: ManifestEntry[] = [];
const decks: { deck: string; file: string; slides: number; probes: number }[] = [];

for (const [deck, page] of byDeck) {
  const grid = DECK_GRID[deck] ?? { cols: 1, rows: 1 };
  const perSlide = grid.cols * grid.rows;
  const cellW = SLIDE_WIDTH / grid.cols;
  const cellH = SLIDE_HEIGHT / grid.rows;

  const slides: string[] = [];
  for (let start = 0; start < page.length; start += perSlide) {
    const chunk = page.slice(start, start + perSlide);
    const shapes: string[] = [];
    // `p:cNvPr/@id` is unique within a part and 1 is taken by the group shape.
    let id = 2;
    chunk.forEach((probe, i) => {
      const col = i % grid.cols;
      const row = Math.floor(i / grid.cols);
      const cell: Rect = {
        x: pt(col * cellW),
        y: pt(row * cellH),
        cx: ptSize(cellW),
        cy: ptSize(cellH),
      };
      const rect = place(probe, cell);
      if (probe.backdrop !== undefined) {
        // Drawn first, so it is behind. An alpha probe measures nothing without
        // a known opaque colour underneath it, and the slide background is not
        // one - a neighbouring probe may be there instead.
        shapes.push(shapeXml({ id: id++, name: `${probe.id}-bg`, ...rect, fill: probe.backdrop }));
      }
      shapes.push(
        shapeXml({
          id: id++,
          name: probe.id,
          ...rect,
          fill: probe.fill,
          prst: probe.prst,
          rot: probe.rot,
          flipH: probe.flipH,
          flipV: probe.flipV,
        }),
      );
      manifest.push({
        id: probe.id,
        deck,
        slide: slides.length + 1,
        index: i,
        group: probe.group,
        question: probe.question,
        sample: probe.sample,
        fill: probe.fill,
        prst: probe.prst ?? 'rect',
        rot: probe.rot ?? 0,
        flipH: probe.flipH ?? false,
        flipV: probe.flipV ?? false,
        hasBackdrop: probe.backdrop !== undefined,
        ...rect,
      });
    });
    slides.push(shapes.join(''));
  }

  const file = `fill-${deck}.pptx`;
  writeFileSync(join(outDir, file), buildPptx({ slides }));
  decks.push({ deck, file, slides: slides.length, probes: page.length });
}

writeFileSync(
  join(outDir, 'fill-inputs.json'),
  JSON.stringify(
    {
      clrScheme: CLR_SCHEME,
      slideSize: { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT },
      emuPerPoint: PT,
      decks,
      probes: manifest,
    },
    null,
    2,
  ),
);

console.log(`${String(probes.length)} probes across ${String(decks.length)} decks`);
const groups = new Map<string, number>();
for (const p of probes) groups.set(p.group, (groups.get(p.group) ?? 0) + 1);
for (const [group, n] of [...groups].sort((a, b) => a[0].localeCompare(b[0])))
  console.log(`  ${group.padEnd(10)} ${String(n)}`);

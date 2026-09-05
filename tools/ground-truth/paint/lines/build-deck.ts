/**
 * Experiment C4, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/build-line-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `lines.ts`, plus `line-inputs.json` carrying
 * every probe's rectangle and its read region, so the analysis never has to
 * guess where anything landed.
 *
 * ## Placement is absolute, not a grid
 *
 * C3 laid its probes out on a grid and let the builder decide where each shape
 * went. That does not work here. A stroke is drawn *around* its geometry, so
 * half of what is being measured is outside the shape's own rectangle - a square
 * cap on a 24pt line paints 12pt past the endpoint, a 24pt shadow blur reaches
 * 24pt past the shape, and a region expressed relative to the shape could not
 * name either. So every probe in `lines.ts` states its rectangle and its read
 * window in points on the slide, and this file only converts points to EMU.
 *
 * Points, because the slide is exactly 960 x 540 of them and 12700 EMU is one,
 * so at 1920 or 3840 pixels wide every whole-point coordinate is a whole-pixel
 * boundary. Sub-point coordinates are the one thing this file refuses.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, shapeXml, CLR_SCHEME, SLIDE_HEIGHT, SLIDE_WIDTH } from './pptx.ts';
import { EMU_PER_POINT, hostileProbes, lineProbes, type Probe } from './lines.ts';

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-line-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const probes = [...lineProbes(), ...hostileProbes()];

/** Points to EMU, refusing anything that is not a whole point. */
function emu(points: number, what: string): number {
  if (!Number.isInteger(points)) {
    throw new Error(`${what} is ${String(points)}pt, which is not a whole point`);
  }
  return points * EMU_PER_POINT;
}

const byDeck = new Map<string, Probe[]>();
for (const probe of probes) {
  const bucket = byDeck.get(probe.deck);
  if (bucket) bucket.push(probe);
  else byDeck.set(probe.deck, [probe]);
}

interface ManifestEntry {
  id: string;
  deck: string;
  group: string;
  question: string;
  /** Points. */
  x: number;
  y: number;
  cx: number;
  cy: number;
  read: Probe['read'];
  line: string;
  effect: string;
  geom: string;
}

const manifest: ManifestEntry[] = [];
const decks: { deck: string; file: string; probes: number }[] = [];

for (const [deck, page] of byDeck) {
  const shapes: string[] = [];
  // `p:cNvPr/@id` is unique within a part, and 1 is taken by the group shape.
  let id = 2;
  for (const probe of page) {
    shapes.push(
      shapeXml({
        id: id++,
        name: probe.id,
        x: emu(probe.x, `${probe.id}.x`),
        y: emu(probe.y, `${probe.id}.y`),
        cx: emu(probe.cx, `${probe.id}.cx`),
        cy: emu(probe.cy, `${probe.id}.cy`),
        geom: probe.geom,
        fill: probe.fill ?? '<a:noFill/>',
        line: probe.line,
        effect: probe.effect,
        rot: probe.rot,
      }),
    );
    manifest.push({
      id: probe.id,
      deck,
      group: probe.group,
      question: probe.question,
      x: probe.x,
      y: probe.y,
      cx: probe.cx,
      cy: probe.cy,
      read: probe.read,
      line: probe.line ?? '',
      effect: probe.effect ?? '',
      geom: probe.geom ?? '',
    });
  }

  const file = `line-${deck}.pptx`;
  writeFileSync(join(outDir, file), buildPptx({ slides: [shapes.join('')] }));
  decks.push({ deck, file, probes: page.length });
}

/**
 * Which decks are exported at 3840 rather than 1920.
 *
 * A dash run of four line-widths at 12pt is 192 px at 3840 and 96 at 1920, and
 * the run boundary is found to about a pixel either way - so the finer export
 * halves the error on the number this whole sub-phase turns on. The arrowhead
 * decks get it for the same reason: the outline is recovered one column at a
 * time and every column is a sample of the marker's half-height.
 */
const FINE_DECKS = [
  'dash',
  'dashw',
  'dashcap',
  'arrow0',
  'arrow1',
  'arrow2',
  'arrow3',
  'arrow4',
  'arrow5',
  'ahscale',
];

writeFileSync(
  join(outDir, 'line-inputs.json'),
  JSON.stringify(
    {
      clrScheme: CLR_SCHEME,
      slideSize: { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT },
      slidePoints: { w: SLIDE_WIDTH / EMU_PER_POINT, h: SLIDE_HEIGHT / EMU_PER_POINT },
      emuPerPoint: EMU_PER_POINT,
      fineDecks: FINE_DECKS,
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
  console.log(`  ${group.padEnd(12)} ${String(n)}`);

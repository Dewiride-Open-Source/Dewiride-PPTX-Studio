/**
 * Experiment C6, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/paint/blips/build-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `probes.ts`, the three probe images beside them,
 * and `blip-inputs.json` carrying every probe's rectangle so the analysis never
 * has to guess where a shape landed.
 *
 * Every rectangle is a whole number of points, so at an export width of 1920 or
 * 3840 each shape edge falls on a pixel boundary and no edge pixel is a blend of
 * the shape with the slide behind it.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { writeBmp } from '../../lib/bmp.ts';
import { buildPptx, shapeXml, SLIDE_HEIGHT, SLIDE_WIDTH, type MediaPart } from '../../lib/pptx.ts';
import {
  assertOnSlide,
  blipProbes,
  quadPixel,
  QUAD_PX,
  RAMP_H,
  RAMP_W,
  rampPixel,
  type BlipProbe,
  type DeckId,
} from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined) {
  throw new Error('usage: node tools/ground-truth/paint/blips/build-deck.ts <out-dir>');
}
mkdirSync(outDir, { recursive: true });

/** Media order is the contract: `buildPptx` hands out rId2 upwards in this order. */
const media: readonly MediaPart[] = [
  { name: 'quad.bmp', bytes: writeBmp(QUAD_PX, QUAD_PX, quadPixel, 96) },
  { name: 'quad150.bmp', bytes: writeBmp(QUAD_PX, QUAD_PX, quadPixel, 150) },
  { name: 'ramp.bmp', bytes: writeBmp(RAMP_W, RAMP_H, rampPixel, 96) },
];

const probes = blipProbes();
assertOnSlide(probes);

const byDeck = new Map<DeckId, BlipProbe[]>();
for (const probe of probes) {
  const bucket = byDeck.get(probe.deck);
  if (bucket === undefined) byDeck.set(probe.deck, [probe]);
  else bucket.push(probe);
}

/** Both resolutions, so a tile period fixed in pixels is separable from one in points. */
const WIDTHS = [1920, 3840];

const decks = [...byDeck].map(([id, list]) => {
  const shapes = list
    .map((probe, index) =>
      shapeXml({
        id: index + 10,
        name: probe.id,
        x: probe.rect.x,
        y: probe.rect.y,
        cx: probe.rect.cx,
        cy: probe.rect.cy,
        fill: probe.fill,
      }),
    )
    .join('');
  const file = `blip-${id}.pptx`;
  writeFileSync(join(outDir, file), buildPptx({ slides: [shapes], media }));
  return {
    id,
    file,
    widths: WIDTHS,
    probes: list.map((probe) => ({
      id: probe.id,
      question: probe.question,
      image: probe.image,
      rect: probe.rect,
      fill: probe.fill,
    })),
  };
});

for (const part of media) writeFileSync(join(outDir, part.name), part.bytes);

const inputs = {
  experiment: 'C6 - image fills',
  slide: { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT },
  images: {
    quad: { file: 'quad.bmp', width: QUAD_PX, height: QUAD_PX, dpi: 96 },
    quad150: { file: 'quad150.bmp', width: QUAD_PX, height: QUAD_PX, dpi: 150 },
    ramp: { file: 'ramp.bmp', width: RAMP_W, height: RAMP_H, dpi: 96 },
  },
  decks,
};
writeFileSync(join(outDir, 'blip-inputs.json'), `${JSON.stringify(inputs, null, 2)}\n`);

console.log(
  `wrote ${String(decks.length)} deck(s), ${String(probes.length)} probe(s) to ${outDir}`,
);

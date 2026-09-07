/**
 * Experiment C7, step 1 - write the probe deck.
 *
 * ```
 * node tools/ground-truth/render/pictures/build-deck.ts <out-dir>
 * ```
 *
 * One deck, the probe image beside it, and `picture-inputs.json` carrying every
 * probe's rectangle so the analysis never has to guess where a shape landed.
 *
 * Every rectangle is a whole number of points, so at an export width of 1920
 * each shape edge falls on a pixel boundary and no sampled pixel is a blend of
 * the shape with the slide behind it.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { writeBmp } from '../../lib/bmp.ts';
import { buildPptx, picXml, shapeXml, type MediaPart } from '../../lib/pptx.ts';
import {
  assertOnSlide,
  OUTLINE_PT,
  pictureProbes,
  PT,
  quadPixel,
  QUAD_PX,
  shapeControl,
} from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined) {
  throw new Error('usage: node tools/ground-truth/render/pictures/build-deck.ts <out-dir>');
}
mkdirSync(outDir, { recursive: true });

/** Media order is the contract: `buildPptx` hands out rId2 upwards in this order. */
const media: readonly MediaPart[] = [
  { name: 'quad.bmp', bytes: writeBmp(QUAD_PX, QUAD_PX, quadPixel, 96) },
];

const probes = pictureProbes();
assertOnSlide(probes);

const shapes = probes
  .map((probe, index) =>
    picXml({
      id: index + 10,
      name: probe.id,
      x: probe.rect.x,
      y: probe.rect.y,
      cx: probe.rect.cx,
      cy: probe.rect.cy,
      embed: 'rId2',
      prst: probe.prst,
      rot: probe.rot,
      flipH: probe.flipH,
      flipV: probe.flipV,
      spPrFill: probe.spPrFill,
      line: probe.line,
    }),
  )
  .join('');

// The same outline on a `p:sp`, so the picture rule is read against a shape
// from one export rather than against 2.8 from memory.
const LINE = `<a:ln w="${String(OUTLINE_PT * PT)}"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>`;
const control = shapeControl();
const controlXml = shapeXml({
  id: 100,
  name: control.id,
  x: control.rect.x,
  y: control.rect.y,
  cx: control.rect.cx,
  cy: control.rect.cy,
  fill: '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
  line: LINE,
});

/** Black, so a pixel the picture did not paint reads as no colour at all. */
const BACKGROUND = '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>';

const file = 'pictures.pptx';
writeFileSync(
  join(outDir, file),
  buildPptx({ slides: [shapes + controlXml], media, backgrounds: [BACKGROUND] }),
);

const inputs = {
  decks: [
    {
      id: 'pictures',
      file,
      widths: [1920],
      control: { id: control.id, rect: control.rect },
      probes: probes.map((probe) => ({
        id: probe.id,
        question: probe.question,
        rect: probe.rect,
        prst: probe.prst,
        flipH: probe.flipH,
        flipV: probe.flipV,
        spPrFill: probe.spPrFill !== undefined,
        line: probe.line !== undefined,
      })),
    },
  ],
};

writeFileSync(join(outDir, 'picture-inputs.json'), `${JSON.stringify(inputs, null, 2)}\n`);
console.log(`wrote ${String(probes.length)} probe(s) to ${join(outDir, file)}`);

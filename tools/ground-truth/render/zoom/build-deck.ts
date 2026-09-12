/**
 * Experiment F2, step 1 - write the probe deck.
 *
 * `node tools/ground-truth/render/zoom/build-deck.ts <out-dir>` writes `zoom-probes.pptx` and
 * `zoom-inputs.json`, which carries every probe's read region in points and the export widths.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { writeBmp } from '../../lib/bmp.ts';
import { buildPptx, type MediaPart } from '../../lib/pptx.ts';
import { assertSeparable, backgroundOf, EXPORT_WIDTHS, SLIDE, zoomProbes } from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined) {
  throw new Error('usage: node tools/ground-truth/render/zoom/build-deck.ts <out-dir>');
}
mkdirSync(outDir, { recursive: true });

const probes = zoomProbes();
assertSeparable(probes);

/** A white image: the picture probes measure the border, and nothing else may ink. */
const media: readonly MediaPart[] = [
  { name: 'white.bmp', bytes: writeBmp(32, 32, () => 0xffffff, 96) },
];

const slideCount = Math.max(...probes.map((p) => p.slide));
const slides: string[] = [];
const backgrounds: (string | undefined)[] = [];
for (let n = 1; n <= slideCount; n++) {
  slides.push(
    probes
      .filter((p) => p.slide === n)
      .map((p) => p.markup)
      .join(''),
  );
  backgrounds.push(backgroundOf(n));
}

writeFileSync(join(outDir, 'zoom-probes.pptx'), buildPptx({ slides, media, backgrounds }));

writeFileSync(
  join(outDir, 'zoom-inputs.json'),
  `${JSON.stringify(
    {
      deck: 'zoom-probes.pptx',
      slidePoints: SLIDE,
      widths: EXPORT_WIDTHS,
      slides: slideCount,
      probes: probes.map(({ markup: _markup, ...rest }) => rest),
    },
    null,
    2,
  )}\n`,
);

console.log(
  `wrote zoom-probes.pptx: ${String(probes.length)} probes on ${String(slideCount)} slides, ` +
    `to export at ${EXPORT_WIDTHS.join(', ')} px wide`,
);

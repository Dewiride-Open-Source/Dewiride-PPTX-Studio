/**
 * Experiment F3, step 1 - write the probe deck.
 *
 * `node tools/ground-truth/render/snap/build-deck.ts <out-dir>` writes `snap-probes.pptx` and
 * `snap-inputs.json`, which carries every probe's read window in points and the export widths.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { writeBmp } from '../../lib/bmp.ts';
import { buildPptx, type MediaPart } from '../../lib/pptx.ts';
import { assertSeparable, EXPORT_WIDTHS, SLIDE, snapProbes } from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined) {
  throw new Error('usage: node tools/ground-truth/render/snap/build-deck.ts <out-dir>');
}
mkdirSync(outDir, { recursive: true });

const probes = snapProbes();
assertSeparable(probes);

/** A black image for the picture-edge probes, and a white one so a border probe inks nothing else. */
const media: readonly MediaPart[] = [
  { name: 'black.bmp', bytes: writeBmp(32, 32, () => 0x000000, 96) },
  { name: 'white.bmp', bytes: writeBmp(32, 32, () => 0xffffff, 96) },
];

const slideCount = Math.max(...probes.map((p) => p.slide));
const slides: string[] = [];
for (let n = 1; n <= slideCount; n++) {
  slides.push(
    probes
      .filter((p) => p.slide === n)
      .map((p) => p.markup)
      .join(''),
  );
}

writeFileSync(join(outDir, 'snap-probes.pptx'), buildPptx({ slides, media }));

writeFileSync(
  join(outDir, 'snap-inputs.json'),
  `${JSON.stringify(
    {
      deck: 'snap-probes.pptx',
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
  `wrote snap-probes.pptx: ${String(probes.length)} probes on ${String(slideCount)} slides, ` +
    `to export at ${EXPORT_WIDTHS.join(', ')} px wide`,
);

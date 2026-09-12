/**
 * Experiment T13's probe fonts, written to a directory for another experiment.
 *
 * ```
 * node tools/ground-truth/fonts/metrics/build-fonts.ts <dir>
 * ```
 *
 * T14 measures the faces a runner has, and a runner's own faces agree with
 * themselves; `--font-dir` pointed here adds fourteen that do not, so a run on
 * any rasteriser separates the three tables by 100 px on `split`. ADR 0052.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { probeFonts } from './probes.ts';

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: build-fonts.ts <dir>');
const out = resolve(dir);
mkdirSync(out, { recursive: true });

const fonts = probeFonts();
for (const font of fonts) {
  writeFileSync(join(out, `${font.id}.ttf`), Buffer.from(font.base64, 'base64'));
}
console.log(`wrote ${String(fonts.length)} probe font(s) to ${out}`);

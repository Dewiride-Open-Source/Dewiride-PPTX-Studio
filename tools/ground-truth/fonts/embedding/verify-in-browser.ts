/**
 * Experiment B, step 2 - is the probe font valid, independently of PowerPoint?
 *
 * ```
 * node tools/ground-truth/fonts/embedding/verify-in-browser.ts <dir-with-probe.ttf>
 * ```
 *
 * Experiment B has two variables in it: a font we built and an EOT wrapper we
 * built. If PowerPoint shows nothing, one of them is wrong and we would not know
 * which. So the font is judged first by a party with no stake in the outcome.
 *
 * Chromium is a good judge. Every font it loads goes through OTS, the
 * OpenType Sanitiser, which is stricter than the specification and rejects
 * outright a great deal that a lenient reader would accept. A font OTS takes is
 * a font whose table directory, checksums, cmap, loca and glyf are all
 * defensible.
 *
 * Then it draws the sample and measures the bars, so a "loaded" that renders
 * nothing does not pass either.
 *
 * Uses the Chromium that `pnpm test` already runs in. Installs nothing.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { chromium } from 'playwright';

/*
 * The body of `page.evaluate` runs in Chromium, not in Node, so the browser
 * globals it uses are not in scope for the type checker - `tools/` is compiled
 * with `types: ["node"]` and no DOM library, deliberately, because every other
 * tool here would be wrong to touch the DOM.
 *
 * Rather than widen the whole tools project's lib, declare exactly the surface
 * this one file uses. Narrow enough to stay honest: if the evaluate body starts
 * using something else, it fails to compile rather than silently becoming `any`.
 */
interface Ctx2D {
  fillStyle: string;
  font: string;
  textBaseline: string;
  fillRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number): void;
  measureText(text: string): { width: number };
  getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray };
}
interface Canvas {
  width: number;
  height: number;
  getContext(id: '2d'): Ctx2D | null;
}
declare const document: {
  getElementById(id: string): Canvas | null;
  fonts: { add(face: BrowserFontFace): void };
};
declare const FontFace: new (family: string, source: ArrayBuffer) => BrowserFontFace;
interface BrowserFontFace {
  load(): Promise<unknown>;
}

const dir = process.argv[2];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/fonts/embedding/verify-in-browser.ts <dir>');

const bytes = [...new Uint8Array(readFileSync(join(dir, 'probe.ttf')))];
const SAMPLE = 'ABCDEFGH';

const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 1000, height: 300 } });
  await page.setContent('<canvas id="c" width="1000" height="300"></canvas>');

  const result = await page.evaluate(
    async ({ data, sample }: { data: number[]; sample: string }) => {
      const buffer = new Uint8Array(data).buffer;
      const face = new FontFace('ProbeUnderTest', buffer);
      try {
        await face.load();
      } catch (error) {
        return {
          loaded: false,
          error: String(error),
          bars: [] as number[],
          widths: [] as number[],
        };
      }
      document.fonts.add(face);

      const canvas = document.getElementById('c')!;
      const ctx = canvas.getContext('2d')!;
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#000';
      ctx.font = '100px ProbeUnderTest';
      ctx.textBaseline = 'alphabetic';
      ctx.fillText(sample, 10, 200);

      const widths = [...sample].map((ch) => ctx.measureText(ch).width);

      // For each glyph cell, the height of the tallest run of dark pixels.
      const image = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      const columnHeight = (x: number): number => {
        let top = -1;
        let bottom = -1;
        for (let y = 0; y < canvas.height; y++) {
          const dark = image[(y * canvas.width + x) * 4]! < 128;
          if (dark) {
            if (top < 0) top = y;
            bottom = y;
          }
        }
        return top < 0 ? 0 : bottom - top + 1;
      };
      const advance = widths[0] ?? 60;
      const bars = [...sample].map((_, i) => columnHeight(Math.round(10 + advance * (i + 0.5))));
      return { loaded: true, error: '', bars, widths };
    },
    { data: bytes, sample: SAMPLE },
  );

  console.log(`FontFace.load(): ${result.loaded ? 'accepted' : 'REJECTED'}`);
  if (!result.loaded) {
    console.log(`  ${result.error}`);
    console.log('\nChromium refused the font. The EOT wrapper is not the problem - the font is.');
    process.exitCode = 1;
  } else {
    console.log(`  advance widths : ${result.widths.map((w) => w.toFixed(1)).join(' ')}`);
    console.log(`  bar heights px : ${result.bars.join(' ')}`);
    // The glyphs are 100..800 font units tall at 1000 upem, drawn at 100px, so
    // the bars should be 10..80px and strictly increasing.
    const increasing = result.bars.every((h, i) => i === 0 || h > result.bars[i - 1]!);
    const inRange = result.bars.every((h) => h >= 5 && h <= 95);
    console.log(`  strictly increasing staircase : ${increasing ? 'yes' : 'NO'}`);
    console.log(`  every bar within 5..95px      : ${inRange ? 'yes' : 'NO'}`);
    if (!increasing || !inRange) process.exitCode = 1;
  }

  const shot = await page.screenshot();
  writeFileSync(join(dir, 'chromium-reference.png'), shot);
  console.log(`\nwrote ${join(dir, 'chromium-reference.png')}`);
} finally {
  await browser.close();
}

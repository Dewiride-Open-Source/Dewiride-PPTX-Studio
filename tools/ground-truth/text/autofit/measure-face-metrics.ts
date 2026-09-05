/**
 * Experiment T4, step 5 - can a browser predict the two per-face constants?
 *
 * ```
 * node tools/ground-truth/text/autofit/measure-face-metrics.ts <work-dir>
 * ```
 *
 * T4 measured two numbers that are properties of a typeface rather than of the
 * format:
 *
 * - **`b`**, the per-size term in `lastLine = 0.75 x advance + b x size`, which
 *   spans 0.2053 on Tahoma to 0.3000 on Courier New.
 * - **`C`**, the multiple of the size below which the last line takes the whole
 *   advance instead. Indistinguishable from 1.2 on five faces and about 1.27 on
 *   Courier New.
 *
 * Six hardcoded faces is not a renderer. Either these are derivable from
 * something a browser can report for *any* font, or every deck using a seventh
 * typeface is laid out on a guess - so this asks the browser what it knows.
 *
 * `TextMetrics.fontBoundingBoxAscent` and `fontBoundingBoxDescent` are the
 * candidates: Chromium derives them from the face's own vertical metrics, they
 * are available in a Worker through `OffscreenCanvas`, and they need no font
 * file to be parsed. `emHeightAscent`/`emHeightDescent` and the alphabetic and
 * ideographic baselines are recorded beside them, because if the first pair
 * does not explain the constants one of the others might, and re-running the
 * browser is cheaper than re-running PowerPoint.
 *
 * This measures; `tools/ground-truth/text/autofit/analyse.ts` decides. Nothing here fits anything.
 */

import { writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

/*
 * The `page.evaluate` body runs in Chromium, not in Node. `tools/` compiles
 * with `types: ["node"]` and no DOM library, deliberately, so the exact surface
 * used is declared here - narrow enough that a body reaching for something else
 * fails to compile rather than becoming `any`.
 */
interface Metrics {
  readonly width: number;
  readonly fontBoundingBoxAscent: number;
  readonly fontBoundingBoxDescent: number;
  readonly emHeightAscent: number;
  readonly emHeightDescent: number;
  readonly actualBoundingBoxAscent: number;
  readonly actualBoundingBoxDescent: number;
  readonly alphabeticBaseline: number;
  readonly ideographicBaseline: number;
  readonly hangingBaseline: number;
}
interface OffscreenCtx {
  font: string;
  fontKerning: string;
  measureText(text: string): Metrics;
}
interface OffscreenCanvasCtor {
  new (width: number, height: number): { getContext(id: '2d'): OffscreenCtx | null };
}
interface FontFaceSetLike {
  check(font: string): boolean;
}

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: measure-face-metrics.ts <out-dir>');
const work = resolve(outDir);

/** The six faces T4 measured against PowerPoint, and nothing else - a face the
 *  experiment did not measure would have nothing to be compared against. */
const FACES = ['Arial', 'Times New Roman', 'Verdana', 'Courier New', 'Georgia', 'Tahoma'];
/** Sizes spanning the range the `last` deck used, to see whether the browser's
 *  numbers are exactly proportional to the size or only nearly so. */
const SIZES = [8, 11, 14, 18, 24, 32, 44, 1000, 4096];

interface FaceMetric {
  readonly face: string;
  readonly sizePt: number;
  readonly installed: boolean;
  readonly metrics: Metrics;
}

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const results = await page.evaluate(
    ({ faces, sizes }: { faces: readonly string[]; sizes: readonly number[] }): FaceMetric[] => {
      const Ctor = (globalThis as unknown as { OffscreenCanvas: OffscreenCanvasCtor })
        .OffscreenCanvas;
      // `document.fonts` on a page; `self.fonts` in a Worker. This runs on a
      // page, so it is the first - the shipped measurer runs in the second and
      // that difference is why the surface is declared rather than assumed.
      const fonts = (globalThis as unknown as { document: { fonts: FontFaceSetLike } }).document
        .fonts;
      const ctx = new Ctor(64, 64).getContext('2d');
      if (ctx === null) throw new Error('no 2d context');
      ctx.fontKerning = 'none';
      const out: FaceMetric[] = [];
      for (const face of faces) {
        for (const sizePt of sizes) {
          const font = `${String(sizePt)}px "${face}"`;
          ctx.font = font;
          // A probe string with an ascender, a descender and a round letter, so
          // the actual-bounding-box numbers are not measuring one glyph's quirk.
          const m = ctx.measureText('Hxgpo');
          out.push({
            face,
            sizePt,
            installed: fonts.check(font),
            metrics: {
              width: m.width,
              fontBoundingBoxAscent: m.fontBoundingBoxAscent,
              fontBoundingBoxDescent: m.fontBoundingBoxDescent,
              emHeightAscent: m.emHeightAscent,
              emHeightDescent: m.emHeightDescent,
              actualBoundingBoxAscent: m.actualBoundingBoxAscent,
              actualBoundingBoxDescent: m.actualBoundingBoxDescent,
              alphabeticBaseline: m.alphabeticBaseline,
              ideographicBaseline: m.ideographicBaseline,
              hangingBaseline: m.hangingBaseline,
            },
          });
        }
      }
      return out;
    },
    { faces: FACES, sizes: SIZES },
  );

  const version = browser.version();
  writeFileSync(
    join(work, 'face-metrics.json'),
    `${JSON.stringify({ chromium: version, results }, null, 2)}\n`,
  );

  console.log(`chromium ${version}`);
  const missing = results.filter((r) => !r.installed).map((r) => r.face);
  if (missing.length > 0) {
    throw new Error(
      `not installed, so the numbers are a fallback face: ${[...new Set(missing)].join(', ')}`,
    );
  }
  // One line per face at 100 units, where every ratio reads directly.
  for (const face of FACES) {
    const at = results.find((r) => r.face === face && r.sizePt === 4096);
    if (at === undefined) throw new Error(`no reading for ${face}`);
    const m = at.metrics;
    console.log(
      `  ${face.padEnd(16)} fontBox ${((m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) / 4096).toFixed(6)}` +
        `  emHeight ${((m.emHeightAscent + m.emHeightDescent) / 4096).toFixed(6)}` +
        `  descent ${(m.fontBoundingBoxDescent / 4096).toFixed(6)}`,
    );
  }
} finally {
  await browser.close();
}

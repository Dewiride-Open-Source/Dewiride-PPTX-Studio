/**
 * Experiment T8, step 3 - the same faces, measured where the renderer runs.
 *
 * ```
 * node tools/ground-truth/render/text/measure-in-browser.ts <work-dir>
 * ```
 *
 * Two questions can only be answered in terms the browser supplies at run time.
 * The baseline sits some way down the line box, and what a renderer needs is
 * that distance as a function of something it can measure itself. The underline
 * sits some way below the baseline, and what a renderer needs to know is whether
 * `text-decoration` puts it in the same place - because if it does, the right
 * implementation is to ask for it rather than to draw it.
 */

import { writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

import { FACES, SIZES } from './probes.ts';

/*
 * The bodies of `page.evaluate` run in Chromium, not in Node, so the browser
 * globals they use are not in scope for the type checker. Declare exactly the
 * surface these two use rather than widening the project's lib.
 */
interface Metrics {
  width: number;
  fontBoundingBoxAscent: number;
  fontBoundingBoxDescent: number;
  actualBoundingBoxAscent: number;
  actualBoundingBoxDescent: number;
  alphabeticBaseline: number;
}
interface Ctx2D {
  font: string;
  letterSpacing: string;
  fontKerning: string;
  textBaseline: string;
  measureText(text: string): Metrics;
  drawImage(image: BrowserImage, x: number, y: number): void;
  getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray };
}
interface Canvas {
  width: number;
  height: number;
  getContext(id: '2d'): Ctx2D | null;
}
interface BrowserImage {
  src: string;
  width: number;
  height: number;
  decode(): Promise<void>;
}
declare const OffscreenCanvas: new (w: number, h: number) => Canvas;
declare const document: { createElement(tag: 'canvas'): Canvas };
declare const Image: new () => BrowserImage;

const dirArg = process.argv[2];
if (dirArg === undefined) throw new Error('usage: measure-in-browser.ts <dir>');
const dir: string = resolve(dirArg);

/** The size the decoration page is set at; one em is this many CSS pixels. */
const DECORATION_PX = 100;

/** Chromium is asked for four device pixels per CSS pixel, so a rule resolves. */
const DEVICE_SCALE = 4;

/** The marker's bottom edge sits on the baseline, which is what makes it useful. */
const MARKER_PX = 8;

function decorationPage(face: string): string {
  const family = face.replace(/["\\]/g, '');
  return (
    '<!doctype html><meta charset="utf-8">' +
    '<style>html,body{margin:0;background:#fff}' +
    `.row{position:absolute;left:40px;font:${String(DECORATION_PX)}px "${family}";` +
    'white-space:pre;color:#000;text-decoration-color:#000}' +
    `.mark{display:inline-block;width:${String(MARKER_PX)}px;height:${String(MARKER_PX)}px;background:#f00}` +
    '</style>' +
    '<div class="row" style="top:40px"><span class="mark"></span><span style="text-decoration:underline">Alpha bravo</span></div>' +
    '<div class="row" style="top:240px"><span class="mark"></span><span style="text-decoration:line-through">Alpha bravo</span></div>'
  );
}

const browser = await chromium.launch();
try {
  /* ---- the font box, for the baseline question --------------------------- */
  const page = await browser.newPage();
  const measured = await page.evaluate(
    ({ faces, sizes }: { faces: readonly string[]; sizes: readonly number[] }) => {
      const canvas = new OffscreenCanvas(64, 64);
      const ctx = canvas.getContext('2d');
      if (ctx === null) throw new Error('no 2d context');

      const read = (family: string, px: number, text: string): Record<string, number> => {
        ctx.font = `${String(px)}px "${family.replace(/["\\]/g, '')}"`;
        ctx.letterSpacing = '0px';
        ctx.fontKerning = 'normal';
        ctx.textBaseline = 'alphabetic';
        const m = ctx.measureText(text);
        return {
          width: m.width,
          fontAscent: m.fontBoundingBoxAscent,
          fontDescent: m.fontBoundingBoxDescent,
          inkAscent: m.actualBoundingBoxAscent,
          inkDescent: m.actualBoundingBoxDescent,
          alphabetic: m.alphabeticBaseline,
        };
      };

      const out: Record<string, unknown> = {};
      for (const family of faces) {
        const perSize: Record<string, unknown> = {};
        for (const px of sizes) perSize[String(px)] = read(family, px, 'xxxxxxxx');
        // Chromium rounds the font box to whole pixels, so a ratio read off a
        // small size is quantised; the ladder shows how fast that settles.
        const ladder: Record<string, unknown> = {};
        for (const px of [100, 400, 1000, 4000]) ladder[String(px)] = read(family, px, 'Hxg');
        out[family] = { ladder, perSize };
      }
      return out;
    },
    { faces: FACES, sizes: SIZES },
  );

  /* ---- the decorations, read off Chromium's own rendering ---------------- */
  const context = await browser.newContext({ deviceScaleFactor: DEVICE_SCALE });
  const shot = await context.newPage();
  await shot.setViewportSize({ width: 900, height: 420 });

  const decorations: Record<string, unknown> = {};
  for (const face of FACES) {
    await shot.setContent(decorationPage(face));
    const png = (await shot.screenshot({ type: 'png' })).toString('base64');
    // Chromium decodes its own screenshot, so nothing here needs a PNG library.
    decorations[face] = await shot.evaluate(
      async ({ data, scale, px }: { data: string; scale: number; px: number; marker: number }) => {
        const image = new Image();
        image.src = `data:image/png;base64,${data}`;
        await image.decode();
        const canvas = document.createElement('canvas');
        canvas.width = image.width;
        canvas.height = image.height;
        const ctx = canvas.getContext('2d');
        if (ctx === null) throw new Error('no 2d context');
        ctx.drawImage(image, 0, 0);
        const { data: pixels } = ctx.getImageData(0, 0, canvas.width, canvas.height);

        const at = (x: number, y: number): readonly [number, number, number] => {
          const i = (y * canvas.width + x) * 4;
          return [pixels[i] ?? 255, pixels[i + 1] ?? 255, pixels[i + 2] ?? 255];
        };

        interface Span {
          y: number;
          left: number;
          right: number;
          count: number;
        }

        const rowOf = (
          y: number,
          from: number,
          keep: (r: number, g: number, b: number) => boolean,
        ): Span | null => {
          let left = -1;
          let right = -1;
          let count = 0;
          for (let x = from; x < canvas.width; x++) {
            const [r, g, b] = at(x, y);
            if (!keep(r, g, b)) continue;
            if (left < 0) left = x;
            right = x;
            count++;
          }
          return left < 0 ? null : { y, left, right, count };
        };

        const merge = (
          spans: readonly Span[],
        ): { from: number; to: number; left: number; right: number }[] => {
          const out: { from: number; to: number; left: number; right: number }[] = [];
          for (const span of spans) {
            const last = out.at(-1);
            if (last !== undefined && span.y === last.to + 1) {
              last.to = span.y;
              last.left = Math.min(last.left, span.left);
              last.right = Math.max(last.right, span.right);
            } else out.push({ from: span.y, to: span.y, left: span.left, right: span.right });
          }
          return out;
        };

        const band = (top: number, height: number): Record<string, number> => {
          const from = Math.round(top * scale);
          const to = Math.min(canvas.height, Math.round((top + height) * scale));

          const red: Span[] = [];
          for (let y = from; y < to; y++) {
            const row = rowOf(y, 0, (r, g, b) => r > 150 && g < 120 && b < 120);
            if (row !== null) red.push(row);
          }
          const marker = merge(red)[0];
          if (marker === undefined) throw new Error('no baseline marker on the page');
          // A marker is an inline block, so its bottom edge sits on the baseline.
          const baseline = (marker.to + 1) / scale;

          const dark: Span[] = [];
          for (let y = from; y < to; y++) {
            const row = rowOf(y, marker.right + 2, (r, g, b) => r < 100 && g < 100 && b < 100);
            if (row !== null) dark.push(row);
          }
          if (dark.length === 0) throw new Error('the page drew no text at all');
          const widest = Math.max(...dark.map((span) => span.right - span.left));
          // A row of glyphs can span the whole width but is never nine tenths
          // ink, which is what separates a rule from the middle of a word.
          const solid = dark.filter(
            (span) =>
              span.right - span.left >= widest - 8 && span.count > (span.right - span.left) * 0.9,
          );
          const rules = merge(solid);
          if (rules.length !== 1) {
            throw new Error(`the page drew ${String(rules.length)} full-width rules, not one`);
          }
          const rule = rules[0];
          if (rule === undefined) throw new Error('unreachable');
          return {
            baseline,
            offset: rule.from / scale - baseline,
            thickness: (rule.to - rule.from + 1) / scale,
            left: rule.left / scale,
            right: (rule.right + 1) / scale,
          };
        };

        const em = (row: Record<string, number>): Record<string, number> => ({
          ...row,
          offsetEm: (row['offset'] ?? 0) / px,
          thicknessEm: (row['thickness'] ?? 0) / px,
        });
        return { underline: em(band(40, 190)), strike: em(band(240, 170)) };
      },
      { data: png, scale: DEVICE_SCALE, px: DECORATION_PX, marker: MARKER_PX },
    );
  }

  const fixture = {
    $comment: 'Chromium font metrics for experiment T8. Regenerate with measure-in-browser.ts.',
    chromium: browser.version(),
    decorationPx: DECORATION_PX,
    deviceScale: DEVICE_SCALE,
    faces: measured,
    decorations,
  };
  const out = join(dir, 'text-render-browser.json');
  writeFileSync(out, `${JSON.stringify(fixture, null, 2)}\n`);
  console.log(`${String(FACES.length)} face(s) measured -> ${out}`);
} finally {
  await browser.close();
}

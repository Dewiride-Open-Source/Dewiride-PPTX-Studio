/**
 * Experiment T3, step 1 - how wide is every prefix of every probe string?
 *
 * ```
 * node tools/ground-truth/measure-break-widths.ts <work-dir>
 * ```
 *
 * The sweep in `line-break.ts` needs a box width that holds exactly `k` code
 * points and no more, which means knowing the width of every prefix before the
 * deck can be written. Chromium answers, because T2 measured it against
 * PowerPoint over 364 comparisons and put the two a median of 0.15% apart -
 * and because half a character advance at 24pt is around 6 points, so a
 * fraction of a percent has room to be wrong in by a factor of a hundred.
 *
 * The estimate is not trusted on its own. Every probe string also gets a
 * `wrap="none"` control shape in the deck, so `analyse-breaks.ts` can compare
 * the whole-string width against PowerPoint's own and report the disagreement
 * per face rather than assume it away.
 *
 * Kerning is turned **off** on both sides - `kern="0"` in the deck, which T2
 * measured as "never kern", and `fontKerning: 'none'` here. Kerning cannot move
 * a break opportunity, only a width, so the only thing it could do to this
 * experiment is put a box a fraction of a point on the wrong side of a
 * boundary. Removing it removes that.
 *
 * `OffscreenCanvas`, not a DOM canvas: it is what the shipped measurer uses,
 * and the two are permitted to differ.
 */

import { writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

import { FACES, PROBE_SIZE, allStrings } from './line-break.ts';

/*
 * The `page.evaluate` body runs in Chromium, not in Node. `tools/` compiles with
 * `types: ["node"]` and no DOM library, deliberately, so the exact surface used
 * is declared here - narrow enough that a body which reaches for something else
 * fails to compile rather than becoming `any`. Same approach as
 * `measure-in-browser.ts`.
 */
interface OffscreenCtx {
  font: string;
  letterSpacing: string;
  fontKerning: string;
  measureText(text: string): { readonly width: number };
}
interface OffscreenCanvasCtor {
  new (width: number, height: number): { getContext(id: '2d'): OffscreenCtx | null };
}
interface FontFaceSetLike {
  check(font: string): boolean;
}

declare const OffscreenCanvas: OffscreenCanvasCtor;
declare const document: { fonts: FontFaceSetLike };

interface Job {
  readonly sizePx: number;
  readonly items: readonly {
    readonly id: string;
    readonly family: string;
    readonly text: string;
  }[];
  readonly families: readonly string[];
}

interface Result {
  readonly id: string;
  readonly family: string;
  /** `widths[k]` is the width of the first `k` code points, in points. `widths[0]` is 0. */
  readonly widths: readonly number[];
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: measure-break-widths.ts <work-dir>');
const work = resolve(dir);

const sizePt = PROBE_SIZE / 100;
const job: Job = {
  sizePx: sizePt,
  items: allStrings().map((s) => ({ id: s.id, family: FACES[s.face], text: s.text })),
  families: [...new Set(allStrings().map((s) => FACES[s.face]))],
};

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const out = await page.evaluate((j: Job) => {
    const canvas = new OffscreenCanvas(8, 8);
    const ctx = canvas.getContext('2d');
    if (ctx === null) throw new Error('no 2d context on an OffscreenCanvas');

    // Quoted, and with no generic tail: a family that is missing must fall back
    // to the browser default and stay *detectable*, which is what `installed`
    // records. T2 found `document.fonts.check` lies about faces it has never
    // seen, so the flag is reported and never used to decide anything.
    const fontOf = (family: string): string => `${String(j.sizePx)}px "${family}"`;

    const installed: Record<string, boolean> = {};
    // The advance of a real hyphen, per face. A soft-hyphen break renders one,
    // and the line it ends has to pay for it - which is only visible as a probe
    // that refuses a break the model expected.
    const hyphen: Record<string, number> = {};
    // The space advance, per face. T2 found PowerPoint reports a text range as
    // one paragraph mark wider than its text, and the mark is exactly this wide,
    // so section G cannot compare the two without it.
    const space: Record<string, number> = {};
    for (const family of j.families) {
      installed[family] = document.fonts.check(fontOf(family));
      ctx.font = fontOf(family);
      ctx.letterSpacing = '0px';
      ctx.fontKerning = 'none';
      hyphen[family] = ctx.measureText('-').width;
      space[family] = ctx.measureText(' ').width;
    }

    const results: Result[] = [];
    for (const item of j.items) {
      ctx.font = fontOf(item.family);
      ctx.letterSpacing = '0px';
      ctx.fontKerning = 'none';
      const points = [...item.text];
      const widths: number[] = [0];
      for (let k = 1; k <= points.length; k += 1) {
        widths.push(ctx.measureText(points.slice(0, k).join('')).width);
      }
      results.push({ id: item.id, family: item.family, widths });
    }
    return { installed, hyphen, space, results };
  }, job);

  writeFileSync(
    join(work, 'break-widths.json'),
    JSON.stringify({ chromium: browser.version(), sizePt, ...out }, null, 2),
  );

  const absent = Object.entries(out.installed)
    .filter(([, present]) => !present)
    .map(([family]) => family);
  console.log(`chromium ${browser.version()}`);
  console.log(`${String(out.results.length)} strings measured at ${String(sizePt)}pt`);
  console.log(`fonts.check says absent: ${absent.length === 0 ? '(none)' : absent.join(', ')}`);
} finally {
  await browser.close();
}

/**
 * Experiment T2, step 3 - ask Chromium the same questions PowerPoint answered.
 *
 * ```
 * node tools/ground-truth/text/metrics/measure-in-browser.ts <work-dir>
 * ```
 *
 * Sub-phase 3.2's deliverable is a measurer that runs in a Web Worker over
 * `OffscreenCanvas`. PowerPoint's numbers are the reference; these are what the
 * product will actually have. Two questions only this script can answer:
 *
 * 1. **Do the advances agree?** If Chromium and PowerPoint disagree about how
 *    wide `Hamburgefonstiv` is at 18pt in Arial, then no amount of correct line
 *    breaking will put the break in the same place, and the size of that
 *    disagreement is the honest ceiling on text fidelity. The plan's risk table
 *    promises "a tracked, improving number - never pixel-perfection", and this
 *    is where the number comes from.
 *
 * 2. **Is the last line's height something a browser can compute?** T2 found
 *    the last line of a text block measures `0.75 x advance + c`, where `c` is
 *    the only font-dependent term in the whole line model. If `c` is one of the
 *    metrics `TextMetrics` reports, the model is reproducible. If it is not, it
 *    has to be carried as measured data and said so out loud.
 *
 * Same `OffscreenCanvas` the Worker will use, not a DOM `<canvas>`, because the
 * two are permitted to differ and the one that matters is the one that ships.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

/*
 * The body of `page.evaluate` runs in Chromium, not in Node. `tools/` is
 * compiled with `types: ["node"]` and no DOM library, deliberately, so declare
 * exactly the surface this file uses - narrow enough that a body which starts
 * using something else fails to compile rather than becoming `any`. Same
 * approach as `verify-render.ts` and `tools/ground-truth/fonts/embedding/verify-in-browser.ts`.
 */
interface BrowserTextMetrics {
  readonly width: number;
  readonly actualBoundingBoxLeft: number;
  readonly actualBoundingBoxRight: number;
  readonly actualBoundingBoxAscent: number;
  readonly actualBoundingBoxDescent: number;
  readonly fontBoundingBoxAscent: number;
  readonly fontBoundingBoxDescent: number;
  readonly emHeightAscent: number;
  readonly emHeightDescent: number;
  readonly alphabeticBaseline: number;
  readonly hangingBaseline: number;
  readonly ideographicBaseline: number;
}
interface OffscreenCtx {
  font: string;
  letterSpacing: string;
  fontKerning: string;
  measureText(text: string): BrowserTextMetrics;
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
  readonly fonts: readonly string[];
  readonly absentFont: string;
  readonly sizes: readonly number[];
  readonly strings: readonly { readonly key: string; readonly text: string }[];
  readonly spcValues: readonly number[];
  readonly kernPairs: string;
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: measure-in-browser.ts <work-dir>');
const work = resolve(dir);

interface Inputs {
  readonly fonts: readonly string[];
  readonly absentFont: string;
  readonly strings: readonly { readonly key: string; readonly text: string }[];
  readonly probes: readonly {
    readonly kind: string;
    readonly font: string;
    readonly sz: number;
    readonly text: string;
  }[];
}
const inputs = JSON.parse(readFileSync(join(work, 'metric-inputs.json'), 'utf8')) as Inputs;

// Exactly the sizes and strings PowerPoint was asked about, taken from the
// probe table rather than restated - two lists that could drift are two lists
// that will.
const sizes = [...new Set(inputs.probes.filter((p) => p.kind === 'advance').map((p) => p.sz))].sort(
  (a, b) => a - b,
);
const lineSizes = [
  ...new Set(inputs.probes.filter((p) => p.kind === 'lastLine').map((p) => p.sz)),
].sort((a, b) => a - b);

const job: Job = {
  fonts: [...inputs.fonts, inputs.absentFont],
  absentFont: inputs.absentFont,
  sizes: [...new Set([...sizes, ...lineSizes])].sort((a, b) => a - b),
  strings: inputs.strings,
  spcValues: [0, 25, 300, -50],
  kernPairs: 'AVATAR Yo To Wa',
};

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const result = await page.evaluate((j: Job) => {
    const canvas = new OffscreenCanvas(8, 8);
    const ctx = canvas.getContext('2d');
    if (ctx === null) throw new Error('no 2d context on an OffscreenCanvas');

    // A quoted family name, so a face whose name has a space is one family and
    // not a fallback list. The generic tail is deliberately absent: a missing
    // face must fall back to the browser default and be *detectable*, which is
    // what the `installed` flag records.
    const fontOf = (family: string, px: number): string => `${String(px)}px "${family}"`;

    const out: {
      family: string;
      installed: boolean;
      sizes: {
        px: number;
        metrics: Record<string, number>;
        widths: Record<string, number>;
        spc: Record<string, number>;
        kernOn: number;
        kernOff: number;
      }[];
    }[] = [];

    for (const family of j.fonts) {
      const entry = {
        family,
        installed: document.fonts.check(fontOf(family, 16)),
        sizes: [] as {
          px: number;
          metrics: Record<string, number>;
          widths: Record<string, number>;
          spc: Record<string, number>;
          kernOn: number;
          kernOff: number;
        }[],
      };
      for (const sz of j.sizes) {
        const px = sz / 100;
        ctx.font = fontOf(family, px);
        ctx.letterSpacing = '0px';
        ctx.fontKerning = 'normal';

        const probe = ctx.measureText('Hxy');
        const metrics: Record<string, number> = {
          fontBoundingBoxAscent: probe.fontBoundingBoxAscent,
          fontBoundingBoxDescent: probe.fontBoundingBoxDescent,
          emHeightAscent: probe.emHeightAscent,
          emHeightDescent: probe.emHeightDescent,
          actualBoundingBoxAscent: probe.actualBoundingBoxAscent,
          actualBoundingBoxDescent: probe.actualBoundingBoxDescent,
          alphabeticBaseline: probe.alphabeticBaseline,
          hangingBaseline: probe.hangingBaseline,
          ideographicBaseline: probe.ideographicBaseline,
        };

        const widths: Record<string, number> = {};
        for (const s of j.strings) widths[s.key] = ctx.measureText(s.text).width;

        // Letter spacing, measured the same way the probe deck asked
        // PowerPoint: ten characters, one baseline, three offsets.
        const spc: Record<string, number> = {};
        for (const v of j.spcValues) {
          ctx.letterSpacing = `${String(v / 100)}px`;
          spc[String(v)] = ctx.measureText('0123456789').width;
        }
        ctx.letterSpacing = '0px';

        ctx.fontKerning = 'normal';
        const kernOn = ctx.measureText(j.kernPairs).width;
        ctx.fontKerning = 'none';
        const kernOff = ctx.measureText(j.kernPairs).width;
        ctx.fontKerning = 'normal';

        entry.sizes.push({ px, metrics, widths, spc, kernOn, kernOff });
      }
      out.push(entry);
    }
    return out;
  }, job);

  writeFileSync(
    join(work, 'browser-metrics.json'),
    JSON.stringify({ chromium: browser.version(), fonts: result }, null, 2),
  );

  const missing = result.filter((r) => !r.installed).map((r) => r.family);
  console.log(`chromium ${browser.version()}`);
  console.log(`${String(result.length)} families measured at ${String(job.sizes.length)} sizes`);
  console.log(`not installed: ${missing.length === 0 ? '(none)' : missing.join(', ')}`);
} finally {
  await browser.close();
}

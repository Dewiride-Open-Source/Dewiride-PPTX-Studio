/**
 * Experiment T13, step 1 - ask Chromium what it reads out of a font.
 *
 * ```
 * node tools/ground-truth/fonts/metrics/measure-in-browser.ts <work-dir>
 * ```
 *
 * The fonts are built here and loaded as data URIs, so nothing on this machine
 * is read and the answer cannot be an installed face answering for the probe -
 * every family name is one that exists nowhere. ADR 0042.
 */

import { writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

import { BASELINE_SIZES, BOX_PX, SIZES, STRINGS, probeFonts } from './probes.ts';

/*
 * `page.evaluate` runs in Chromium; `tools/` is compiled with `types: ["node"]`
 * and no DOM library, so the surface used is declared rather than imported.
 */
interface Ctx2D {
  font: string;
  letterSpacing: string;
  fontKerning: string;
  measureText(text: string): {
    readonly width: number;
    readonly fontBoundingBoxAscent: number;
    readonly fontBoundingBoxDescent: number;
    readonly alphabeticBaseline: number;
    readonly hangingBaseline: number;
    readonly ideographicBaseline: number;
  };
}
interface CanvasLike {
  getContext(id: '2d'): Ctx2D | null;
}
interface FontFaceLike {
  load(): Promise<FontFaceLike>;
}
declare const OffscreenCanvas: new (w: number, h: number) => CanvasLike;
declare const FontFace: new (family: string, source: string) => FontFaceLike;
declare const document: { fonts: { add(face: FontFaceLike): void } };

interface Job {
  readonly fonts: readonly {
    id: string;
    family: string;
    base64: string;
    ideograph: number | undefined;
  }[];
  readonly strings: readonly string[];
  readonly sizes: readonly number[];
  readonly boxPx: number;
  readonly baselineSizes: readonly number[];
}

/** What one `measureText` says about where the baselines sit. */
interface Baselines {
  alphabetic: number;
  hanging: number;
  ideographic: number;
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: measure-in-browser.ts <work-dir>');
const work = resolve(dir);

const fonts = probeFonts();
const job: Job = {
  fonts: fonts.map((f) => ({
    id: f.id,
    family: f.family,
    base64: f.base64,
    ideograph: f.spec.ideograph,
  })),
  strings: STRINGS,
  sizes: SIZES,
  boxPx: BOX_PX,
  baselineSizes: BASELINE_SIZES,
};

// No PINNED_ARGS here: `--disable-remote-fonts` would refuse every probe.
const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const measured = await page.evaluate(async (j: Job) => {
    const ctx = new OffscreenCanvas(8, 8).getContext('2d');
    if (ctx === null) throw new Error('no 2d context on an OffscreenCanvas');

    const out: {
      id: string;
      loaded: boolean;
      widths: Record<string, Record<string, number>>;
      box: { ascent: number; descent: number };
      baselines: Record<string, Baselines>;
      hanBaselines: Record<string, Baselines & { width: number }> | null;
    }[] = [];

    for (const font of j.fonts) {
      const face = new FontFace(font.family, `url(data:font/ttf;base64,${font.base64})`);
      let loaded = true;
      try {
        document.fonts.add(await face.load());
      } catch {
        loaded = false;
      }

      const widths: Record<string, Record<string, number>> = {};
      for (const px of j.sizes) {
        ctx.font = `${String(px)}px "${font.family}"`;
        ctx.letterSpacing = '0px';
        ctx.fontKerning = 'normal';
        const row: Record<string, number> = {};
        for (const text of j.strings) row[text] = ctx.measureText(text).width;
        widths[String(px)] = row;
      }

      ctx.font = `${String(j.boxPx)}px "${font.family}"`;
      const metrics = ctx.measureText('AH');

      const read = (m: ReturnType<Ctx2D['measureText']>): Baselines => ({
        alphabetic: m.alphabeticBaseline,
        hanging: m.hangingBaseline,
        ideographic: m.ideographicBaseline,
      });
      const baselines: Record<string, Baselines> = {};
      // Only a font that maps U+4E00 is asked about it: a font that does not
      // would be answered by whatever face the machine has, which is the one
      // thing this suite refuses.
      const hanBaselines: Record<string, Baselines & { width: number }> | null =
        font.ideograph === undefined ? null : {};
      for (const px of j.baselineSizes) {
        ctx.font = `${String(px)}px "${font.family}"`;
        ctx.letterSpacing = '0px';
        ctx.fontKerning = 'normal';
        baselines[String(px)] = read(ctx.measureText('A'));
        if (hanBaselines !== null && font.ideograph !== undefined) {
          const han = ctx.measureText(String.fromCodePoint(font.ideograph));
          hanBaselines[String(px)] = { ...read(han), width: han.width };
        }
      }

      out.push({
        id: font.id,
        loaded,
        widths,
        box: {
          ascent: metrics.fontBoundingBoxAscent,
          descent: metrics.fontBoundingBoxDescent,
        },
        baselines,
        hanBaselines,
      });
    }
    return out;
  }, job);

  const unloaded = measured.filter((m) => !m.loaded).map((m) => m.id);
  if (unloaded.length > 0) {
    throw new Error(`chromium refused these probe fonts: ${unloaded.join(', ')}`);
  }

  writeFileSync(
    join(work, 'browser-metrics.json'),
    `${JSON.stringify(
      {
        chromium: browser.version(),
        boxPx: BOX_PX,
        sizes: SIZES,
        baselineSizes: BASELINE_SIZES,
        strings: STRINGS,
        fonts: fonts.map((f) => ({ id: f.id, asks: f.asks, family: f.family, spec: f.spec })),
        measured,
      },
      null,
      2,
    )}\n`,
  );

  console.log(`chromium ${browser.version()}`);
  for (const row of measured) {
    const at100 = row.widths['100'] ?? {};
    const base = row.baselines['1000'];
    const han = row.hanBaselines?.['1000'];
    console.log(
      `${row.id.padEnd(16)} box ${String(row.box.ascent)}/${String(row.box.descent)}` +
        `  A=${String(at100['A'])} AB=${String(at100['AB'])}` +
        `  ideo=${String(base?.ideographic)} hang=${String(base?.hanging)}` +
        `  han ideo=${String(han?.ideographic)} width=${String(han?.width)}`,
    );
  }
  console.log(`\nwrote ${join(work, 'browser-metrics.json')}`);
} finally {
  await browser.close();
}

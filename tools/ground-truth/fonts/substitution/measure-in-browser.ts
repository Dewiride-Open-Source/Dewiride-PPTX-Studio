/**
 * Experiment T7, step 3 - ask Chromium which of these faces it actually has.
 *
 * ```
 * node tools/ground-truth/fonts/substitution/measure-in-browser.ts <work-dir>
 * ```
 *
 * PowerPoint's answer is the reference; this is what the product will have. The
 * detector has to be scored here or it is not measured at all. ADR 0033.
 */

import { writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { chromium } from 'playwright';

import { ABSENT_FACES, ALIAS_FACES, INSTALLED_FACES } from './probes.ts';

/*
 * `page.evaluate` runs in Chromium; `tools/` is compiled with `types: ["node"]`
 * and no DOM library, so the surface used is declared rather than imported.
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

/**
 * The pool the analysis picks the fingerprint's three glyphs out of, by
 * exhaustive search - a triple chosen by eye is a triple nothing refutes.
 */
const GLYPHS: readonly string[] = [
  ...'abcdefghijklmnopqrstuvwxyz',
  ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
  ...'0123456789',
  'Hamburgefonstiv',
];

/** A generic a family is stacked on top of, so absence is observable. */
const ANCHORS: readonly string[] = ['monospace', 'serif', 'sans-serif'];

/** Two sizes, so "the fingerprint is scale-free" is checked rather than assumed. */
const SIZES: readonly number[] = [100, 16];

interface Job {
  readonly families: readonly string[];
  readonly glyphs: readonly string[];
  readonly anchors: readonly string[];
  readonly sizes: readonly number[];
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: measure-in-browser.ts <work-dir>');
const work = resolve(dir);

const families = [...ABSENT_FACES, ...INSTALLED_FACES, ...ALIAS_FACES];
const job: Job = { families, glyphs: GLYPHS, anchors: ANCHORS, sizes: SIZES };

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const result = await page.evaluate((j: Job) => {
    const ctx = new OffscreenCanvas(8, 8).getContext('2d');
    if (ctx === null) throw new Error('no 2d context on an OffscreenCanvas');

    const widths = (stack: string, px: number): number[] => {
      ctx.font = `${String(px)}px ${stack}`;
      ctx.letterSpacing = '0px';
      ctx.fontKerning = 'normal';
      return j.glyphs.map((g) => ctx.measureText(g).width);
    };
    const bySize = (stack: string): Record<string, number[]> =>
      Object.fromEntries(j.sizes.map((px) => [String(px), widths(stack, px)]));

    // The anchors alone: where a family that resolves to nothing ends up.
    const anchorOnly: Record<string, Record<string, number[]>> = {};
    for (const anchor of j.anchors) anchorOnly[anchor] = bySize(anchor);

    const out = j.families.map((family) => ({
      family,
      check: document.fonts.check(`16px "${family}"`),
      bare: bySize(`"${family}"`),
      stacked: Object.fromEntries(
        j.anchors.map((anchor) => [anchor, bySize(`"${family}", ${anchor}`)]),
      ),
    }));

    return { anchorOnly, families: out };
  }, job);

  writeFileSync(
    join(work, 'browser-fonts.json'),
    `${JSON.stringify(
      {
        chromium: browser.version(),
        sizes: SIZES,
        glyphs: GLYPHS,
        anchors: ANCHORS,
        anchorOnly: result.anchorOnly,
        families: result.families,
      },
      null,
      2,
    )}\n`,
  );

  const checked = result.families.filter((f) => f.check).length;
  console.log(`chromium ${browser.version()}`);
  console.log(
    `${String(result.families.length)} families x ${String(GLYPHS.length)} glyphs at ${SIZES.map(String).join('/')}px`,
  );
  console.log(`document.fonts.check said yes to ${String(checked)} of them`);
} finally {
  await browser.close();
}

/**
 * Experiment T14, step 1 - the shipped reader and Chromium, on the same bytes.
 *
 * ```
 * node tools/ground-truth/fonts/real-faces/measure-in-browser.ts <work-dir> \
 *   --font-dir <d> [--font-dir <d>]
 * ```
 *
 * `--font-dir` is required and repeatable and the index is taken with
 * `system: false`, so there is no branch here that can reach a machine's own
 * font directories. Each face is served to the page from `page.route`, so the
 * browser reads the bytes the reader read and no fontconfig alias can enter the
 * comparison. ADR 0042, open questions 2 and 3.
 */

import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { platform, release } from 'node:os';
import { extname, join, resolve } from 'node:path';

import { chromium } from 'playwright';

import {
  createFontMeasurer,
  indexFonts,
  type Face,
  type FontLibrary,
} from '../../../../packages/cli/dist/index.js';
import { AGREEMENT, BOX_PX, SAMPLES, SIZES, sampleCharacters, type Reading } from './probes.ts';
import type { Row, Run, RunFace } from './score.ts';

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
  };
}
interface CanvasLike {
  getContext(id: '2d'): Ctx2D | null;
}
interface FontFaceLike {
  load(): Promise<FontFaceLike>;
}
declare const OffscreenCanvas: new (width: number, height: number) => CanvasLike;
declare const FontFace: new (family: string, source: string) => FontFaceLike;
declare const document: {
  fonts: {
    add(face: FontFaceLike): void;
    delete(face: FontFaceLike): void;
    check(font: string): boolean;
  };
};

/** An origin that resolves nowhere, so a missed route cannot reach the network. */
const ORIGIN = 'https://real-faces.invalid';

const HTML_TYPE = 'text/html; charset=utf-8';

/** The fixed-point step T13 measured Chromium reporting an advance in. */
const FIXED = 65536;

function truncateAt(px: number): number {
  return Math.trunc(px * FIXED) / FIXED;
}

function roundAt(px: number): number {
  return Math.round(px * FIXED) / FIXED;
}

/* -------------------------------------------------------------------------- */
/* the command line                                                           */
/* -------------------------------------------------------------------------- */

const USAGE = 'usage: measure-in-browser.ts <work-dir> --font-dir <d> [--font-dir <d>]';

const args = process.argv.slice(2);
const work = args[0];
if (work === undefined || work.startsWith('--')) throw new Error(USAGE);

const directories: string[] = [];
for (let i = 1; i < args.length; i += 1) {
  if (args[i] !== '--font-dir') throw new Error(`${USAGE}\nunexpected argument ${String(args[i])}`);
  const named = args[i + 1];
  if (named === undefined) throw new Error(`${USAGE}\n--font-dir needs a directory`);
  directories.push(resolve(named));
  i += 1;
}
if (directories.length === 0) throw new Error(`${USAGE}\nat least one --font-dir is required`);

const workDir = resolve(work);
mkdirSync(workDir, { recursive: true });

/* -------------------------------------------------------------------------- */
/* what the reader says                                                       */
/* -------------------------------------------------------------------------- */

interface Rivals {
  readonly exact: number;
  readonly advanceRounded: number;
  readonly unkerned: number;
  readonly wholeStringTruncated: number;
}

/** The four candidate readings, from the same `Face` the product would use. */
function rivals(face: Face, text: string, px: number): Rivals {
  const em = face.metrics.unitsPerEm;
  const points = [...text];
  let exact = 0;
  let advanceRounded = 0;
  let unkerned = 0;
  for (let i = 0; i < points.length; i += 1) {
    const code = points[i]?.codePointAt(0) ?? 0;
    const advance = ((face.advanceOf(code) ?? 0) * px) / em;
    exact += advance;
    advanceRounded += roundAt(advance);
    unkerned += truncateAt(advance);
    const following = points[i + 1];
    if (following === undefined) continue;
    const adjust = face.kernBetween(code, following.codePointAt(0) ?? 0);
    if (adjust === 0) continue;
    const kern = (adjust * px) / em;
    exact += kern;
    advanceRounded += roundAt(kern);
  }
  return { exact, advanceRounded, unkerned, wholeStringTruncated: truncateAt(exact) };
}

/** One face, presented as a whole library, so no cross-face fallback can fire. */
function libraryOf(face: Face, file: string, family: string): FontLibrary {
  return {
    indexed: [{ face, file }],
    directories,
    resolve: (asked: string) => ({ face, asked, drawn: family, file, substituted: false }),
  };
}

/* -------------------------------------------------------------------------- */
/* which files can be asked                                                   */
/* -------------------------------------------------------------------------- */

const library = indexFonts({ extra: directories, system: false });

const byFile = new Map<string, Face[]>();
for (const entry of library.indexed) {
  const bucket = byFile.get(entry.file);
  if (bucket === undefined) byFile.set(entry.file, [entry.face]);
  else bucket.push(entry.face);
}

const skipped: { file: string; why: string }[] = [];
const candidates: { file: string; face: Face }[] = [];
const COLLECTION = new Set(['.ttc', '.otc']);
for (const [file, faces] of byFile) {
  const first = faces[0];
  // CSS cannot address a collection by index, so the browser and the reader
  // would not be looking at the same face.
  if (faces.length > 1 || COLLECTION.has(extname(file).toLowerCase()) || first === undefined) {
    skipped.push({ file, why: 'collection' });
    continue;
  }
  candidates.push({ file, face: first });
}
candidates.sort((a, b) => a.file.localeCompare(b.file));

const ALL_CHARACTERS = sampleCharacters();

/** A code point this face certainly has, so the box is not a fallback's box. */
function boxCharacter(face: Face): string | undefined {
  for (const code of ALL_CHARACTERS) {
    if (face.advanceOf(code) !== undefined) return String.fromCodePoint(code);
  }
  return undefined;
}

/** The samples every one of whose code points this face draws itself. */
function covered(face: Face): typeof SAMPLES {
  return SAMPLES.filter((sample) =>
    [...sample.text].every(
      (character) => face.advanceOf(character.codePointAt(0) ?? 0) !== undefined,
    ),
  );
}

interface Asked {
  readonly file: string;
  readonly face: Face;
  readonly family: string;
  readonly boxText: string;
  readonly samples: typeof SAMPLES;
}

const asked: Asked[] = [];
for (const candidate of candidates) {
  const boxText = boxCharacter(candidate.face);
  if (boxText === undefined) {
    skipped.push({ file: candidate.file, why: 'no measurable glyph' });
    continue;
  }
  asked.push({
    file: candidate.file,
    face: candidate.face,
    family: `t14f${String(asked.length)}`,
    boxText,
    samples: covered(candidate.face),
  });
}

/* -------------------------------------------------------------------------- */
/* what the browser says                                                      */
/* -------------------------------------------------------------------------- */

interface Job {
  readonly family: string;
  readonly url: string;
  readonly boxPx: number;
  readonly boxText: string;
  readonly sizes: readonly number[];
  readonly texts: readonly { readonly id: string; readonly text: string }[];
}

interface Measured {
  readonly loaded: boolean;
  readonly checked: boolean;
  readonly widths: Record<string, Record<string, number>>;
  readonly box: { ascent: number; descent: number };
}

const routes = new Map(asked.map((entry) => [`/font/${entry.family}`, entry.file]));

// Default flags on purpose: `PINNED_ARGS` carries `--disable-remote-fonts`,
// which would refuse every face here.
const browser = await chromium.launch();
const faces: RunFace[] = [];
try {
  const page = await browser.newPage();
  await page.route(`${ORIGIN}/**`, async (route) => {
    const { pathname } = new URL(route.request().url());
    if (pathname === '/harness.html') {
      await route.fulfill({
        status: 200,
        contentType: HTML_TYPE,
        body: '<!doctype html><meta charset="utf-8"><title>real-faces</title>',
      });
      return;
    }
    const file = routes.get(pathname);
    if (file === undefined) {
      await route.fulfill({ status: 404, body: 'not a face this run asked for' });
      return;
    }
    const type = extname(file).toLowerCase() === '.otf' ? 'font/otf' : 'font/ttf';
    await route.fulfill({ status: 200, contentType: type, body: readFileSync(file) });
  });
  await page.goto(`${ORIGIN}/harness.html`);

  for (const entry of asked) {
    const job: Job = {
      family: entry.family,
      url: `${ORIGIN}/font/${entry.family}`,
      boxPx: BOX_PX,
      boxText: entry.boxText,
      sizes: SIZES,
      texts: entry.samples.map((sample) => ({ id: sample.id, text: sample.text })),
    };

    const measured: Measured = await page.evaluate(async (j: Job) => {
      const ctx = new OffscreenCanvas(8, 8).getContext('2d');
      if (ctx === null) throw new Error('no 2d context on an OffscreenCanvas');

      const face = new FontFace(j.family, `url("${j.url}")`);
      let loaded = true;
      try {
        document.fonts.add(await face.load());
      } catch {
        loaded = false;
      }
      // A face that never loaded still measures, in the platform default, and
      // those numbers look plausible - so absence is asserted, not assumed.
      const checked = loaded && document.fonts.check(`16px "${j.family}"`);

      const widths: Record<string, Record<string, number>> = {};
      let box = { ascent: 0, descent: 0 };
      if (checked) {
        for (const px of j.sizes) {
          ctx.font = `${String(px)}px "${j.family}"`;
          ctx.letterSpacing = '0px';
          ctx.fontKerning = 'normal';
          const row: Record<string, number> = {};
          for (const sample of j.texts) row[sample.id] = ctx.measureText(sample.text).width;
          widths[String(px)] = row;
        }
        ctx.font = `${String(j.boxPx)}px "${j.family}"`;
        const metrics = ctx.measureText(j.boxText);
        box = { ascent: metrics.fontBoundingBoxAscent, descent: metrics.fontBoundingBoxDescent };
      }

      if (loaded) document.fonts.delete(face);
      return { loaded, checked, widths, box };
    }, job);

    if (!measured.checked) {
      skipped.push({ file: entry.file, why: 'browser refused' });
      continue;
    }

    const reader = createFontMeasurer(libraryOf(entry.face, entry.file, entry.family));
    const rows: Row[] = [];
    for (const px of SIZES) {
      for (const sample of entry.samples) {
        const shipped = reader.measurer.measure(sample.text, {
          family: entry.family,
          sz: px * 100,
        });
        const rest = rivals(entry.face, sample.text, px);
        const readings: Record<Reading, number> = { shipped: shipped.width, ...rest };
        rows.push({
          sample: sample.id,
          px,
          browser: measured.widths[String(px)]?.[sample.id] ?? Number.NaN,
          readings,
        });
      }
    }
    // The coverage filter and the measurer have to agree about this face, or
    // one of them is looking at glyphs the other is not.
    const absent = reader.missing();
    if (absent.length > 0) {
      throw new Error(
        `${entry.file}: the measurer found ${String(absent.length)} code point(s) the ` +
          'coverage filter said were present',
      );
    }

    faces.push({
      file: entry.file,
      sha256: createHash('sha256').update(readFileSync(entry.file)).digest('hex'),
      family: entry.face.family,
      subfamily: entry.face.subfamily,
      unitsPerEm: entry.face.metrics.unitsPerEm,
      metricsSource: entry.face.metrics.source,
      box: {
        browser: measured.box,
        reader: { ascent: entry.face.metrics.ascent, descent: entry.face.metrics.descent },
      },
      uncovered: SAMPLES.filter((sample) => !entry.samples.includes(sample)).map(
        (sample) => sample.id,
      ),
      rows,
    });
  }

  const run: Run = {
    experiment: 'T14',
    subPhase: '3.10',
    adr: 'docs/adr/phase-3-text/0042-rendering-without-a-browser.md',
    chromium: browser.version(),
    image: {
      platform: platform(),
      release: release(),
      imageOs: process.env['ImageOS'] ?? null,
      imageVersion: process.env['ImageVersion'] ?? null,
    },
    directories: library.directories,
    sizes: SIZES,
    boxPx: BOX_PX,
    agreement: AGREEMENT,
    samples: SAMPLES.map((sample) => ({
      id: sample.id,
      script: sample.script,
      gated: sample.gated,
      shaped: sample.shaped,
      asks: sample.asks,
      text: sample.text,
    })),
    indexed: { files: byFile.size, faces: library.indexed.length, measured: faces.length },
    skipped,
    faces,
  };

  const at = join(workDir, 'real-faces.json');
  writeFileSync(at, `${JSON.stringify(run, null, 2)}\n`);
  console.log(`chromium ${run.chromium}`);
  console.log(
    `${String(run.indexed.measured)} of ${String(run.indexed.faces)} face(s) in ` +
      `${String(run.indexed.files)} file(s) from ${directories.join(', ')}`,
  );
  console.log(
    `${String(skipped.length)} skipped, ${String(faces.reduce((n, f) => n + f.rows.length, 0))} comparisons`,
  );
  console.log(`wrote ${at}`);
} finally {
  await browser.close();
}

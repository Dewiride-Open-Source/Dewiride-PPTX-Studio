/**
 * What `@pptx-studio/cli` makes of every sample deck, rendered ahead of time.
 *
 * ```
 * node prerender/decks.mjs
 * ```
 *
 * A static export has no server, so the CLI runs here, on the machine that
 * builds the site: `renderDeck` over `public/decks/*` into `public/rendered/`,
 * with the font report and the `inspect` and `validate` output beside each
 * deck. The page shows exactly this, labelled with the machine that made it.
 * Throws on the first failure rather than leaving a partial directory.
 */

import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { main, renderDeck } from '@pptx-studio/cli';

const WIDTH = 1280;
const DECK = /\.ppt[xm]$/;
const DECKS = join(import.meta.dirname, '..', 'public', 'decks');
const OUT = join(import.meta.dirname, '..', 'public', 'rendered');

/** @param {string} verb @param {string} file */
function capture(verb, file) {
  let text = '';
  const exitCode = main([verb, file], {
    out: (chunk) => {
      text += chunk;
    },
    err: (chunk) => {
      text += chunk;
    },
  });
  return { text, exitCode };
}

/** @param {number} code */
const codePoint = (code) => `U+${code.toString(16).toUpperCase().padStart(4, '0')}`;

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });

const files = readdirSync(DECKS)
  .filter((name) => DECK.test(name))
  .sort();
if (files.length === 0) throw new Error(`no decks under ${DECKS}`);

const decks = [];
for (const file of files) {
  const startedAt = performance.now();
  const bytes = new Uint8Array(readFileSync(join(DECKS, file)));
  const rendered = renderDeck(bytes, { width: WIDTH });
  const ms = performance.now() - startedAt;

  const id = file.replace(DECK, '');
  const dir = join(OUT, id);
  mkdirSync(dir);
  const slides = rendered.slides.map((slide) => {
    const name = `slide-${String(slide.number).padStart(3, '0')}.svg`;
    writeFileSync(join(dir, name), slide.svg);
    return { number: slide.number, file: name, bytes: Buffer.byteLength(slide.svg) };
  });

  const inspect = capture('inspect', join(DECKS, file));
  const validate = capture('validate', join(DECKS, file));
  writeFileSync(join(dir, 'inspect.txt'), inspect.text);
  writeFileSync(join(dir, 'validate.txt'), validate.text);

  const report = {
    id,
    file,
    width: rendered.width,
    height: rendered.height,
    slides,
    fonts: rendered.fonts,
    missing: rendered.missing.map(codePoint),
    fontDirectories: rendered.fontDirectories,
    facesIndexed: rendered.facesIndexed,
    ms: Math.round(ms),
    validateExitCode: validate.exitCode,
  };
  writeFileSync(join(dir, 'render.json'), `${JSON.stringify(report, null, 2)}\n`);
  decks.push({ id, file, slides: slides.length, ms: report.ms });
  console.log(
    `${file}: ${String(slides.length)} slide(s) in ${String(report.ms)} ms, ` +
      `${String(rendered.fonts.filter((face) => face.substituted).length)} substituted`,
  );
}

const index = {
  width: WIDTH,
  command: `npx @pptx-studio/cli render <deck> --width ${String(WIDTH)} --out rendered/`,
  machine: { node: process.version, platform: process.platform, arch: process.arch },
  decks,
};
writeFileSync(join(OUT, 'index.json'), `${JSON.stringify(index, null, 2)}\n`);
console.log(`${String(decks.length)} deck(s) rendered into public/rendered`);

/**
 * `pptx-studio render` - slides to SVG, with no browser anywhere.
 *
 * Every part of the picture except text was already reachable from Node:
 * geometry, fills, strokes, effects and images are pure functions over the
 * model, and an image's size comes out of its own header rather than a decoder.
 * Text was the one thing that needed a canvas, and `measure.ts` is what replaces
 * it. ADR 0042.
 */

import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { renderSlide, type MediaResolver } from '@pptx-studio/render-svg';

import { RenderError } from './errors.js';
import { indexFonts, type FontLibrary } from './faces.js';
import { createFontMeasurer, type FaceUse } from './measure.js';

/** What a slide is drawn at when the caller says nothing. PowerPoint's own. */
export const DEFAULT_WIDTH = 1920;

export interface RenderOptions {
  /** 1-based, or `null` for every slide. */
  readonly slide: number | null;
  readonly width: number;
  /** A directory for many slides, a file for one, or `null` for stdout. */
  readonly out: string | null;
  readonly fontDirs: readonly string[];
  readonly systemFonts: boolean;
  readonly text: boolean;
  readonly json: boolean;
  readonly quiet: boolean;
}

export const RENDER_DEFAULTS = {
  width: DEFAULT_WIDTH,
  systemFonts: true,
  text: true,
} as const;

export interface RenderedSlide {
  /** 1-based, as the flag and the file name spell it. */
  readonly number: number;
  readonly svg: string;
}

export interface RenderResult {
  readonly slides: readonly RenderedSlide[];
  readonly width: number;
  readonly height: number;
  /** Every typeface the deck asked for, and what drew it. Empty with `--no-text`. */
  readonly fonts: readonly FaceUse[];
  /** Code points no indexed face could draw. */
  readonly missing: readonly number[];
  readonly fontDirectories: readonly string[];
  readonly facesIndexed: number;
}

/** The media resolver: an rId means the rels of the part the fill was written in. */
function mediaFrom(store: PartStore): MediaResolver {
  return (embed: string, part: string) => {
    const target = store.relationships(part).targetOf(embed);
    if (target === undefined) return undefined;
    const contentType = store.contentTypeOf(target);
    if (contentType === undefined) return undefined;
    return { bytes: store.read(target), contentType };
  };
}

/**
 * Render a deck that is already in memory.
 *
 * Separate from `runRender` so that the whole pipeline is exercised by the test
 * suite without a file system, a process or a captured stdout.
 */
export function renderDeck(bytes: Uint8Array, options: RenderOptions): RenderResult {
  const store = PartStore.open(bytes);
  const document = loadDocument(store);
  const size = document.slideSize;
  const height = Math.round((options.width * size.cy) / size.cx);

  if (options.slide !== null) {
    const count = document.slides.length;
    if (!Number.isInteger(options.slide) || options.slide < 1 || options.slide > count) {
      throw new RenderError(
        'CLI_NO_SLIDE',
        `--slide ${String(options.slide)}: the deck has ${String(count)} slide(s)`,
        String(options.slide),
      );
    }
  }

  let library: FontLibrary | null = null;
  const fonts = options.text
    ? createFontMeasurer(
        (library = indexFonts({ extra: options.fontDirs, system: options.systemFonts })),
      )
    : null;

  const media = mediaFrom(store);
  const wanted =
    options.slide === null
      ? document.slides.map((sheet, at) => ({ sheet, number: at + 1 }))
      : [{ sheet: document.slides[options.slide - 1]!, number: options.slide }];

  const slides = wanted.map(({ sheet, number }) => ({
    number,
    svg: renderSlide(sheet, size, {
      width: options.width,
      height,
      idPrefix: `s${String(number)}`,
      media,
      text:
        fonts === null
          ? false
          : {
              defaultTextStyle: document.defaultTextStyle,
              measurer: fonts.measurer,
              faceBox: fonts.faceBox,
            },
    }),
  }));

  return {
    slides,
    width: options.width,
    height,
    fonts: fonts?.used() ?? [],
    missing: fonts?.missing() ?? [],
    fontDirectories: library?.directories ?? [],
    facesIndexed: library?.indexed.length ?? 0,
  };
}

/** `slide-03.svg`, so a directory listing sorts the way the deck reads. */
function fileNameFor(number: number, count: number): string {
  const width = Math.max(2, String(count).length);
  return `slide-${String(number).padStart(width, '0')}.svg`;
}

function isDirectory(path: string): boolean {
  try {
    return statSync(path).isDirectory();
  } catch {
    return false;
  }
}

/**
 * Where each slide's markup goes.
 *
 * One slide to a path that is not an existing directory writes that file;
 * anything else writes a file per slide into a directory, created if needed.
 * A path ending in `.svg` with more than one slide is a mistake worth naming
 * rather than silently turning into a directory.
 */
function write(result: RenderResult, out: string): readonly string[] {
  const single = result.slides.length === 1 && !isDirectory(out);
  if (single) {
    const first = result.slides[0]!;
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, first.svg, 'utf8');
    return [out];
  }
  if (out.toLowerCase().endsWith('.svg')) {
    throw new RenderError(
      'CLI_OUTPUT_PATH',
      `--out ${out} names one file but ${String(result.slides.length)} slides were rendered; ` +
        'name a directory, or pass --slide',
      out,
    );
  }
  mkdirSync(out, { recursive: true });
  return result.slides.map((slide) => {
    const path = join(out, fileNameFor(slide.number, result.slides.length));
    writeFileSync(path, slide.svg, 'utf8');
    return path;
  });
}

function report(result: RenderResult, written: readonly string[], options: RenderOptions): string {
  const lines: string[] = [];
  const slides = `${String(result.slides.length)} slide(s) at ${String(result.width)}x${String(result.height)}`;
  lines.push(
    written.length === 0
      ? slides
      : `${slides} -> ${written.length === 1 ? written[0]! : `${String(written.length)} files`}`,
  );

  if (options.text) {
    const substituted = result.fonts.filter((font) => font.substituted);
    lines.push(
      `${String(result.fonts.length)} typeface(s) from ${String(result.facesIndexed)} indexed face(s)` +
        (substituted.length === 0 ? '' : `, ${String(substituted.length)} substituted`),
    );
    for (const font of substituted) lines.push(`  ${font.asked} -> ${font.drawn}`);
    if (result.missing.length > 0) {
      const shown = result.missing
        .slice(0, 8)
        .map((code) => `U+${code.toString(16).toUpperCase().padStart(4, '0')}`);
      lines.push(
        `  no face has ${String(result.missing.length)} code point(s): ${shown.join(' ')}` +
          (result.missing.length > shown.length ? ' ...' : ''),
      );
    }
  }
  return `${lines.join('\n')}\n`;
}

/** Read, render, write, and return an exit code. Never calls `process.exit`. */
export function runRender(
  file: string,
  options: RenderOptions,
  out: (text: string) => void,
): number {
  const result = renderDeck(new Uint8Array(readFileSync(file)), options);

  if (options.json) {
    const written = options.out === null ? [] : write(result, options.out);
    out(
      `${JSON.stringify(
        {
          slides: result.slides.map((slide) => ({
            number: slide.number,
            bytes: slide.svg.length,
          })),
          width: result.width,
          height: result.height,
          fonts: result.fonts,
          missing: result.missing.map((code) => `U+${code.toString(16).toUpperCase()}`),
          fontDirectories: result.fontDirectories,
          facesIndexed: result.facesIndexed,
          written,
        },
        null,
        2,
      )}\n`,
    );
    return 0;
  }

  if (options.out === null) {
    for (const slide of result.slides) out(slide.svg);
    return 0;
  }

  const written = write(result, options.out);
  if (!options.quiet) out(report(result, written, options));
  return 0;
}

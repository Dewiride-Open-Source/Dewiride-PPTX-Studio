/**
 * A slide, and PowerPoint's picture of it, through the same rasteriser.
 *
 * Both sides are decoded by the same Chromium, through the same `<img>` and
 * `drawImage`, and reduced by the same injected `reduceRgba`. So there is no
 * second decode path to keep in agreement and no PNG decoder in this
 * repository at all. ADR 0035.
 */

import type { Page } from 'playwright';

import { FidelityError } from '../errors.ts';
import { reduceRgba, type Grid } from '../metric/reduce.ts';

/*
 * The bodies of `page.evaluate` run in Chromium, not in Node, and `tools/` is
 * compiled with `types: ["node"]` and no DOM library. As in
 * `tools/ground-truth/render/transforms/verify-render.ts`, declare exactly the
 * surface these two functions use: narrow enough that a body which starts using
 * something else fails to compile instead of becoming `any`.
 */
interface Ctx2D {
  fillStyle: string;
  fillRect(x: number, y: number, w: number, h: number): void;
  drawImage(image: BrowserImage, x: number, y: number, w: number, h: number): void;
  getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray };
}
interface Canvas {
  width: number;
  height: number;
  getContext(id: '2d', options?: { willReadFrequently: boolean }): Ctx2D | null;
}
interface BrowserImage {
  src: string;
  readonly naturalWidth: number;
  readonly naturalHeight: number;
  decode(): Promise<void>;
}
declare const document: { createElement(tag: 'canvas'): Canvas };
declare const Image: new () => BrowserImage;

/** The side of one luma cell, in raster pixels. */
export const CELL = 8;

/** Every slide is drawn this wide, whatever its aspect. */
export const RASTER_WIDTH = 960;

export interface Geometry {
  /** What the SVG is asked to draw at, which is its own aspect. */
  readonly drawWidth: number;
  readonly drawHeight: number;
  /** The canvas, grown to whole cells and filled white behind the slide. */
  readonly padWidth: number;
  readonly padHeight: number;
}

/**
 * The raster a slide of this shape is compared at.
 *
 * The width is fixed and the height follows the slide, so a 4:3 deck is not
 * squashed into a 16:9 frame; the canvas is then grown to whole cells, because
 * a partial cell would average over pixels that do not exist.
 */
export function geometryOf(cx: number, cy: number): Geometry {
  if (!(cx > 0) || !(cy > 0)) {
    throw new FidelityError('FID_RASTER_GEOMETRY', `slide is ${String(cx)} by ${String(cy)} EMU`);
  }
  const drawHeight = Math.round((RASTER_WIDTH * cy) / cx);
  return {
    drawWidth: RASTER_WIDTH,
    drawHeight,
    padWidth: Math.ceil(RASTER_WIDTH / CELL) * CELL,
    padHeight: Math.ceil(drawHeight / CELL) * CELL,
  };
}

/**
 * Install the reduction in the page, as the same source Node runs.
 *
 * A script tag rather than `new Function`: the source crossing into the page is
 * this module's own compiled function and nothing else, so there is no string
 * here that anything outside the repository could reach.
 */
export async function injectReduce(page: Page): Promise<void> {
  await page.addScriptTag({ content: `globalThis.reduce = ${reduceRgba.toString()};` });
}

export interface SlideRaster {
  readonly grid: Grid;
  /** SHA-256 of the padded RGBA, which is what the gate compares. */
  readonly rasterSha256: string;
  /** SHA-256 of the SVG source, which localises a change to markup or paint. */
  readonly svgSha256: string;
  /** The CSS families this slide was drawn in, taken from the markup itself. */
  readonly families: readonly string[];
  readonly geometry: Geometry;
}

/**
 * Draw one slide and reduce it.
 *
 * The digest is taken inside the page: shipping eight million bytes per slide
 * over the debugging protocol is what made an earlier version of this idea take
 * longer than ten minutes on fifty decks.
 */
export async function renderSlide(
  page: Page,
  deckUrl: string,
  slideIndex: number,
): Promise<SlideRaster> {
  const result = await page.evaluate(
    async ({ url, index, cell, width }) => {
      const api = (globalThis as unknown as { pptx: Record<string, never> }).pptx as unknown as {
        opc: { PartStore: { open: (bytes: Uint8Array) => unknown } };
        model: {
          loadDocument: (store: unknown) => {
            slides: readonly unknown[];
            slideSize: { cx: number; cy: number };
            defaultTextStyle: unknown;
          };
        };
        rsvg: { renderSlide: (sheet: unknown, size: unknown, options: unknown) => string };
      };
      const reduce = (globalThis as unknown as { reduce: (input: unknown) => unknown }).reduce;

      const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
      const document_ = api.model.loadDocument(api.opc.PartStore.open(bytes));
      const size = document_.slideSize;
      const drawHeight = Math.round((width * size.cy) / size.cx);
      const padWidth = Math.ceil(width / cell) * cell;
      const padHeight = Math.ceil(drawHeight / cell) * cell;

      const sheet = document_.slides[index];
      if (sheet === undefined) throw new Error(`no slide ${String(index)}`);
      const markup = api.rsvg.renderSlide(sheet, size, {
        width,
        height: drawHeight,
        idPrefix: `s${String(index)}`,
        text: { defaultTextStyle: document_.defaultTextStyle },
      });

      const image = new Image();
      // A blob rather than a data URI: `encodeURIComponent` throws on a lone
      // surrogate, which is a property of the escape and not of the markup.
      const blob = new Blob([markup], { type: 'image/svg+xml' });
      const href = URL.createObjectURL(blob);
      image.src = href;
      await image.decode();
      const natural = { width: image.naturalWidth, height: image.naturalHeight };

      const canvas = document.createElement('canvas');
      canvas.width = padWidth;
      canvas.height = padHeight;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      if (context === null) throw new Error('no 2d context');
      context.fillStyle = '#FFFFFF';
      context.fillRect(0, 0, padWidth, padHeight);
      context.drawImage(image, 0, 0, width, drawHeight);
      URL.revokeObjectURL(href);

      const pixels = context.getImageData(0, 0, padWidth, padHeight).data;
      const hex = async (buffer: ArrayBufferView | ArrayBuffer): Promise<string> =>
        [...new Uint8Array(await crypto.subtle.digest('SHA-256', buffer as ArrayBuffer))]
          .map((b) => b.toString(16).padStart(2, '0'))
          .join('');

      // The first family of each stack, which is the face the run asked for;
      // what it falls back to is the environment probe's question, not this one.
      const families = new Set<string>();
      for (const match of markup.matchAll(/font-family="([^"]*)"/g)) {
        const first = (match[1] ?? '').split(',')[0] ?? '';
        const name = first
          .trim()
          .replace(/^&quot;|&quot;$/g, '')
          .replace(/^'|'$/g, '');
        if (name.length > 0) families.add(name);
      }

      return {
        grid: reduce({ data: pixels, width: padWidth, height: padHeight, cell }),
        rasterSha256: await hex(pixels),
        svgSha256: await hex(new TextEncoder().encode(markup)),
        families: [...families],
        natural,
        drawHeight,
      };
    },
    { url: deckUrl, index: slideIndex, cell: CELL, width: RASTER_WIDTH },
  );

  const geometry = {
    drawWidth: RASTER_WIDTH,
    drawHeight: result.drawHeight,
    padWidth: Math.ceil(RASTER_WIDTH / CELL) * CELL,
    padHeight: Math.ceil(result.drawHeight / CELL) * CELL,
  };
  if (
    result.natural.width !== geometry.drawWidth ||
    result.natural.height !== geometry.drawHeight
  ) {
    throw new FidelityError(
      'FID_SVG_INTRINSIC_SIZE',
      `the SVG is intrinsically ${String(result.natural.width)}x${String(result.natural.height)}, ` +
        `not the ${String(geometry.drawWidth)}x${String(geometry.drawHeight)} it was asked for`,
      deckUrl,
    );
  }

  return {
    grid: result.grid as Grid,
    rasterSha256: result.rasterSha256,
    svgSha256: result.svgSha256,
    families: result.families,
    geometry,
  };
}

/** PowerPoint's PNG, decoded and reduced by the same code path as our SVG. */
export async function oracleGrid(
  page: Page,
  pngUrl: string,
  geometry: Geometry,
): Promise<{ grid: Grid; rasterSha256: string }> {
  const result = await page.evaluate(
    async ({ url, cell, geo }) => {
      const reduce = (globalThis as unknown as { reduce: (input: unknown) => unknown }).reduce;
      const image = new Image();
      image.src = url;
      await image.decode();
      const canvas = document.createElement('canvas');
      canvas.width = geo.padWidth;
      canvas.height = geo.padHeight;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      if (context === null) throw new Error('no 2d context');
      context.fillStyle = '#FFFFFF';
      context.fillRect(0, 0, geo.padWidth, geo.padHeight);
      context.drawImage(image, 0, 0, geo.drawWidth, geo.drawHeight);
      const pixels = context.getImageData(0, 0, geo.padWidth, geo.padHeight).data;
      const digest = await crypto.subtle.digest('SHA-256', new Uint8Array(pixels));
      return {
        grid: reduce({ data: pixels, width: geo.padWidth, height: geo.padHeight, cell }),
        rasterSha256: [...new Uint8Array(digest)]
          .map((b) => b.toString(16).padStart(2, '0'))
          .join(''),
      };
    },
    { url: pngUrl, cell: CELL, geo: geometry },
  );
  return { grid: result.grid as Grid, rasterSha256: result.rasterSha256 };
}

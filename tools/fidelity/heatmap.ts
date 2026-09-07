/**
 * A reduced grid, and a difference field, as pictures.
 *
 * SVG rather than a raster, for two reasons that matter more than the file
 * size. A cell is the unit the score is computed in, so drawing cells as
 * rectangles shows exactly what was compared rather than a resampling of it;
 * and it needs no image encoder in this repository at all. ADR 0038.
 */

import type { Grid } from './metric/reduce.ts';
import { HISTOGRAM_EDGES } from './metric/score.ts';

/**
 * The nine bands, one per `HISTOGRAM_EDGES` bucket, dark where the two agree.
 *
 * Ordered so that brighter and warmer is worse; the report prints the edges
 * beside it, because a ramp with no scale beside it is decoration.
 */
export const RAMP: readonly string[] = [
  '#0d1117',
  '#132a4a',
  '#17408a',
  '#1f5fc0',
  '#1e88c7',
  '#22b07a',
  '#c9c020',
  '#ef7d18',
  '#f5382e',
];

/** Which band a difference falls in, by the same edges the histogram counts. */
export function bandOf(difference: number): number {
  let bucket = 0;
  while (bucket < HISTOGRAM_EDGES.length && difference >= (HISTOGRAM_EDGES[bucket] ?? 0)) {
    bucket += 1;
  }
  return bucket;
}

function hex2(value: number): string {
  const clamped = value < 0 ? 0 : value > 255 ? 255 : Math.round(value);
  return clamped.toString(16).padStart(2, '0');
}

/**
 * BT.601 back to RGB: the inverse of the transform `reduceRgba` applied.
 *
 * The round trip is lossy - the forward pass averaged a cell and decimated the
 * chroma - so this reconstructs what was measured, not what was drawn.
 */
function rgbOf(luma: number, cb: number, cr: number): string {
  const u = cb - 128;
  const v = cr - 128;
  return `#${hex2(luma + 1.402 * v)}${hex2(luma - 0.344136 * u - 0.714136 * v)}${hex2(luma + 1.772 * u)}`;
}

/**
 * One row of cells as `<rect>`s, with adjacent equal colours merged.
 *
 * The merge is what keeps a 120x68 field to a few hundred elements instead of
 * eight thousand, and a background is one rectangle rather than a wall of them.
 */
function rowRects(
  colours: readonly string[],
  y: number,
  cell: number,
  skip: string | null,
): string {
  const out: string[] = [];
  let start = 0;
  for (let x = 1; x <= colours.length; x++) {
    if (x < colours.length && colours[x] === colours[start]) continue;
    const colour = colours[start] ?? '#000000';
    if (colour !== skip) {
      out.push(
        `<rect x="${String(start * cell)}" y="${String(y * cell)}" ` +
          `width="${String((x - start) * cell)}" height="${String(cell)}" fill="${colour}"/>`,
      );
    }
    start = x;
  }
  return out.join('');
}

function svgOf(width: number, height: number, body: string): string {
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${String(width)} ${String(height)}" ` +
    `preserveAspectRatio="none" shape-rendering="crispEdges">${body}</svg>`
  );
}

/** A grid as the picture it was reduced from, at one rectangle per cell run. */
export function gridSvg(grid: Grid): string {
  const rows: string[] = [];
  for (let y = 0; y < grid.height; y++) {
    const colours: string[] = new Array<string>(grid.width);
    for (let x = 0; x < grid.width; x++) {
      const chroma = (y >> 1) * grid.chromaWidth + (x >> 1);
      colours[x] = rgbOf(
        grid.luma[y * grid.width + x] ?? 0,
        grid.cb[chroma] ?? 128,
        grid.cr[chroma] ?? 128,
      );
    }
    rows.push(rowRects(colours, y, grid.cell, null));
  }
  return svgOf(grid.width * grid.cell, grid.height * grid.cell, rows.join(''));
}

/**
 * A difference field as a heatmap, transparent where the two sides agree.
 *
 * Transparent rather than black, so the map can be laid over either side and
 * read as "here, and this much" instead of replacing the slide it describes.
 */
export function heatmapSvg(
  difference: readonly number[],
  width: number,
  height: number,
  cell: number,
): string {
  const rows: string[] = [];
  for (let y = 0; y < height; y++) {
    const colours: string[] = new Array<string>(width);
    for (let x = 0; x < width; x++) {
      colours[x] = RAMP[bandOf(difference[y * width + x] ?? 0)] ?? RAMP[0] ?? '#000000';
    }
    rows.push(rowRects(colours, y, cell, RAMP[0] ?? null));
  }
  return svgOf(width * cell, height * cell, rows.join(''));
}

/**
 * Pattern fills: the tile, its size, and its two colours.
 *
 * The tiles themselves are in `pattern-tiles.ts`, generated from the C3 fixture.
 * This file is the small amount of arithmetic around them, and the two things
 * about a pattern fill that a renderer gets wrong if it guesses.
 */

import { PaintError } from '../errors.js';
import { PATTERN_PIXEL_EMU, PATTERN_TILES, isPresetPatternName } from './pattern-tiles.js';
import { resolveColor, type ColorContext } from '../colors/resolve.js';
import type { PatternFill } from './fill.js';
import type { Rgba } from '../types.js';

/** Black, the colour PowerPoint paints when `a:fgClr` is missing. */
const DEFAULT_FG: Rgba = { r: 0, g: 0, b: 0, a: 1 };
/** White, the colour PowerPoint paints when `a:bgClr` is missing. */
const DEFAULT_BG: Rgba = { r: 1, g: 1, b: 1, a: 1 };

export interface ResolvedPattern {
  /** Width of the repeat, in EMU. */
  readonly width: number;
  /** Height of the repeat, in EMU. */
  readonly height: number;
  /** Coverage per pixel, row-major, 0 to 1. `length` is `cols * rows`. */
  readonly coverage: readonly number[];
  readonly cols: number;
  readonly rows: number;
  readonly fg: Rgba;
  readonly bg: Rgba;
}

/**
 * A pattern fill, ready to draw.
 *
 * ## The tile is a physical size
 *
 * Eight pixels at 96 DPI - one twelfth of an inch, six points, 76200 EMU - and
 * that is neither a fraction of the shape nor a fixed number of device pixels.
 * Exporting the same shape at 640, 1280, 1920 and 2560 pixels across gives
 * periods of 4, 8, 12 and 16, so the tile scales with the rendering resolution
 * and not with the shape.
 *
 * For SVG that means `patternUnits="userSpaceOnUse"` with `width` and `height`
 * in slide units. `objectBoundingBox` is the trap: it makes the hatching of a
 * wide shape coarser than the same hatching on a narrow one, and slants every
 * diagonal.
 *
 * ## The two colours both default
 *
 * `a:fgClr` and `a:bgClr` are both optional and PowerPoint accepts a package
 * with either missing. Measured: a missing foreground paints black and a missing
 * background paints white, regardless of the theme.
 */
export function resolvePattern(fill: PatternFill, ctx: ColorContext = {}): ResolvedPattern {
  if (!isPresetPatternName(fill.prst)) {
    throw new PaintError(
      'FILL_PATTERN_UNKNOWN',
      `pattFill prst="${fill.prst}" is not one of the 54 preset pattern names`,
      fill.prst,
    );
  }
  const tile = PATTERN_TILES[fill.prst]!;
  const coverage: number[] = [];
  for (let i = 0; i + 2 <= tile.cover.length; i += 2) {
    coverage.push(parseInt(tile.cover.slice(i, i + 2), 16) / 255);
  }
  return {
    width: tile.w * PATTERN_PIXEL_EMU,
    height: tile.h * PATTERN_PIXEL_EMU,
    coverage,
    cols: tile.w,
    rows: tile.h,
    fg: fill.fg === null ? DEFAULT_FG : resolveColor(fill.fg, ctx),
    bg: fill.bg === null ? DEFAULT_BG : resolveColor(fill.bg, ctx),
  };
}

/**
 * The colour of one tile pixel.
 *
 * Coverage is composited in sRGB, which is what 2.6 measured PowerPoint doing
 * for alpha and what the three antialiased tiles - `dnDiag`, `upDiag` and
 * `diagCross` - are drawn with.
 */
export function patternPixel(pattern: ResolvedPattern, col: number, row: number): Rgba {
  const c = ((col % pattern.cols) + pattern.cols) % pattern.cols;
  const r = ((row % pattern.rows) + pattern.rows) % pattern.rows;
  const a = pattern.coverage[r * pattern.cols + c] ?? 0;
  const mix = (x: number, y: number): number => x + (y - x) * a;
  return {
    r: mix(pattern.bg.r, pattern.fg.r),
    g: mix(pattern.bg.g, pattern.fg.g),
    b: mix(pattern.bg.b, pattern.fg.b),
    a: mix(pattern.bg.a, pattern.fg.a),
  };
}

/**
 * What fraction of the tile is foreground.
 *
 * Useful for a renderer choosing whether a pattern is worth drawing at a small
 * zoom, and for the record: the `pctNN` names are labels rather than
 * measurements. `pct75` covers 88 per cent of its tile and `pct20` covers 13.
 */
export function patternInk(name: string): number {
  if (!isPresetPatternName(name)) {
    throw new PaintError(
      'FILL_PATTERN_UNKNOWN',
      `pattFill prst="${name}" is not one of the 54 preset pattern names`,
      name,
    );
  }
  const tile = PATTERN_TILES[name]!;
  let sum = 0;
  let n = 0;
  for (let i = 0; i + 2 <= tile.cover.length; i += 2) {
    sum += parseInt(tile.cover.slice(i, i + 2), 16) / 255;
    n++;
  }
  return n === 0 ? 0 : sum / n;
}

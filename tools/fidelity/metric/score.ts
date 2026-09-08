/**
 * Two grids into the numbers the report is written from.
 *
 * The score is **reported, never gated**: our renderer and PowerPoint differ by
 * a few levels everywhere no matter what we do, so a pass/fail on this number
 * would need a tolerance, and a tolerance is the thing that gets loosened. What
 * is gated is the digest of our own raster, which needs no tolerance at all.
 * ADR 0035.
 */

import { FidelityError } from '../errors.ts';

import type { Grid } from './reduce.ts';

/** The nine dyadic buckets `hist` counts cells into, by their difference. */
export const HISTOGRAM_EDGES: readonly number[] = [1, 2, 4, 8, 16, 32, 64, 128];

/** A run of adjacent cells that differ, and how much they differ by. */
export interface Region {
  /** Cells in the component. */
  readonly cells: number;
  readonly sumD: number;
  readonly maxD: number;
  /** Inclusive cell bounds: left, top, right, bottom. */
  readonly bbox: readonly [number, number, number, number];
}

export interface SlideScore {
  /** Luma cells compared, which is the denominator of `meanBp`. */
  readonly cells: number;
  readonly sumD: number;
  /**
   * Agreement in basis points: 10000 is identical, 0 is maximally different.
   *
   * `10000 - round(10000 * sumD / (255 * cells))`, rounded half up in integers
   * so that two machines cannot round it differently.
   */
  readonly meanBp: number;
  readonly maxD: number;
  readonly hist: readonly number[];
}

/**
 * Agreement over `cells` cells summing to `sumD`, in basis points.
 *
 * Integer arithmetic rounded half up, so two machines cannot round it
 * differently, and the one place the ratio is spelled out.
 */
export function agreementBp(sumD: number, cells: number): number {
  if (cells === 0) return 10000;
  return 10000 - Math.floor((20000 * sumD + 255 * cells) / (510 * cells));
}

/** The per-cell difference of two grids, as luma cells in row-major order. */
export function differenceOf(ours: Grid, theirs: Grid): readonly number[] {
  if (
    ours.width !== theirs.width ||
    ours.height !== theirs.height ||
    ours.chromaWidth !== theirs.chromaWidth ||
    ours.chromaHeight !== theirs.chromaHeight
  ) {
    throw new FidelityError(
      'FID_GRID_MISMATCH',
      `grids are ${String(ours.width)}x${String(ours.height)} and ` +
        `${String(theirs.width)}x${String(theirs.height)}`,
    );
  }

  const out: number[] = new Array<number>(ours.width * ours.height);
  for (let y = 0; y < ours.height; y++) {
    for (let x = 0; x < ours.width; x++) {
      const at = y * ours.width + x;
      const chroma = (y >> 1) * ours.chromaWidth + (x >> 1);
      const dy = Math.abs((ours.luma[at] ?? 0) - (theirs.luma[at] ?? 0));
      const db = Math.abs((ours.cb[chroma] ?? 0) - (theirs.cb[chroma] ?? 0));
      const dr = Math.abs((ours.cr[chroma] ?? 0) - (theirs.cr[chroma] ?? 0));
      // The largest of the three, not their sum or their mean: an error in one
      // channel is an error, and averaging it with two agreeing channels is how
      // a wrong hue at equal luma becomes invisible.
      out[at] = Math.max(dy, db, dr);
    }
  }
  return out;
}

export function scoreOf(difference: readonly number[]): SlideScore {
  const cells = difference.length;
  if (cells === 0) throw new FidelityError('FID_GRID_MISMATCH', 'an empty grid has no score');

  let sumD = 0;
  let maxD = 0;
  const hist: number[] = new Array<number>(HISTOGRAM_EDGES.length + 1).fill(0);
  for (const d of difference) {
    sumD += d;
    if (d > maxD) maxD = d;
    let bucket = 0;
    while (bucket < HISTOGRAM_EDGES.length && d >= (HISTOGRAM_EDGES[bucket] ?? 0)) bucket += 1;
    hist[bucket] = (hist[bucket] ?? 0) + 1;
  }

  return { cells, sumD, meanBp: agreementBp(sumD, cells), maxD, hist };
}

/**
 * The connected runs of cells at or above `floor`, worst first.
 *
 * Ranked rather than thresholded, so nothing is invisible: a defect that is
 * spread thin still has a component, it just sorts below a concentrated one.
 * Four-connected, because a diagonal touch is two defects that happen to meet
 * at a corner and calling them one hides the smaller.
 */
export function regionsOf(
  difference: readonly number[],
  width: number,
  height: number,
  floor: number,
): readonly Region[] {
  if (width * height !== difference.length) {
    throw new FidelityError(
      'FID_GRID_MISMATCH',
      `${String(difference.length)} cells is not ${String(width)}x${String(height)}`,
    );
  }

  const seen = new Uint8Array(difference.length);
  const found: Region[] = [];
  const stack: number[] = [];

  for (let start = 0; start < difference.length; start++) {
    if (seen[start] === 1 || (difference[start] ?? 0) < floor) continue;
    stack.push(start);
    seen[start] = 1;
    let cells = 0;
    let sumD = 0;
    let maxD = 0;
    let left = width;
    let top = height;
    let right = 0;
    let bottom = 0;

    while (stack.length > 0) {
      const at = stack.pop() ?? 0;
      const d = difference[at] ?? 0;
      const x = at % width;
      const y = (at - x) / width;
      cells += 1;
      sumD += d;
      if (d > maxD) maxD = d;
      if (x < left) left = x;
      if (x > right) right = x;
      if (y < top) top = y;
      if (y > bottom) bottom = y;

      const neighbours = [
        x > 0 ? at - 1 : -1,
        x + 1 < width ? at + 1 : -1,
        y > 0 ? at - width : -1,
        y + 1 < height ? at + width : -1,
      ];
      for (const next of neighbours) {
        if (next < 0 || seen[next] === 1 || (difference[next] ?? 0) < floor) continue;
        seen[next] = 1;
        stack.push(next);
      }
    }
    found.push({ cells, sumD, maxD, bbox: [left, top, right, bottom] });
  }

  // A total order, so the report does not churn when the pixels did not.
  return found.sort((a, b) => b.sumD - a.sumD || a.bbox[1] - b.bbox[1] || a.bbox[0] - b.bbox[0]);
}

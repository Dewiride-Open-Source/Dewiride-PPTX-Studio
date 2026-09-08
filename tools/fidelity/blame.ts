/**
 * Whose cell it is.
 *
 * `regionsOf` says where on a slide two rasters differ, in cells. This says
 * which shape was under it, so a report names `blipFill tile flip` rather than
 * `[88,12,115,45]`. ADR 0035.
 */

import { FidelityError } from './errors.ts';
import type { Region } from './metric/score.ts';

/** A placed shape as the renderer laid it out, in raster pixels. */
export interface PlacedBox {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  /** The **unrotated** extent, which is what `Frame` carries. */
  readonly cx: number;
  readonly cy: number;
  /** Degrees, clockwise. */
  readonly rot: number;
}

/**
 * The axis-aligned box a rotated rectangle occupies.
 *
 * Half-extents |cos|cx + |sin|cy and |sin|cx + |cos|cy about the same centre. A
 * flip maps the rectangle onto itself, so it does not enter.
 */
export function boxOf(placed: PlacedBox): ShapeBox {
  const radians = (placed.rot * Math.PI) / 180;
  const spread = Math.abs(Math.cos(radians)) * placed.cx + Math.abs(Math.sin(radians)) * placed.cy;
  const rise = Math.abs(Math.sin(radians)) * placed.cx + Math.abs(Math.cos(radians)) * placed.cy;
  return {
    id: placed.id,
    name: placed.name,
    left: placed.x + (placed.cx - spread) / 2,
    top: placed.y + (placed.cy - rise) / 2,
    right: placed.x + (placed.cx + spread) / 2,
    bottom: placed.y + (placed.cy + rise) / 2,
  };
}

/** A placed shape's axis-aligned bounding box, in raster pixels. */
export interface ShapeBox {
  readonly id: number;
  readonly name: string;
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** What one shape is answerable for, over the cells it painted last. */
export interface Blame {
  /** `null` for the cells no shape covers. */
  readonly id: number | null;
  readonly name: string;
  /** Cells that differ at all. */
  readonly cells: number;
  /** Cells the shape painted last, differing or not, and the denominator. */
  readonly covered: number;
  readonly sumD: number;
  readonly maxD: number;
}

/** What a shape whose `p:cNvPr` carries no name is reported as. */
export const UNNAMED = '(unnamed)';

/** What a differing cell under no shape at all is reported as. */
export const NO_SHAPE = '(no shape)';

/** The inclusive cell bounds a box touches; `x1 < x0` where it touches none. */
function cellsOf(box: ShapeBox, cell: number): { x0: number; y0: number; x1: number; y1: number } {
  return {
    x0: Math.floor(box.left / cell),
    y0: Math.floor(box.top / cell),
    x1: Math.ceil(box.right / cell) - 1,
    y1: Math.ceil(box.bottom / cell) - 1,
  };
}

/**
 * Which shape painted each cell last, as an index into `shapes`, or -1.
 *
 * Painter's order rather than smallest-box: the shape a viewer sees at a cell
 * is the last one drawn over it, and that is the one to go and look at.
 */
function ownersOf(
  shapes: readonly ShapeBox[],
  width: number,
  height: number,
  cell: number,
): Int32Array {
  const owner = new Int32Array(width * height).fill(-1);
  shapes.forEach((box, index) => {
    const span = cellsOf(box, cell);
    const top = Math.max(0, span.y0);
    const bottom = Math.min(height - 1, span.y1);
    const left = Math.max(0, span.x0);
    const right = Math.min(width - 1, span.x1);
    for (let y = top; y <= bottom; y++) {
      for (let x = left; x <= right; x++) owner[y * width + x] = index;
    }
  });
  return owner;
}

/** Ties break on the name, so two runs of the same tree sort identically. */
function byName(a: string, b: string): number {
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

function nameOf(shapes: readonly ShapeBox[], index: number): string {
  if (index < 0) return NO_SHAPE;
  const name = shapes[index]?.name ?? '';
  return name === '' ? UNNAMED : name;
}

/**
 * The difference, split over the shapes that painted it, worst first.
 *
 * There is no floor and no threshold: every differing cell is attributed to
 * exactly one shape, so the totals add up to the slide's own `sumD` and nothing
 * is filtered out of the answer before a reader sees it.
 */
export function blameOf(
  difference: readonly number[],
  width: number,
  height: number,
  cell: number,
  shapes: readonly ShapeBox[],
): readonly Blame[] {
  if (width * height !== difference.length) {
    throw new FidelityError(
      'FID_GRID_MISMATCH',
      `${String(difference.length)} cells is not ${String(width)}x${String(height)}`,
    );
  }

  const owner = ownersOf(shapes, width, height, cell);
  const tally = new Map<number, { cells: number; covered: number; sumD: number; maxD: number }>();
  for (let at = 0; at < difference.length; at++) {
    const index = owner[at] ?? -1;
    const row = tally.get(index) ?? { cells: 0, covered: 0, sumD: 0, maxD: 0 };
    row.covered += 1;
    const d = difference[at] ?? 0;
    if (d > 0) {
      row.cells += 1;
      row.sumD += d;
      if (d > row.maxD) row.maxD = d;
    }
    tally.set(index, row);
  }

  const out: Blame[] = [];
  for (const [index, row] of tally) {
    if (row.sumD === 0) continue;
    out.push({
      id: index < 0 ? null : (shapes[index]?.id ?? null),
      name: nameOf(shapes, index),
      ...row,
    });
  }
  // A total order, so the report does not churn when the pixels did not.
  return out.sort((a, b) => b.sumD - a.sumD || b.maxD - a.maxD || byName(a.name, b.name));
}

/**
 * The shapes under a region's bounding box, by how much of it they cover.
 *
 * A `Region` carries its bounds and not its cells, so this is the box and not
 * the component: a diagonal region names everything its box crosses.
 */
export function regionShapes(
  region: Region,
  cell: number,
  shapes: readonly ShapeBox[],
  limit: number,
): readonly string[] {
  const [left, top, right, bottom] = region.bbox;
  const scored: { name: string; area: number }[] = [];
  for (const box of shapes) {
    const span = cellsOf(box, cell);
    const overlapX = Math.min(right, span.x1) - Math.max(left, span.x0) + 1;
    const overlapY = Math.min(bottom, span.y1) - Math.max(top, span.y0) + 1;
    if (overlapX <= 0 || overlapY <= 0) continue;
    scored.push({ name: box.name === '' ? UNNAMED : box.name, area: overlapX * overlapY });
  }
  scored.sort((a, b) => b.area - a.area || byName(a.name, b.name));

  const seen = new Set<string>();
  const out: string[] = [];
  for (const entry of scored) {
    if (seen.has(entry.name)) continue;
    seen.add(entry.name);
    out.push(entry.name);
    if (out.length === limit) break;
  }
  return out;
}

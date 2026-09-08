/**
 * Which shape a differing cell is charged to.
 *
 * The report's whole claim is that the totals add up and nothing is quietly
 * dropped, so the invariant is asserted directly rather than inferred from a
 * ranking. ADR 0035.
 */

import { describe, expect, it } from 'vitest';

import { NO_SHAPE, UNNAMED, blameOf, boxOf, regionShapes, type ShapeBox } from './blame.ts';
import { isFidelityError } from './errors.ts';
import { regionsOf } from './metric/score.ts';

const CELL = 8;

function box(name: string, left: number, top: number, right: number, bottom: number): ShapeBox {
  return { id: left * 1000 + top, name, left, top, right, bottom };
}

/** A 4x3 difference field, written as rows so a case reads like the picture. */
function field(rows: readonly (readonly number[])[]): readonly number[] {
  return rows.flat();
}

const FOUR_BY_THREE = { width: 4, height: 3 };

function blame(rows: readonly (readonly number[])[], shapes: readonly ShapeBox[]) {
  return blameOf(field(rows), FOUR_BY_THREE.width, FOUR_BY_THREE.height, CELL, shapes);
}

describe('blaming a shape for a cell', () => {
  it('charges every differing cell to exactly one shape, and the totals add up', () => {
    const rows = [
      [10, 20, 0, 5],
      [0, 30, 4, 0],
      [1, 0, 0, 2],
    ];
    const shapes = [box('left half', 0, 0, 16, 24), box('right half', 16, 0, 32, 24)];
    const blamed = blame(rows, shapes);
    const sumD = rows.flat().reduce((a, b) => a + b, 0);
    expect(blamed.reduce((total, row) => total + row.sumD, 0)).toBe(sumD);
    expect(blamed.reduce((total, row) => total + row.cells, 0)).toBe(
      rows.flat().filter((d) => d > 0).length,
    );
  });

  it('gives a cell to the shape that painted it last, not the largest', () => {
    const rows = [
      [0, 0, 0, 0],
      [0, 9, 0, 0],
      [0, 0, 0, 0],
    ];
    const under = box('underneath', 0, 0, 32, 24);
    const over = box('on top', 8, 8, 16, 16);
    expect(blame(rows, [under, over])[0]?.name).toBe('on top');
    // The same two shapes the other way round is a different answer, and that
    // is the point: paint order decides, not size.
    expect(blame(rows, [over, under])[0]?.name).toBe('underneath');
  });

  it('charges a cell no shape covers to nobody, by name', () => {
    const rows = [
      [7, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    const blamed = blame(rows, [box('elsewhere', 16, 16, 32, 24)]);
    expect(blamed).toHaveLength(1);
    expect(blamed[0]?.name).toBe(NO_SHAPE);
    expect(blamed[0]?.id).toBeNull();
  });

  it('counts covered cells whether they differ or not, and differing cells only once', () => {
    const rows = [
      [3, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    // Two cells wide, one tall: covered is 2 and only one of them differs.
    const blamed = blame(rows, [box('a pair', 0, 0, 16, 8)]);
    expect(blamed[0]?.covered).toBe(2);
    expect(blamed[0]?.cells).toBe(1);
  });

  it('leaves out a shape that differs nowhere', () => {
    const rows = [
      [0, 0, 0, 6],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    const blamed = blame(rows, [box('clean', 0, 0, 8, 8), box('dirty', 24, 0, 32, 8)]);
    expect(blamed.map((row) => row.name)).toEqual(['dirty']);
  });

  it('takes a partly covered cell, because a shape paints into it', () => {
    const rows = [
      [4, 4, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    // 4 to 12 pixels crosses the boundary at 8, so it owns cell 0 and cell 1.
    expect(blame(rows, [box('astride', 4, 0, 12, 8)])[0]?.covered).toBe(2);
  });

  it('clips a shape that hangs off the raster instead of writing past it', () => {
    const rows = [
      [0, 0, 0, 5],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    const blamed = blame(rows, [box('overhanging', -40, -40, 400, 400)]);
    expect(blamed[0]?.name).toBe('overhanging');
    expect(blamed[0]?.covered).toBe(12);
  });

  it('does not let a shape off the right edge wrap onto the next row', () => {
    const rows = [
      [0, 0, 0, 0],
      [9, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    // Row 0 only, but 48 pixels wide against a 32-pixel raster. Unclamped, the
    // overhang runs into row 1 and steals the cell that actually differs.
    const blamed = blame(rows, [box('too wide', 0, 0, 48, 8)]);
    expect(blamed[0]?.name).toBe(NO_SHAPE);
  });

  it('names a shape whose p:cNvPr carried none', () => {
    const rows = [
      [8, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    expect(blame(rows, [box('', 0, 0, 8, 8)])[0]?.name).toBe(UNNAMED);
  });

  it('ranks by total difference, then by the worst cell, then by name', () => {
    const rows = [
      [10, 1, 1, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ];
    const blamed = blame(rows, [
      box('one big', 0, 0, 8, 8),
      box('b two small', 8, 0, 16, 8),
      box('a two small', 16, 0, 24, 8),
    ]);
    expect(blamed.map((row) => row.name)).toEqual(['one big', 'a two small', 'b two small']);
  });

  it('refuses a difference field that is not the grid it was given', () => {
    try {
      blameOf([1, 2, 3], 4, 3, CELL, []);
      expect.unreachable('a 3-cell field is not 4x3');
    } catch (error) {
      expect(isFidelityError(error) && error.code).toBe('FID_GRID_MISMATCH');
    }
  });
});

/** Equal to within float noise: sin(pi) is not zero and never will be. */
function sameBox(got: ShapeBox, want: ShapeBox): void {
  expect(got.left).toBeCloseTo(want.left, 9);
  expect(got.top).toBeCloseTo(want.top, 9);
  expect(got.right).toBeCloseTo(want.right, 9);
  expect(got.bottom).toBeCloseTo(want.bottom, 9);
}

describe('the box a placed shape occupies', () => {
  const square = { id: 1, name: 'square', x: 10, y: 20, cx: 40, cy: 40, rot: 0 };
  const wide = { ...square, cx: 40, cy: 10 };

  it('is the rectangle itself when nothing is turned', () => {
    expect(boxOf(square)).toEqual({
      id: 1,
      name: 'square',
      left: 10,
      top: 20,
      right: 50,
      bottom: 60,
    });
  });

  it('swaps the extents at a quarter turn, about the same centre', () => {
    const turned = boxOf({ ...wide, rot: 90 });
    expect(turned.right - turned.left).toBeCloseTo(10, 9);
    expect(turned.bottom - turned.top).toBeCloseTo(40, 9);
    expect((turned.left + turned.right) / 2).toBeCloseTo(30, 9);
    expect((turned.top + turned.bottom) / 2).toBeCloseTo(25, 9);
  });

  it('grows a square by root two on the diagonal', () => {
    const turned = boxOf({ ...square, rot: 45 });
    expect(turned.right - turned.left).toBeCloseTo(40 * Math.SQRT2, 9);
    expect(turned.bottom - turned.top).toBeCloseTo(40 * Math.SQRT2, 9);
  });

  it('gives a half turn the same box as no turn at all', () => {
    sameBox(boxOf({ ...wide, rot: 180 }), boxOf(wide));
  });

  it('turns clockwise or anticlockwise to the same box, because it is a box', () => {
    sameBox(boxOf({ ...wide, rot: 30 }), boxOf({ ...wide, rot: 330 }));
  });
});

describe('naming a region', () => {
  const difference = field([
    [0, 0, 0, 0],
    [0, 40, 40, 0],
    [0, 0, 0, 0],
  ]);

  it('names the shapes its bounding box crosses, most covered first', () => {
    const region = regionsOf(difference, 4, 3, 16)[0];
    expect(region).toBeDefined();
    if (region === undefined) return;
    const names = regionShapes(
      region,
      CELL,
      [box('narrow', 8, 8, 16, 16), box('wide', 8, 0, 24, 24)],
      2,
    );
    expect(names).toEqual(['wide', 'narrow']);
  });

  it('stops at the limit and never repeats a name', () => {
    const region = regionsOf(difference, 4, 3, 16)[0];
    expect(region).toBeDefined();
    if (region === undefined) return;
    const twice = [
      box('same', 8, 8, 16, 16),
      box('same', 8, 8, 24, 16),
      box('other', 8, 8, 16, 16),
    ];
    expect(regionShapes(region, CELL, twice, 1)).toEqual(['same']);
    expect(regionShapes(region, CELL, twice, 5)).toEqual(['same', 'other']);
  });

  it('names nothing where nothing overlaps', () => {
    const region = regionsOf(difference, 4, 3, 16)[0];
    expect(region).toBeDefined();
    if (region === undefined) return;
    expect(regionShapes(region, CELL, [box('far away', 200, 200, 240, 240)], 3)).toEqual([]);
  });
});

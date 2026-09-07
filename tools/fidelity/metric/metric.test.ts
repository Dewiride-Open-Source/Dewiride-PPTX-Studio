/**
 * The metric, against arithmetic done by hand.
 *
 * Every expected number below is derived in ADR 0035 from the BT.601
 * coefficients and the rounding rule, and typed in from there. None of it was
 * read off a run: a fixture recorded from the implementation would agree with
 * any implementation, including a wrong one.
 */

import { describe, expect, it } from 'vitest';

import { FidelityError } from '../errors.ts';

import { decodeGrid, encodeGrid } from './grid.ts';
import { reduceRgba, type Grid } from './reduce.ts';
import { differenceOf, regionsOf, scoreOf } from './score.ts';

const CELL = 8;
/** Four cells square, so the arithmetic has a denominator anyone can divide by. */
const SIDE = 4 * CELL;

type Rgb = readonly [number, number, number];

function raster(fill: Rgb, width = SIDE, height = SIDE): Uint8ClampedArray {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let at = 0; at < data.length; at += 4) {
    data[at] = fill[0];
    data[at + 1] = fill[1];
    data[at + 2] = fill[2];
    data[at + 3] = 255;
  }
  return data;
}

function paint(
  data: Uint8ClampedArray,
  colour: Rgb,
  box: { x: number; y: number; w: number; h: number },
  width = SIDE,
): void {
  for (let y = box.y; y < box.y + box.h; y++) {
    for (let x = box.x; x < box.x + box.w; x++) {
      const at = (y * width + x) * 4;
      data[at] = colour[0];
      data[at + 1] = colour[1];
      data[at + 2] = colour[2];
      data[at + 3] = 255;
    }
  }
}

function gridOf(data: Uint8ClampedArray, width = SIDE, height = SIDE): Grid {
  return reduceRgba({ data, width, height, cell: CELL });
}

const WHITE: Rgb = [255, 255, 255];
const BLACK: Rgb = [0, 0, 0];
/** Chosen so that its luma is 77, which is also pure red's. */
const GREEN: Rgb = [0, 131, 0];
const RED: Rgb = [255, 0, 0];

describe('the reduction', () => {
  it('gives white and black the luma the coefficients say, and neither any chroma', () => {
    const white = gridOf(raster(WHITE));
    const black = gridOf(raster(BLACK));
    expect(white.luma[0]).toBe(255);
    expect(black.luma[0]).toBe(0);
    // A grey has no colour, so both sit at the 128 that means "no chroma".
    expect([white.cb[0], white.cr[0]]).toEqual([128, 128]);
    expect([black.cb[0], black.cr[0]]).toEqual([128, 128]);
  });

  it('gives that green and pure red the same luma and the same Cb', () => {
    const green = gridOf(raster(GREEN));
    const red = gridOf(raster(RED));
    expect(green.luma[0]).toBe(77);
    expect(red.luma[0]).toBe(77);
    expect(green.cb[0]).toBe(85);
    expect(red.cb[0]).toBe(85);
    // Only Cr separates them, and red's is clamped rather than allowed to wrap.
    expect(green.cr[0]).toBe(73);
    expect(red.cr[0]).toBe(255);
  });

  it('rounds a cell mean half up, so one pixel at 32 of 64 is one level', () => {
    const data = raster(BLACK);
    paint(data, [32, 32, 32], { x: 0, y: 0, w: 1, h: 1 });
    expect(gridOf(data).luma[0]).toBe(1);
  });

  it('halves the chroma resolution in each axis and leaves luma alone', () => {
    const grid = gridOf(raster(WHITE));
    expect([grid.width, grid.height]).toEqual([4, 4]);
    expect([grid.chromaWidth, grid.chromaHeight]).toEqual([2, 2]);
    expect(grid.luma.length).toBe(16);
    expect(grid.cb.length).toBe(4);
  });

  it('refuses a raster that is not a whole number of cells', () => {
    expect(() =>
      reduceRgba({ data: raster(WHITE, 30, 32), width: 30, height: 32, cell: CELL }),
    ).toThrow(/not whole cells/);
  });
});

describe('the score', () => {
  it('is 10000 when the two grids are the same', () => {
    const score = scoreOf(differenceOf(gridOf(raster(WHITE)), gridOf(raster(WHITE))));
    expect(score.meanBp).toBe(10000);
    expect(score.maxD).toBe(0);
  });

  it('is 9375 for one black cell of sixteen on white', () => {
    // d = 255 on one cell of 16. 10000 - floor((20000*255 + 4080)/8160)
    //   = 10000 - floor(5104080/8160) = 10000 - 625.
    const ours = raster(WHITE);
    paint(ours, BLACK, { x: 0, y: 0, w: CELL, h: CELL });
    const score = scoreOf(differenceOf(gridOf(ours), gridOf(raster(WHITE))));
    expect(score.sumD).toBe(255);
    expect(score.meanBp).toBe(9375);
    expect(score.maxD).toBe(255);
    expect(score.hist).toEqual([15, 0, 0, 0, 0, 0, 0, 0, 1]);
  });

  it('is 8216 for a hue that differs at equal luma, which a luma metric calls identical', () => {
    // Red over one whole chroma cell of green: dY = 0, dCb = 0, dCr = 255-73 = 182,
    // on 4 luma cells. 10000 - floor((20000*728 + 4080)/8160) = 10000 - 1784.
    const ours = raster(GREEN);
    paint(ours, RED, { x: 0, y: 0, w: 2 * CELL, h: 2 * CELL });
    const difference = differenceOf(gridOf(ours), gridOf(raster(GREEN)));
    const score = scoreOf(difference);
    expect(difference[0]).toBe(182);
    expect(score.sumD).toBe(728);
    expect(score.meanBp).toBe(8216);
    expect(score.hist).toEqual([12, 0, 0, 0, 0, 0, 0, 0, 4]);
  });

  it('moves for a single pixel one level off, because there is no deadband', () => {
    // The whole point of the metric having no tolerance: the smallest change a
    // cell mean can express still moves the number.
    const ours = raster(BLACK);
    paint(ours, [32, 32, 32], { x: 0, y: 0, w: 1, h: 1 });
    const score = scoreOf(differenceOf(gridOf(ours), gridOf(raster(BLACK))));
    expect(score.sumD).toBe(1);
    expect(score.meanBp).toBe(9998);
    expect(score.hist).toEqual([15, 1, 0, 0, 0, 0, 0, 0, 0]);
  });

  it('rounds its own division half up, which two cells one level off are enough to show', () => {
    // sumD = 2, cells = 16.  10000*2/(255*16) = 4.902..., so rounding gives 5
    // and truncating gives 4 - a whole basis point of difference from two pixels.
    const ours = raster(BLACK);
    paint(ours, [32, 32, 32], { x: 0, y: 0, w: 1, h: 1 });
    paint(ours, [32, 32, 32], { x: CELL, y: 0, w: 1, h: 1 });
    const score = scoreOf(differenceOf(gridOf(ours), gridOf(raster(BLACK))));
    expect(score.sumD).toBe(2);
    expect(score.meanBp).toBe(9995);
  });

  it('rounds the chroma it decimates, so a quarter of a level is not thrown away', () => {
    // (0,0,3) has Cb 130 against black's 128 and the same luma. One such cell in
    // a chroma cell of four sums to 514, which is 129 rounded and 128 truncated.
    const ours = raster(BLACK);
    paint(ours, [0, 0, 3], { x: 0, y: 0, w: CELL, h: CELL });
    const grid = gridOf(ours);
    expect(grid.luma[0]).toBe(0);
    expect(grid.cb[0]).toBe(129);
    const score = scoreOf(differenceOf(grid, gridOf(raster(BLACK))));
    expect(score.sumD).toBe(4);
    expect(score.meanBp).toBe(9990);
  });

  it('reaches 0 only when every cell is maximally different', () => {
    const score = scoreOf(differenceOf(gridOf(raster(WHITE)), gridOf(raster(BLACK))));
    expect(score.meanBp).toBe(0);
  });

  it('refuses two grids that are not the same shape', () => {
    const wide = reduceRgba({
      data: raster(WHITE, SIDE * 2, SIDE),
      width: SIDE * 2,
      height: SIDE,
      cell: CELL,
    });
    expect(() => differenceOf(gridOf(raster(WHITE)), wide)).toThrow(FidelityError);
  });
});

describe('the regions a difference falls into', () => {
  /** A 4x4 field of differences, written out so the shape is visible. */
  function field(rows: readonly (readonly number[])[]): readonly number[] {
    return rows.flat();
  }

  it('counts a diagonal as separate runs, because a corner touch is not adjacency', () => {
    const diagonal = field([
      [9, 0, 0, 0],
      [0, 9, 0, 0],
      [0, 0, 9, 0],
      [0, 0, 0, 9],
    ]);
    expect(regionsOf(diagonal, 4, 4, 1).length).toBe(4);
  });

  it('joins cells that share an edge', () => {
    const line = field([
      [9, 9, 9, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ]);
    const regions = regionsOf(line, 4, 4, 1);
    expect(regions.length).toBe(1);
    expect(regions[0]?.cells).toBe(3);
    expect(regions[0]?.bbox).toEqual([0, 0, 2, 0]);
  });

  it('ranks the heaviest run first however the cells are ordered', () => {
    const two = field([
      [3, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 40, 40],
      [0, 0, 0, 0],
    ]);
    const regions = regionsOf(two, 4, 4, 1);
    expect(regions.map((r) => r.sumD)).toEqual([80, 3]);
  });

  it('ranks by how much differs, not by how many cells do', () => {
    const lopsided = field([
      [2, 2, 2, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 50],
      [0, 0, 0, 0],
    ]);
    const regions = regionsOf(lopsided, 4, 4, 1);
    expect(regions.map((r) => r.sumD)).toEqual([50, 6]);
    expect(regions.map((r) => r.cells)).toEqual([1, 3]);
  });

  it('includes a cell sitting exactly on the floor', () => {
    const onTheLine = field([
      [1, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ]);
    expect(regionsOf(onTheLine, 4, 4, 1).length).toBe(1);
  });

  it('sees nothing below the floor it is given', () => {
    const faint = field([
      [3, 3, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
      [0, 0, 0, 0],
    ]);
    expect(regionsOf(faint, 4, 4, 1).length).toBe(1);
    expect(regionsOf(faint, 4, 4, 32).length).toBe(0);
  });
});

describe('the committed grid file', () => {
  it('round-trips every plane', () => {
    const data = raster(GREEN);
    paint(data, RED, { x: 0, y: 0, w: CELL, h: CELL });
    const grid = gridOf(data);
    const back = decodeGrid(encodeGrid(grid), 'test');
    expect(back).toEqual(grid);
  });

  it('is header plus one byte per luma cell and one per chroma cell per plane', () => {
    // 10 + 16 + 4 + 4, which is the arithmetic the storage budget rests on.
    expect(encodeGrid(gridOf(raster(WHITE))).length).toBe(10 + 16 + 4 + 4);
  });

  it('refuses a file that does not begin with the magic', () => {
    const bytes = encodeGrid(gridOf(raster(WHITE)));
    bytes[0] = 0;
    expect(() => decodeGrid(bytes, 'test')).toThrow(FidelityError);
  });

  it('refuses a file whose length does not match its own header', () => {
    const bytes = encodeGrid(gridOf(raster(WHITE)));
    expect(() => decodeGrid(bytes.subarray(0, bytes.length - 1), 'test')).toThrow(/bytes, not the/);
  });
});

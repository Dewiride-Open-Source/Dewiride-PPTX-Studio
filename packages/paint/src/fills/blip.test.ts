/**
 * Every expected value here comes from `corpus/ground-truth/blips.json` or from
 * the arithmetic its rules state - never from reading `blip.ts`. The adversarial
 * cases assert that the refuted readings are wrong, because a suite that only
 * agrees with the implementation passes forever and proves nothing.
 */

import { describe, expect, it } from 'vitest';

import { blipPlacement, brightOffset, contrastScale, BLIP_LUMA } from './blip.js';
import { WHOLE_RECT, type BlipFill, type RelativeRect, type TileAlign } from './fill.js';
import { PaintError } from '../errors.js';

/** The C6 probe image: 32 pixels square at 96 dpi, so one tile is 24 pt. */
const IMAGE = { widthPx: 32, heightPx: 32, dpi: 96 };
const PT = 12700;
const NATURAL = 24 * PT;

/** The C6 geometry deck's shape: 100 x 80 pt at the origin. */
const BOX = { x: 0, y: 0, w: 100 * PT, h: 80 * PT };

function stretchFill(
  fillRect: RelativeRect = WHOLE_RECT,
  srcRect: RelativeRect = WHOLE_RECT,
): BlipFill {
  return {
    type: 'blip',
    embed: 'rId2',
    svgEmbed: null,
    part: '/ppt/slides/slide1.xml',
    srcRect,
    mode: { kind: 'stretch', fillRect },
    effects: [],
    dpi: 0,
    rotWithShape: true,
  };
}

function tileFill(
  over: Partial<{
    tx: number;
    ty: number;
    sx: number;
    sy: number;
    algn: TileAlign;
    flip: 'none' | 'x' | 'y' | 'xy';
  }> = {},
  srcRect: RelativeRect = WHOLE_RECT,
): BlipFill {
  return {
    type: 'blip',
    embed: 'rId2',
    svgEmbed: null,
    part: '/ppt/slides/slide1.xml',
    srcRect,
    mode: {
      kind: 'tile',
      tx: over.tx ?? 0,
      ty: over.ty ?? 0,
      sx: over.sx ?? 100000,
      sy: over.sy ?? 100000,
      flip: over.flip ?? 'none',
      algn: over.algn ?? 'tl',
    },
    effects: [],
    dpi: 0,
    rotWithShape: true,
  };
}

describe('a:stretch', () => {
  it('maps the whole image onto the whole box, aspect ignored', () => {
    const placement = blipPlacement(stretchFill(), BOX, IMAGE);
    expect(placement).toMatchObject({ kind: 'stretch', image: BOX, clip: BOX });
  });

  it('insets the destination by a:fillRect', () => {
    const inset = { l: 10000, t: 10000, r: 10000, b: 10000 };
    const placement = blipPlacement(stretchFill(inset), BOX, IMAGE);
    expect(placement.kind).toBe('stretch');
    if (placement.kind !== 'stretch') return;
    expect(placement.clip).toEqual({ x: 10 * PT, y: 8 * PT, w: 80 * PT, h: 64 * PT });
  });

  it('outsets the destination when a:fillRect is negative', () => {
    const placement = blipPlacement(stretchFill({ l: -20000, t: 0, r: 0, b: 0 }), BOX, IMAGE);
    expect(placement.kind).toBe('stretch');
    if (placement.kind !== 'stretch') return;
    expect(placement.clip.x).toBe(-20 * PT);
    expect(placement.clip.w).toBe(120 * PT);
  });

  it('grows the image so the crop a:srcRect leaves lands on the box', () => {
    // 25% off the left and 12.5% off the right leaves 0.625 of the width, so
    // the whole image is 1/0.625 = 1.6 times the box and starts left of it.
    const placement = blipPlacement(
      stretchFill(WHOLE_RECT, { l: 25000, t: 25000, r: 12500, b: 12500 }),
      BOX,
      IMAGE,
    );
    expect(placement.kind).toBe('stretch');
    if (placement.kind !== 'stretch') return;
    expect(placement.image.w).toBeCloseTo(160 * PT, 6);
    expect(placement.image.x).toBeCloseTo(-40 * PT, 6);
    expect(placement.image.h).toBeCloseTo(128 * PT, 6);
    expect(placement.image.y).toBeCloseTo(-32 * PT, 6);
  });

  it('refuses an a:srcRect that leaves no source', () => {
    expect(() =>
      blipPlacement(stretchFill(WHOLE_RECT, { l: 60000, t: 0, r: 40000, b: 0 }), BOX, IMAGE),
    ).toThrow(PaintError);
  });
});

describe('a:tile', () => {
  it('sizes one tile from the image resolution, not from the shape', () => {
    const placement = blipPlacement(tileFill(), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.tileW).toBe(NATURAL);
    expect(placement.tileH).toBe(NATURAL);
    // The refuted reading: a percentage of the shape would give 100 x 80 pt.
    expect(placement.tileW).not.toBe(BOX.w);
  });

  it('follows the image own resolution when it is not 96 dpi', () => {
    // 32 px at 150 dpi is 15.36 pt, which is what dpi-image150 measured.
    const placement = blipPlacement(tileFill(), BOX, { ...IMAGE, dpi: 150 });
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.tileW).toBeCloseTo(15.36 * PT, 3);
  });

  it('scales each axis by its own percentage of the natural size', () => {
    const placement = blipPlacement(tileFill({ sx: 50000, sy: 150000 }), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.tileW).toBe(12 * PT);
    expect(placement.tileH).toBe(36 * PT);
  });

  it.each([
    ['tl', 0, 0],
    ['t', (100 - 24) / 2, 0],
    ['tr', 100 - 24, 0],
    ['l', 0, (80 - 24) / 2],
    ['ctr', (100 - 24) / 2, (80 - 24) / 2],
    ['r', 100 - 24, (80 - 24) / 2],
    ['bl', 0, 80 - 24],
    ['b', (100 - 24) / 2, 80 - 24],
    ['br', 100 - 24, 80 - 24],
  ] as const)('puts one whole tile flush to the %s edge', (algn, x, y) => {
    const placement = blipPlacement(tileFill({ algn }), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.originX).toBeCloseTo(x * PT, 6);
    expect(placement.originY).toBeCloseTo(y * PT, 6);
  });

  it('moves the lattice from that anchor by @tx and @ty, positive right and down', () => {
    const placement = blipPlacement(tileFill({ algn: 'ctr', tx: 6 * PT, ty: 6 * PT }), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.originX).toBeCloseTo((38 + 6) * PT, 6);
    expect(placement.originY).toBeCloseTo((28 + 6) * PT, 6);
  });

  it('takes a negative offset the other way', () => {
    const placement = blipPlacement(tileFill({ tx: -9 * PT, ty: -6 * PT }), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.originX).toBe(-9 * PT);
    expect(placement.originY).toBe(-6 * PT);
  });

  it.each([
    ['none', false, false],
    ['x', true, false],
    ['y', false, true],
    ['xy', true, true],
  ] as const)('reads @flip="%s" as mirroring on those axes', (flip, flipX, flipY) => {
    const placement = blipPlacement(tileFill({ flip }), BOX, IMAGE);
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.flipX).toBe(flipX);
    expect(placement.flipY).toBe(flipY);
  });

  it('sizes the tile from the whole image even when a:srcRect crops it', () => {
    // `tile-size-follows-crop` scores 0.19 against this reading's 0.86.
    const placement = blipPlacement(
      tileFill({}, { l: 25000, t: 12500, r: 12500, b: 0 }),
      BOX,
      IMAGE,
    );
    expect(placement.kind).toBe('tile');
    if (placement.kind !== 'tile') return;
    expect(placement.tileW).toBe(NATURAL);
    expect(placement.tileH).toBe(NATURAL);
    // The crop still decides what is drawn: the image is bigger than the tile.
    expect(placement.image.w).toBeCloseTo(NATURAL / 0.625, 3);
  });

  it('refuses a tile scaled to nothing', () => {
    expect(() => blipPlacement(tileFill({ sx: 0 }), BOX, IMAGE)).toThrow(PaintError);
  });
});

describe('the measured effect laws', () => {
  it('weighs luminance by BT.709 and not by BT.601', () => {
    // Red at full scale measured 54 of 255; BT.601 would have put it at 76.
    expect(Math.round(BLIP_LUMA.r * 255)).toBe(54);
    expect(Math.round(BLIP_LUMA.g * 255)).toBe(182);
    expect(Math.round(BLIP_LUMA.b * 255)).toBe(18);
  });

  it('divides for positive contrast and adds for negative', () => {
    expect(contrastScale(60000)).toBeCloseTo(2.5, 10);
    expect(contrastScale(30000)).toBeCloseTo(1 / 0.7, 10);
    expect(contrastScale(-60000)).toBeCloseTo(0.4, 10);
    expect(contrastScale(-70000)).toBeCloseTo(0.3, 10);
    expect(contrastScale(0)).toBe(1);
    // The refuted reading: one formula for both signs.
    expect(contrastScale(-60000)).not.toBeCloseTo(1 / 1.6, 6);
  });

  it('grows the brightness term with the contrast slope', () => {
    // Twelve probes put the coefficient at 127.5 x (1 + k), which in 0..1 is
    // 0.5 x (1 + k); treating brightness as independent of k scores zero.
    expect(brightOffset(40000, 1)).toBeCloseTo(0.4, 10);
    expect(brightOffset(70000, 0.3)).toBeCloseTo(0.7 * 0.5 * 1.3, 10);
    expect(brightOffset(70000, 0.3)).not.toBeCloseTo(0.7, 6);
  });
});

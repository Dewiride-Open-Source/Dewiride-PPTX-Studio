/**
 * Where an `a:blipFill` puts its image, as measured against real PowerPoint.
 *
 * Every rule here is scored against `corpus/ground-truth/blips.json`, where the
 * plausible wrong reading of each one is scored beside it. Two are worth naming
 * because they are the ones anybody writes first: `@flip` mirrors ALTERNATE
 * tiles rather than every tile, and `a:srcRect` changes what is drawn inside a
 * tile without changing how big the tile is. ADR 0036.
 */

import { EMU_PER_INCH } from '../types.js';
import { PaintError } from '../errors.js';
import type { BlipFill, RelativeRect, TileAlign } from './fill.js';

/** A rectangle in whatever units its producer named. */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** What the caller knows about the decoded image. */
export interface ImageSize {
  readonly widthPx: number;
  readonly heightPx: number;
  /** The image's own declared resolution, which is what sizes a tile. */
  readonly dpi: number;
}

/** One copy of the source stretched onto a destination. */
export interface StretchPlacement {
  readonly kind: 'stretch';
  /** Where the whole image goes, in the shape's own EMU, before cropping. */
  readonly image: Rect;
  /** The part of that destination the crop leaves visible. */
  readonly clip: Rect;
}

/** The lattice, and the cell that repeats on it. */
export interface TilePlacement {
  readonly kind: 'tile';
  /** One tile, in the shape's own EMU. */
  readonly tileW: number;
  readonly tileH: number;
  /** The top-left corner of the tile the `@algn` anchor placed. */
  readonly originX: number;
  readonly originY: number;
  readonly flipX: boolean;
  readonly flipY: boolean;
  /** Where the whole image sits inside one tile, once the crop is undone. */
  readonly image: Rect;
}

export type BlipPlacement = StretchPlacement | TilePlacement;

const ANCHOR: Readonly<Record<TileAlign, { x: number; y: number }>> = {
  tl: { x: 0, y: 0 },
  t: { x: 0.5, y: 0 },
  tr: { x: 1, y: 0 },
  l: { x: 0, y: 0.5 },
  ctr: { x: 0.5, y: 0.5 },
  r: { x: 1, y: 0.5 },
  bl: { x: 0, y: 1 },
  b: { x: 0.5, y: 1 },
  br: { x: 1, y: 1 },
};

/** A `CT_RelativeRect` as the fraction of the whole it leaves. */
function spans(rect: RelativeRect): { x: number; y: number; w: number; h: number } {
  const w = 1 - (rect.l + rect.r) / 100000;
  const h = 1 - (rect.t + rect.b) / 100000;
  return { x: rect.l / 100000, y: rect.t / 100000, w, h };
}

/**
 * The size one unscaled tile occupies.
 *
 * From the FULL image, not the cropped part: `tile-size-follows-crop` scores
 * 0.19 against this reading's 0.86 on the same probe.
 */
function naturalTile(fill: BlipFill, image: ImageSize): { w: number; h: number } {
  const dpi = fill.dpi > 0 ? fill.dpi : image.dpi;
  if (!(dpi > 0)) {
    throw new PaintError(
      'FILL_BLIP_RESOLUTION',
      `an image fill at ${String(dpi)} dpi`,
      'a:blipFill',
    );
  }
  return {
    w: (image.widthPx / dpi) * EMU_PER_INCH,
    h: (image.heightPx / dpi) * EMU_PER_INCH,
  };
}

/**
 * Where the image goes inside a box, given the crop that box shows.
 *
 * The visible fraction has to land exactly on the box, so the whole image is
 * larger than the box by the reciprocal of what the crop leaves.
 */
function undoCrop(box: Rect, srcRect: RelativeRect): Rect {
  const src = spans(srcRect);
  if (!(src.w > 0) || !(src.h > 0)) {
    throw new PaintError('FILL_BLIP_CROP', 'a:srcRect leaves no source', 'a:srcRect');
  }
  const w = box.w / src.w;
  const h = box.h / src.h;
  return { x: box.x - src.x * w, y: box.y - src.y * h, w, h };
}

/**
 * BT.709, which is what `a:grayscl`, `a:biLevel` and `a:duotone` all weigh by.
 *
 * BT.601 - the reading anyone writes from memory - scores 0.32 against this
 * one's 0.999 on the grayscale ramp.
 */
export const BLIP_LUMA = { r: 0.2126, g: 0.7152, b: 0.0722 } as const;

/**
 * The slope `a:lum/@contrast` applies about mid-grey.
 *
 * Not one formula: positive contrast divides and negative contrast adds, which
 * twelve probes agree on to four decimals.
 */
export function contrastScale(contrast: number): number {
  const c = contrast / 100000;
  return c >= 0 ? 1 / (1 - c) : 1 + c;
}

/**
 * What `a:lum/@bright` adds, in 0..1, once contrast has scaled the channel.
 *
 * The term grows with the contrast slope; treating brightness as independent of
 * it scores zero on every probe that carries both.
 */
export function brightOffset(bright: number, scale: number): number {
  return (bright / 100000) * 0.5 * (1 + scale);
}

/** Where an image fill draws, in the shape box's own coordinates. */
export function blipPlacement(fill: BlipFill, box: Rect, image: ImageSize): BlipPlacement {
  if (!(image.widthPx > 0) || !(image.heightPx > 0)) {
    throw new PaintError(
      'FILL_BLIP_SIZE',
      `an image ${String(image.widthPx)} by ${String(image.heightPx)} pixels`,
      'a:blip',
    );
  }

  if (fill.mode.kind === 'stretch') {
    const dest = spans(fill.mode.fillRect);
    const clip = {
      x: box.x + dest.x * box.w,
      y: box.y + dest.y * box.h,
      w: box.w * dest.w,
      h: box.h * dest.h,
    };
    return { kind: 'stretch', clip, image: undoCrop(clip, fill.srcRect) };
  }

  const tile = fill.mode;
  const natural = naturalTile(fill, image);
  const tileW = (natural.w * tile.sx) / 100000;
  const tileH = (natural.h * tile.sy) / 100000;
  if (!(tileW > 0) || !(tileH > 0)) {
    throw new PaintError('FILL_BLIP_SCALE', 'a:tile scales the image to nothing', 'a:tile');
  }

  const anchor = ANCHOR[tile.algn];
  return {
    kind: 'tile',
    tileW,
    tileH,
    originX: box.x + anchor.x * (box.w - tileW) + tile.tx,
    originY: box.y + anchor.y * (box.h - tileH) + tile.ty,
    flipX: tile.flip === 'x' || tile.flip === 'xy',
    flipY: tile.flip === 'y' || tile.flip === 'xy',
    image: undoCrop({ x: 0, y: 0, w: tileW, h: tileH }, fill.srcRect),
  };
}

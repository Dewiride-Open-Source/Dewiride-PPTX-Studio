/**
 * What a DrawingML fill is, before anything has been decided about it.
 *
 * Same laziness as `Color` in `types.ts`, for the same reason: a gradient stop
 * holds a `Color`, not an `Rgba`, so a theme swap re-resolves it for free and
 * the inspector can still say that a stop is `accent1` with a `tint`. Nothing
 * here resolves anything; `gradient.ts` and `pattern.ts` do that, and only when
 * asked.
 *
 * Values keep the units the file uses - hundred-thousandths for a percentage,
 * sixtieths of a degree for an angle - for the reason set out in `types.ts`.
 */

import type { Color } from '../types.js';

/* -------------------------------------------------------------------------- */
/* shared shapes                                                              */
/* -------------------------------------------------------------------------- */

/**
 * `CT_RelativeRect`: four **insets**, in hundred-thousandths, each defaulting to
 * zero.
 *
 * Insets, not coordinates. `r="100000"` means the right edge is inset by the
 * whole width, so it lands on the left edge - which is how PowerPoint's own
 * from-corner gradient writes a point focus at the top left.
 */
export interface RelativeRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

export const WHOLE_RECT: RelativeRect = { l: 0, t: 0, r: 0, b: 0 };

/** `ST_TileFlipMode`. */
export type TileFlipMode = 'none' | 'x' | 'y' | 'xy';

/* -------------------------------------------------------------------------- */
/* gradients                                                                  */
/* -------------------------------------------------------------------------- */

export interface GradientStop {
  /** `ST_PositiveFixedPercentage`, hundred-thousandths, as written. */
  readonly pos: number;
  readonly color: Color;
}

/** `a:lin` - a straight ramp at an angle. */
export interface LinearShade {
  readonly kind: 'linear';
  /**
   * Sixtieths of a degree. Zero points along +x so stop 0 is on the left, and
   * the angle increases **clockwise on screen**, y pointing down: `5400000` runs
   * top to bottom. Measured on six angles against a square shape, each landing
   * within a tenth of a degree of the value written.
   */
  readonly ang: number;
  /**
   * `@scaled`. False means the written angle is the visual angle. True means it
   * is measured in a space where the shape is a unit square - see
   * `linearGradientVector`, which is where the aspect correction actually
   * happens and where the direction it bends is recorded.
   */
  readonly scaled: boolean;
}

/** `a:path` - a ramp that spreads out from a focus. */
export interface PathShade {
  readonly kind: 'path';
  readonly path: 'shape' | 'circle' | 'rect';
  /**
   * `a:fillToRect`, or `null` when the element was absent - which is a different
   * focus from four zero insets, measured. ADR 0022.
   */
  readonly fillToRect: RelativeRect | null;
}

export type GradientShade = LinearShade | PathShade;

export interface GradientFill {
  readonly type: 'gradient';
  /** As written, in document order. `resolveGradientStops` is what sorts them. */
  readonly stops: readonly GradientStop[];
  /** `null` when the markup carried neither `a:lin` nor `a:path`. */
  readonly shade: GradientShade | null;
  /** `a:tileRect`, or `null` for the default. Insets, and they may be negative. */
  readonly tileRect: RelativeRect | null;
  /**
   * `@flip`. Recorded so it round-trips, and otherwise unused: measured to have
   * no effect on what PowerPoint paints. See `gradient.ts`.
   */
  readonly flip: TileFlipMode;
  /** `@rotWithShape`. Defaults to true, and PowerPoint writes it explicitly. */
  readonly rotWithShape: boolean;
}

/* -------------------------------------------------------------------------- */
/* the rest of EG_FillProperties                                              */
/* -------------------------------------------------------------------------- */

export interface PatternFill {
  readonly type: 'pattern';
  /** `ST_PresetPatternVal`. One of the 54 in `pattern-tiles.ts`. */
  readonly prst: string;
  /** The marks. `null` when the element omitted `a:fgClr`. */
  readonly fg: Color | null;
  /** The ground. `null` when the element omitted `a:bgClr`. */
  readonly bg: Color | null;
}

export interface SolidFill {
  readonly type: 'solid';
  readonly color: Color;
}

export interface NoFill {
  readonly type: 'none';
}

/**
 * `a:grpFill` - not `noFill`.
 *
 * It means "show the enclosing group's fill through this shape", which for a
 * gradient is the *slice* of the group's gradient under this shape's rectangle,
 * not a copy of it. Nothing here can compute that: it needs the group's
 * rectangle, which is 2.10's business. Modelled so a parser has somewhere to put
 * it and so a `switch` over `Fill` is exhaustive.
 */
export interface GroupFill {
  readonly type: 'group';
}

/** Where `a:tile` puts one whole tile against the shape box. ADR 0036. */
export type TileAlign = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';

/**
 * One `a:blip` colour effect, in document order.
 *
 * Every curve here is measured against real PowerPoint in
 * `corpus/ground-truth/blips.json`; the weights are BT.709 and the alpha is an
 * opacity. ADR 0036.
 */
export type BlipEffect =
  | { readonly kind: 'grayscale' }
  | { readonly kind: 'biLevel'; readonly thresh: number }
  | { readonly kind: 'lum'; readonly bright: number; readonly contrast: number }
  | { readonly kind: 'duotone'; readonly from: Color; readonly to: Color }
  | { readonly kind: 'alphaModFix'; readonly amt: number }
  | {
      readonly kind: 'clrChange';
      readonly from: Color;
      readonly to: Color;
      readonly useAlpha: boolean;
    };

/** `a:stretch` - one copy of the source, mapped onto the shape box. */
export interface BlipStretch {
  readonly kind: 'stretch';
  /** `a:fillRect`: insets of the destination, and they may be negative. */
  readonly fillRect: RelativeRect;
}

/** `a:tile` - the source repeated on a lattice. */
export interface BlipTile {
  readonly kind: 'tile';
  /** `@tx`/`@ty` in EMU, positive right and down, from the `@algn` anchor. */
  readonly tx: number;
  readonly ty: number;
  /** `@sx`/`@sy` in hundred-thousandths of the natural tile size. */
  readonly sx: number;
  readonly sy: number;
  /** `@flip`. Mirrors ALTERNATE tiles, so the visible period doubles. */
  readonly flip: TileFlipMode;
  readonly algn: TileAlign;
}

/** `a:blipFill` - an image fill. */
export interface BlipFill {
  readonly type: 'blip';
  /**
   * `a:blip/@r:embed`, which the caller resolves against the part's own rels -
   * the same contract `buBlip` uses.
   *
   * `null` when the blip names no raster at all, which is what PowerPoint
   * writes for a picture that is only an SVG. ADR 0037.
   */
  readonly embed: string | null;
  /** `asvg:svgBlip/@r:embed`: the vector original, drawn from 10.7. */
  readonly svgEmbed: string | null;
  /** The part whose rels `embed` is scoped to; rIds are per-part, never global. */
  readonly part: string;
  /** `a:srcRect`: insets of the SOURCE, as fractions of the image. */
  readonly srcRect: RelativeRect;
  readonly mode: BlipStretch | BlipTile;
  readonly effects: readonly BlipEffect[];
  /** `@dpi`. Zero, which is what PowerPoint writes, means the image's own. */
  readonly dpi: number;
  readonly rotWithShape: boolean;
}

export type Fill = NoFill | SolidFill | GradientFill | PatternFill | GroupFill | BlipFill;

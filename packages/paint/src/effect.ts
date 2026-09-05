/**
 * `a:effectLst` - shadows, glow, soft edges and blur.
 *
 * ## One blur, one constant
 *
 * DrawingML spells a blur radius four times on four elements - `a:outerShdw`
 * and `a:innerShdw` call it `@blurRad`, `a:blur`, `a:glow` and `a:softEdge` call
 * it `@rad` - and never says what a radius is. An SVG filter needs a standard
 * deviation, so the difference between the two is the whole problem.
 *
 * Measured across fourteen radii on four elements, it is one operator with one
 * constant:
 *
 *   **sigma = r / 3**, and the measured ratio never left 0.3300..0.3338.
 *
 * The edge itself moves for two of the four. A glow grows the shape before
 * blurring it and a soft edge shrinks it, both by **0.9 r**; a shadow and a
 * plain blur leave the edge alone. And a glow's `r` is *half* the radius it
 * declares, which is the only reason its numbers look different from the rest.
 *
 * The blur is isotropic and Gaussian, and neither is assumed: fitting an error
 * function to the measured edge left a worst residual of 0.84/255 over every
 * radius, and the ramp across a horizontal edge matched the ramp across a
 * vertical one to **0/255** over 75 samples.
 *
 * The evidence is `corpus/ground-truth/lines.json`; the reasoning is in
 * `docs/adr/phase-2-geometry-and-paint/0023-lines.md`.
 */

import { PaintError } from './errors.js';
import type { Color } from './types.js';

/** Sixtieths of a degree in one degree, as everywhere else in DrawingML. */
const ANGLE = 60000;

/**
 * The Gaussian standard deviation for a DrawingML blur radius, in the same unit.
 *
 * Three, exactly. Fourteen measurements across `a:outerShdw/@blurRad`,
 * `a:blur/@rad`, `a:glow/@rad` and `a:softEdge/@rad` put the ratio between
 * 0.3300 and 0.3338, and the sweep that produced them had a step of 0.02.
 */
export const BLUR_RADIUS_TO_SIGMA = 1 / 3;

/**
 * How far a glow grows the shape, and a soft edge shrinks it, per unit of `r`.
 *
 * Nine tenths, and it is a measurement rather than a derivation. Five soft-edge
 * radii from 2 to 32 points put the inset at 0.875, 0.9375, 0.906 and 0.902 of
 * the radius, and three glow radii put the outset at 0.469 of the *declared*
 * radius, which is 0.938 of the half-radius a glow actually uses. Every one of
 * those is within half a pixel of 0.9 at the resolution they were measured at,
 * and none of them is within half a pixel of 1.
 *
 * A renderer that grows by the whole radius is out by a tenth of it - three and
 * a half points on a 32pt soft edge, which is visible.
 */
export const EDGE_GROWTH_PER_RADIUS = 0.9;

/**
 * A glow uses half the radius it declares.
 *
 * With that halving, a glow's sigma and its outset are the same two constants
 * every other effect uses. Without it a glow appears to have a sigma of r/6 and
 * an outset of 0.47r, which look like two more constants to remember and are
 * not.
 */
export const GLOW_RADIUS_SCALE = 0.5;

export type RectAlignment = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';

/**
 * The anchor a shadow's scale and skew are applied about, when `@algn` is absent.
 *
 * Bottom centre - which is what the schema says, and, unusually for this
 * sub-phase, what Office does. A shadow doubled with `sx="200000" sy="200000"`
 * and no `@algn` grew 50% to each side horizontally and entirely upward
 * vertically, leaving its bottom edge where it was.
 */
export const DEFAULT_SHADOW_ALIGN: RectAlignment = 'b';

export interface ShadowGeometry {
  /** EMU. */
  readonly blurRad: number;
  /** EMU. */
  readonly dist: number;
  /** Sixtieths of a degree. Zero points along +x and the angle turns toward +y. */
  readonly dir: number;
  /** Thousandths of a percent, as written. 100000 is unity. */
  readonly sx: number;
  readonly sy: number;
  /** Sixtieths of a degree. */
  readonly kx: number;
  readonly ky: number;
  readonly algn: RectAlignment;
  readonly rotWithShape: boolean;
}

export interface OuterShadow extends ShadowGeometry {
  readonly kind: 'outerShdw';
  readonly color: Color;
}

export interface InnerShadow extends ShadowGeometry {
  readonly kind: 'innerShdw';
  readonly color: Color;
}

/**
 * `a:prstShdw`.
 *
 * Twenty presets that nothing documents, that no open renderer implements, and
 * that PowerPoint writes today: twenty of its forty-three legacy shadow styles
 * save as this element rather than as `a:outerShdw`. What each of the twenty
 * paints is recorded in `corpus/ground-truth/lines.json` under `presetShadows`.
 * They are not all plain offset shadows - several squash, stretch or skew.
 */
export interface PresetShadow {
  readonly kind: 'prstShdw';
  readonly prst: string;
  readonly dist: number;
  readonly dir: number;
  readonly color: Color;
}

export interface Glow {
  readonly kind: 'glow';
  /** EMU. */
  readonly rad: number;
  readonly color: Color;
}

export interface SoftEdge {
  readonly kind: 'softEdge';
  /** EMU. */
  readonly rad: number;
}

export interface Blur {
  readonly kind: 'blur';
  /** EMU. */
  readonly rad: number;
  readonly grow: boolean;
}

/**
 * `a:reflection`. Thirteen attributes, all of which PowerPoint's own nine
 * presets name explicitly, so none of them has to be guessed at. Not modelled
 * further here: 2.8 records what a reflection is and 2.10 decides how to paint
 * one.
 */
export interface Reflection {
  readonly kind: 'reflection';
  readonly blurRad: number;
  readonly stA: number;
  readonly stPos: number;
  readonly endA: number;
  readonly endPos: number;
  readonly dist: number;
  readonly dir: number;
  readonly fadeDir: number;
  readonly sx: number;
  readonly sy: number;
  readonly kx: number;
  readonly ky: number;
  readonly algn: RectAlignment;
  readonly rotWithShape: boolean;
}

export type Effect = OuterShadow | InnerShadow | PresetShadow | Glow | SoftEdge | Blur | Reflection;

/** The standard deviation for any DrawingML blur radius. Same unit in and out. */
export function blurSigma(radius: number): number {
  if (radius < 0) {
    throw new PaintError(
      'EFFECT_NEGATIVE_RADIUS',
      `a blur radius of ${String(radius)} EMU is negative`,
      String(radius),
    );
  }
  return radius * BLUR_RADIUS_TO_SIGMA;
}

/** How far a glow grows the shape and how much it blurs it, both in EMU. */
export function glowGeometry(rad: number): { grow: number; sigma: number } {
  const r = rad * GLOW_RADIUS_SCALE;
  return { grow: r * EDGE_GROWTH_PER_RADIUS, sigma: blurSigma(r) };
}

/** How far a soft edge pulls the shape in and how much it blurs it, in EMU. */
export function softEdgeGeometry(rad: number): { shrink: number; sigma: number } {
  return { shrink: rad * EDGE_GROWTH_PER_RADIUS, sigma: blurSigma(rad) };
}

/**
 * The shadow's translation, in EMU.
 *
 * `dir` is zero along +x and turns toward +y, which on a slide is clockwise
 * because y runs down the page. Measured at all eight multiples of 45 degrees
 * and cross-checked against `Shape.Shadow.OffsetX`/`OffsetY`, which agreed to
 * within the half-pixel the bitmap can resolve.
 */
export function shadowOffset(dist: number, dir: number): { dx: number; dy: number } {
  const radians = (dir / ANGLE) * (Math.PI / 180);
  return { dx: dist * Math.cos(radians), dy: dist * Math.sin(radians) };
}

export interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** The point of a box that `@algn` names. */
export function anchorPoint(algn: RectAlignment, box: Box): { x: number; y: number } {
  const left = algn === 'tl' || algn === 'l' || algn === 'bl';
  const right = algn === 'tr' || algn === 'r' || algn === 'br';
  const top = algn === 'tl' || algn === 't' || algn === 'tr';
  const bottom = algn === 'bl' || algn === 'b' || algn === 'br';
  return {
    x: left ? box.x : right ? box.x + box.w : box.x + box.w / 2,
    y: top ? box.y : bottom ? box.y + box.h : box.y + box.h / 2,
  };
}

/** A 2-D affine transform, in the order SVG's `matrix(a b c d e f)` wants them. */
export type Matrix = readonly [number, number, number, number, number, number];

/**
 * The whole transform a shadow applies to its shape: scale and skew about the
 * anchor `@algn` names, then the `dist`/`dir` translation.
 *
 * Each part was measured on its own against an untransformed control thrown
 * clear of the shape. `sx="200000"` doubled the box's width and left its
 * horizontal centre where it was; `sy="200000"` doubled its height and left its
 * *bottom* where it was; `sy="-100000"` mirrored it about that same bottom edge;
 * `kx="1200000"` widened the box by exactly `h * tan(20 degrees)`. All four are
 * consistent with one anchor, and with `@algn` absent that anchor is bottom
 * centre.
 */
export function shadowMatrix(shadow: ShadowGeometry, box: Box): Matrix {
  const anchor = anchorPoint(shadow.algn, box);
  const sx = shadow.sx / 100000;
  const sy = shadow.sy / 100000;
  const kx = Math.tan((shadow.kx / ANGLE) * (Math.PI / 180));
  const ky = Math.tan((shadow.ky / ANGLE) * (Math.PI / 180));
  const { dx, dy } = shadowOffset(shadow.dist, shadow.dir);

  // Translate the anchor to the origin, skew, scale, translate back, then offset.
  const a = sx;
  const b = ky * sx;
  const c = kx * sy;
  const d = sy;
  const e = anchor.x - (a * anchor.x + c * anchor.y) + dx;
  const f = anchor.y - (b * anchor.x + d * anchor.y) + dy;
  return [a, b, c, d, e, f];
}

/* -------------------------------------------------------------------------- */
/* the filter graph                                                           */
/* -------------------------------------------------------------------------- */

/**
 * One SVG filter primitive, as data.
 *
 * Data rather than markup, for the same reason `svgStops` in `gradient.ts` is:
 * this package has no DOM and no string templating, and the two renderers in
 * 2.10 build their nodes differently. A serializer for the debug gallery lives
 * with the gallery.
 */
export type FilterPrimitive =
  | {
      readonly op: 'gaussian';
      readonly in: string;
      readonly result: string;
      readonly stdDeviation: number;
    }
  | {
      readonly op: 'offset';
      readonly in: string;
      readonly result: string;
      readonly dx: number;
      readonly dy: number;
    }
  | {
      readonly op: 'flood';
      readonly result: string;
      readonly color: string;
      readonly opacity: number;
    }
  | {
      readonly op: 'composite';
      readonly in: string;
      readonly in2: string;
      readonly result: string;
      readonly operator: 'in' | 'out' | 'over' | 'atop' | 'xor';
    }
  | { readonly op: 'merge'; readonly inputs: readonly string[]; readonly result: string }
  | {
      readonly op: 'morphology';
      readonly in: string;
      readonly result: string;
      readonly operator: 'dilate' | 'erode';
      readonly radius: number;
    }
  | {
      readonly op: 'matrix';
      readonly in: string;
      readonly result: string;
      readonly matrix: Matrix;
    };

export interface FilterGraph {
  /** How far the effects reach outside the shape's own box, in EMU on each side. */
  readonly margin: { left: number; top: number; right: number; bottom: number };
  readonly primitives: readonly FilterPrimitive[];
  /** The `result` name holding the finished picture. */
  readonly result: string;
}

/** Three sigmas covers 99.7% of a Gaussian; beyond that there is nothing to see. */
const SIGMA_REACH = 3;

/**
 * How far one effect paints outside the shape it belongs to, in EMU.
 *
 * The filter region has to hold every one of them or SVG clips the result, and
 * the default region of -10%/+20% is not enough for an ordinary 24pt shadow on a
 * small shape.
 */
export function effectMargin(
  effect: Effect,
  box: Box,
): {
  left: number;
  top: number;
  right: number;
  bottom: number;
} {
  const none = { left: 0, top: 0, right: 0, bottom: 0 };
  switch (effect.kind) {
    case 'glow': {
      const { grow, sigma } = glowGeometry(effect.rad);
      const reach = grow + sigma * SIGMA_REACH;
      return { left: reach, top: reach, right: reach, bottom: reach };
    }
    case 'blur': {
      const reach = effect.grow ? blurSigma(effect.rad) * SIGMA_REACH : 0;
      return { left: reach, top: reach, right: reach, bottom: reach };
    }
    case 'outerShdw':
    case 'prstShdw': {
      const blurRad = effect.kind === 'outerShdw' ? effect.blurRad : 0;
      const reach = blurSigma(blurRad) * SIGMA_REACH;
      const corners =
        effect.kind === 'outerShdw'
          ? shadowBox(effect, box)
          : (() => {
              const { dx, dy } = shadowOffset(effect.dist, effect.dir);
              return { x: box.x + dx, y: box.y + dy, w: box.w, h: box.h };
            })();
      return {
        left: Math.max(0, box.x - (corners.x - reach)),
        top: Math.max(0, box.y - (corners.y - reach)),
        right: Math.max(0, corners.x + corners.w + reach - (box.x + box.w)),
        bottom: Math.max(0, corners.y + corners.h + reach - (box.y + box.h)),
      };
    }
    // An inner shadow and a soft edge paint only inside the shape.
    case 'innerShdw':
    case 'softEdge':
      return none;
    case 'reflection': {
      const { dx, dy } = shadowOffset(effect.dist, effect.dir);
      const height = box.h * Math.abs(effect.sy / 100000);
      return {
        left: Math.max(0, -dx),
        top: Math.max(0, -dy),
        right: Math.max(0, dx),
        bottom: Math.max(0, dy + height),
      };
    }
    default:
      return none;
  }
}

/** Where a shadow's own rectangle lands, given its full transform. */
export function shadowBox(shadow: ShadowGeometry, box: Box): Box {
  const [a, b, c, d, e, f] = shadowMatrix(shadow, box);
  const corners = [
    [box.x, box.y],
    [box.x + box.w, box.y],
    [box.x, box.y + box.h],
    [box.x + box.w, box.y + box.h],
  ].map(([x, y]) => [a * x! + c * y! + e, b * x! + d * y! + f] as const);
  const xs = corners.map((p) => p[0]);
  const ys = corners.map((p) => p[1]);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/**
 * The filter graph for one effect list.
 *
 * Painting order is a fact about Office and not about the document: the schema
 * makes `a:effectLst` a sequence, so a file cannot express an order, and the one
 * PowerPoint uses had to be measured. A shape carrying both a green glow and a
 * red shadow paints green over red where the two overlap - so the glow is above
 * the outer shadow, which is the reverse of the order the schema lists them in.
 *
 * The order here is therefore: outer shadow behind, then the shape with any
 * blur, soft edge and inner shadow applied to it, then the glow.
 */
export function effectFilter(
  effects: readonly Effect[],
  box: Box,
  colorOf: (color: Color) => { css: string; alpha: number },
): FilterGraph {
  const primitives: FilterPrimitive[] = [];
  const layers: string[] = [];
  let n = 0;
  const name = (): string => `e${String(n++)}`;

  const margin = { left: 0, top: 0, right: 0, bottom: 0 };
  for (const effect of effects) {
    const m = effectMargin(effect, box);
    margin.left = Math.max(margin.left, m.left);
    margin.top = Math.max(margin.top, m.top);
    margin.right = Math.max(margin.right, m.right);
    margin.bottom = Math.max(margin.bottom, m.bottom);
  }

  // Behind: every outer and preset shadow, in the order they appear.
  for (const effect of effects) {
    if (effect.kind !== 'outerShdw' && effect.kind !== 'prstShdw') continue;
    const { css, alpha } = colorOf(effect.color);
    const shifted = name();
    if (effect.kind === 'outerShdw') {
      primitives.push({
        op: 'matrix',
        in: 'SourceAlpha',
        result: shifted,
        matrix: shadowMatrix(effect, box),
      });
    } else {
      const { dx, dy } = shadowOffset(effect.dist, effect.dir);
      primitives.push({ op: 'offset', in: 'SourceAlpha', result: shifted, dx, dy });
    }
    let current = shifted;
    if (effect.kind === 'outerShdw' && effect.blurRad > 0) {
      const blurred = name();
      primitives.push({
        op: 'gaussian',
        in: current,
        result: blurred,
        stdDeviation: blurSigma(effect.blurRad),
      });
      current = blurred;
    }
    const flooded = name();
    primitives.push({ op: 'flood', result: flooded, color: css, opacity: alpha });
    const tinted = name();
    primitives.push({ op: 'composite', in: flooded, in2: current, result: tinted, operator: 'in' });
    layers.push(tinted);
  }

  // The shape itself, with anything that acts on it.
  let shape = 'SourceGraphic';
  for (const effect of effects) {
    if (effect.kind === 'blur') {
      const blurred = name();
      primitives.push({
        op: 'gaussian',
        in: shape,
        result: blurred,
        stdDeviation: blurSigma(effect.rad),
      });
      shape = blurred;
    } else if (effect.kind === 'softEdge') {
      const { shrink, sigma } = softEdgeGeometry(effect.rad);
      const eroded = name();
      primitives.push({
        op: 'morphology',
        in: shape,
        result: eroded,
        operator: 'erode',
        radius: shrink,
      });
      const blurred = name();
      primitives.push({ op: 'gaussian', in: eroded, result: blurred, stdDeviation: sigma });
      shape = blurred;
    }
  }
  layers.push(shape);

  // Inner shadows sit on top of the shape and are clipped to it.
  for (const effect of effects) {
    if (effect.kind !== 'innerShdw') continue;
    const { css, alpha } = colorOf(effect.color);
    // The hole is the shape's own alpha moved *along* `dir`, which darkens the
    // edge the direction points at. That is the opposite of what an outer
    // shadow does with the same angle, and it was measured in both directions
    // because one reading is not a rule.
    const { dx, dy } = shadowOffset(effect.dist, effect.dir);
    const moved = name();
    primitives.push({ op: 'offset', in: 'SourceAlpha', result: moved, dx, dy });
    let hole = moved;
    if (effect.blurRad > 0) {
      const blurred = name();
      primitives.push({
        op: 'gaussian',
        in: hole,
        result: blurred,
        stdDeviation: blurSigma(effect.blurRad),
      });
      hole = blurred;
    }
    const flooded = name();
    primitives.push({ op: 'flood', result: flooded, color: css, opacity: alpha });
    const outside = name();
    primitives.push({ op: 'composite', in: flooded, in2: hole, result: outside, operator: 'out' });
    const clipped = name();
    primitives.push({
      op: 'composite',
      in: outside,
      in2: 'SourceAlpha',
      result: clipped,
      operator: 'in',
    });
    layers.push(clipped);
  }

  // In front: the glow, which measurement puts above the outer shadow.
  for (const effect of effects) {
    if (effect.kind !== 'glow') continue;
    const { css, alpha } = colorOf(effect.color);
    const { grow, sigma } = glowGeometry(effect.rad);
    const grown = name();
    primitives.push({
      op: 'morphology',
      in: 'SourceAlpha',
      result: grown,
      operator: 'dilate',
      radius: grow,
    });
    const blurred = name();
    primitives.push({ op: 'gaussian', in: grown, result: blurred, stdDeviation: sigma });
    const flooded = name();
    primitives.push({ op: 'flood', result: flooded, color: css, opacity: alpha });
    const tinted = name();
    primitives.push({ op: 'composite', in: flooded, in2: blurred, result: tinted, operator: 'in' });
    // Behind the shape but in front of the shadow, so it goes under the shape
    // layer rather than over it.
    layers.splice(layers.indexOf(shape), 0, tinted);
  }

  const result = name();
  primitives.push({ op: 'merge', inputs: layers, result });
  return { margin, primitives, result };
}

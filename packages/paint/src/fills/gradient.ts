/**
 * Gradients: which stops win, what curve joins them, and where the ramp runs.
 *
 * Everything here was measured against PowerPoint in experiment C3 - 144 probes
 * across 23 packages, `corpus/ground-truth/fills.json`, written up in
 * `docs/adr/phase-2-geometry-and-paint/0022-fills.md`. Where a claim is inferred rather than measured it
 * says so at the point it is made.
 */

import { PaintError } from '../errors.js';
import { fromRampSpace, toRampSpace, TWO_STOP_RAMP, twoStopWeight } from './gradient-ramp.js';
import { resolveColor, type ColorContext } from '../colors/resolve.js';
import { toByte } from '../colors/transfer.js';
import type { GradientFill, GradientStop, LinearShade, PathShade, RelativeRect } from './fill.js';
import type { Rgba } from '../types.js';

const PERCENT = 100000;
const ANGLE = 60000;

/* -------------------------------------------------------------------------- */
/* the stop list                                                              */
/* -------------------------------------------------------------------------- */

/**
 * The stops in the order PowerPoint paints them: sorted by `@pos`, stably.
 *
 * **Document order is not paint order**, and this is not a hypothetical. Ask
 * PowerPoint through its own object model for a from-centre two-colour gradient
 * and it writes
 *
 * ```xml
 * <a:gs pos="50000"><a:srgbClr val="000000"/></a:gs>
 * <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
 * <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
 * ```
 *
 * and paints white at both ends with black in the middle. Its object model hands
 * the stops back in *file* order - `GradientStops.Item(1).Position` is 0.5 - so
 * the sort belongs to the renderer, not to the parser, and both facts have to be
 * kept: a writer that re-orders the list on save has edited a part nobody asked
 * it to touch.
 *
 * Ties keep document order, which matters for `h-duppos`: two stops at the same
 * position are legal, PowerPoint opens the package, and it paints a hard edge
 * from the first to the second.
 */
export function sortStops(stops: readonly GradientStop[]): readonly GradientStop[] {
  return stops
    .map((stop, i) => ({ stop, i }))
    .sort((a, b) => a.stop.pos - b.stop.pos || a.i - b.i)
    .map((x) => x.stop);
}

/** A stop with its colour resolved, position as a fraction. */
export interface ResolvedStop {
  /** 0 to 1. */
  readonly pos: number;
  readonly color: Rgba;
}

/** A gradient PowerPoint draws with its gamma-corrected two-colour path. */
export interface TwoColorRamp {
  /** The colour at factor 0 - whatever sits at position 0. */
  readonly from: Rgba;
  /** The other colour. */
  readonly to: Rgba;
  /** One factor per sorted stop: 0 where the colour is `from`, 1 where it is `to`. */
  readonly factors: readonly number[];
}

const key = (c: Rgba): string =>
  `${String(toByte(c.r))},${String(toByte(c.g))},${String(toByte(c.b))}`;

/**
 * Is this a gradient PowerPoint draws with a curve rather than linearly?
 *
 * Two conditions, and neither is the obvious one.
 *
 * **Two distinct colours**, not two stops. `[0 white][50% black][100% white]` -
 * which is what PowerPoint's own gallery writes for variant 4 of every gradient
 * style, and therefore what is in an enormous number of real decks - has three
 * stops and *is* curved: it paints `BABABA` a quarter of the way, where linear
 * would paint `808080`. `[0 black][50% 808080][100% white]` has the same three
 * positions, three distinct colours, and is linear to within a byte.
 *
 * **A stop at each end.** Two stops at 20% and 80% are linear, even though there
 * are only two colours.
 *
 * Those two conditions together are exactly the gradient GDI+ can express as a
 * two-colour `LinearGradientBrush` with a blend array - a construct whose blend
 * factors must start at 0 and end at 1, and which is the only one of the two
 * brush constructors that applies gamma correction. That is a mechanism, not a
 * coincidence, and it predicts all four of the measured cases.
 *
 * Returns the two colours and the per-stop factors, or `null` for an ordinary
 * multi-colour gradient.
 */
export function twoColorRamp(stops: readonly ResolvedStop[]): TwoColorRamp | null {
  if (stops.length < 2) return null;
  const first = stops[0]!;
  const last = stops[stops.length - 1]!;
  if (first.pos !== 0 || last.pos !== 1) return null;

  const zero = key(first.color);
  const other = stops.find((s) => key(s.color) !== zero);
  if (other === undefined) {
    // One colour throughout. Nothing to curve, but alpha may still vary, so it
    // is still the two-colour path with `to` equal to `from`.
    return { from: first.color, to: last.color, factors: stops.map((_, i) => (i === 0 ? 0 : 1)) };
  }
  const to = key(other.color);
  if (stops.some((s) => key(s.color) !== zero && key(s.color) !== to)) return null;

  return {
    from: first.color,
    to: other.color,
    factors: stops.map((s) => (key(s.color) === zero ? 0 : 1)),
  };
}

/**
 * Blend two colours at factor `f`.
 *
 * The colour channels go through the measured curve in the 2.2 power space;
 * **alpha is linear in the factor**. Measured: a white ramp from alpha 0 to
 * alpha 100% over red paints `FF4040` a quarter of the way, which is a linear
 * quarter and not the curve's 0.142 - the same split GDI+ makes, where gamma
 * correction touches colour and never the alpha channel.
 */
function blend(ramp: TwoColorRamp, f: number): Rgba {
  const w = twoStopWeight(f);
  const mix = (x: number, y: number): number =>
    fromRampSpace(toRampSpace(x) + (toRampSpace(y) - toRampSpace(x)) * w);
  return {
    r: mix(ramp.from.r, ramp.to.r),
    g: mix(ramp.from.g, ramp.to.g),
    b: mix(ramp.from.b, ramp.to.b),
    a: ramp.from.a + (ramp.to.a - ramp.from.a) * f,
  };
}

/**
 * The stops a renderer should emit, resolved and in paint order.
 *
 * For a two-colour gradient each segment where the factor changes is expanded
 * into the fifteen pre-sampled knots, because SVG and CSS both interpolate
 * linearly in sRGB and neither offers a way to ask for anything else -
 * `color-interpolation` is defined for gradients and implemented for them by
 * nobody. Pre-sampling is the only way to draw the curve at all.
 */
export function resolveGradientStops(
  fill: GradientFill,
  ctx: ColorContext = {},
): readonly ResolvedStop[] {
  const sorted = sortStops(fill.stops);
  if (sorted.length === 0) {
    throw new PaintError('FILL_NO_STOPS', 'a:gradFill has no gradient stops', null);
  }

  const resolved = sorted.map((stop) => ({
    pos: stop.pos / PERCENT,
    color: resolveColor(stop.color, ctx),
  }));

  const ramp = twoColorRamp(resolved);
  if (ramp === null) return resolved;

  const out: ResolvedStop[] = [];
  const push = (pos: number, color: Rgba): void => {
    if (out.length > 0 && out[out.length - 1]!.pos === pos) return;
    out.push({ pos, color });
  };

  for (let i = 0; i < resolved.length - 1; i++) {
    const a = resolved[i]!;
    const b = resolved[i + 1]!;
    const fa = ramp.factors[i]!;
    const fb = ramp.factors[i + 1]!;
    if (fa === fb) {
      push(a.pos, blend(ramp, fa));
      push(b.pos, blend(ramp, fb));
      continue;
    }
    // The knots are a curve from factor 0 to factor 1. A segment running the
    // other way walks them backwards.
    const rising = fb > fa;
    const knots = rising ? TWO_STOP_RAMP : [...TWO_STOP_RAMP].reverse();
    for (const [factor] of knots) {
      const u = rising ? factor : 1 - factor;
      push(a.pos + (b.pos - a.pos) * u, blend(ramp, factor));
    }
  }
  return out;
}

/**
 * The colour at `t` along a gradient - what a viewer will actually paint.
 *
 * Deliberately defined as linear interpolation between the stops
 * `resolveGradientStops` emits, for both the curved and the ordinary case. That
 * is not a shortcut: it is the only definition that agrees with what we can
 * cause a browser to draw, since SVG and CSS interpolate stop colours linearly
 * in sRGB and offer no way to ask for anything else.
 *
 * Defining it any other way makes this function and the renderer disagree, and
 * the renderer is the one the user sees. Evaluating the measured weight curve
 * directly instead is nineteen bytes away from PowerPoint on the steep dark end
 * of a black-to-white ramp; this is within two on every ramp in the fixture.
 *
 * Outside the first and last stop the colour is held flat rather than
 * extrapolated - measured on a ramp whose stops are at 20% and 80%, which is
 * pure black for its first fifth and pure white for its last.
 */
export function gradientColorAt(fill: GradientFill, t: number, ctx: ColorContext = {}): Rgba {
  const stops = resolveGradientStops(fill, ctx);
  const at = Math.min(1, Math.max(0, t));

  if (at <= stops[0]!.pos) return stops[0]!.color;
  const last = stops[stops.length - 1]!;
  if (at >= last.pos) return last.color;

  for (let i = 1; i < stops.length; i++) {
    const b = stops[i]!;
    if (at <= b.pos) {
      const a = stops[i - 1]!;
      const u = b.pos === a.pos ? 1 : (at - a.pos) / (b.pos - a.pos);
      return {
        r: a.color.r + (b.color.r - a.color.r) * u,
        g: a.color.g + (b.color.g - a.color.g) * u,
        b: a.color.b + (b.color.b - a.color.b) * u,
        a: a.color.a + (b.color.a - a.color.a) * u,
      };
    }
  }
  return last.color;
}

/* -------------------------------------------------------------------------- */
/* the linear ramp's geometry                                                 */
/* -------------------------------------------------------------------------- */

export interface GradientVector {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

/**
 * Where a linear gradient's ramp starts and ends, in shape coordinates.
 *
 * Two measured rules, and both are easy to get backwards.
 *
 * **The direction.** With `scaled="0"` the written angle is the visual angle -
 * a 45 degree ramp on a 3:1 shape measures 44.3 degrees. With `scaled="1"` the
 * angle is measured in a space where the shape is a unit square, and what gets
 * stretched is the *bands*, not the arrow: the direction of steepest change is
 * the **normal** to the stretched bands, so it comes out as
 * `atan2(w sin a, h cos a)`, not `atan2(h sin a, w cos a)`. On the same 3:1
 * shape a written 45 degrees measures 71.4 against a predicted 71.6, and the
 * inverted reading predicts 18.4. The 1:3 shape measures 18.8, and 30 degrees on
 * the wide shape measures 59.7 against 60.1 - three shapes, one formula.
 *
 * This is also why `gradientUnits="objectBoundingBox"` is the wrong SVG
 * construct for either value of `@scaled`. It reproduces the bend but it also
 * shears the ramp, putting the corner of a 3:1 shape at 0.21 where PowerPoint
 * puts it at 0.01. Compute the vector here and emit `userSpaceOnUse`.
 *
 * **The extent.** The ramp spans the shape's full projection onto that
 * direction, centred on the shape's centre, so the extreme corner along the
 * direction is exactly stop 0 and the opposite corner exactly the last stop.
 * Verified from the corners: at 45 degrees on a square, the two off-axis corners
 * both read 0.50; on a 3:1 shape with `scaled="0"` they read 0.75 and 0.25
 * against a predicted 0.751 and 0.249.
 */
export function linearGradientVector(shade: LinearShade, w: number, h: number): GradientVector {
  const written = ((shade.ang % (360 * ANGLE)) + 360 * ANGLE) % (360 * ANGLE);
  const radians = (written / ANGLE) * (Math.PI / 180);
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  // The aspect correction, when asked for. `w` and `h` are only ever a ratio
  // here, so a zero-extent shape degenerates to the written angle rather than to
  // a NaN.
  const bent = shade.scaled && w > 0 && h > 0 ? { x: h * cos, y: w * sin } : { x: cos, y: sin };
  const norm = Math.hypot(bent.x, bent.y);
  const dx = norm === 0 ? 1 : bent.x / norm;
  const dy = norm === 0 ? 0 : bent.y / norm;

  const length = Math.abs(w * dx) + Math.abs(h * dy);
  const cx = w / 2;
  const cy = h / 2;
  return {
    x1: cx - (length / 2) * dx,
    y1: cy - (length / 2) * dy,
    x2: cx + (length / 2) * dx,
    y2: cy + (length / 2) * dy,
  };
}

/* -------------------------------------------------------------------------- */
/* the path ramp's geometry                                                   */
/* -------------------------------------------------------------------------- */

/** Turn a `CT_RelativeRect`'s insets into a rectangle in 0..1 shape fractions. */
export function insetRect(rect: RelativeRect): {
  l: number;
  t: number;
  r: number;
  b: number;
} {
  return {
    l: rect.l / PERCENT,
    t: rect.t / PERCENT,
    r: 1 - rect.r / PERCENT,
    b: 1 - rect.b / PERCENT,
  };
}

export interface PathGeometry {
  /** The focus, in shape units from the shape's top-left corner. */
  readonly fx: number;
  readonly fy: number;
  /**
   * How the distance from the focus is measured.
   *
   * `radial` is a true circle - isotropic in shape units, so on a wide shape the
   * bands are circles and not ellipses. `box` is the Chebyshev distance, giving
   * concentric rectangles. `outline` follows the shape's own geometry, which for
   * a rectangle is `box` and for an ellipse is the ellipse.
   */
  readonly metric: 'radial' | 'box' | 'outline';
  /** For `radial`: the circle the last stop lands on, always the shape's own. */
  readonly cx: number;
  readonly cy: number;
  readonly radius: number;
  /** The shape, in the same units, which is what `box` normalises against. */
  readonly w: number;
  readonly h: number;
}

/**
 * Where a path gradient's ramp starts and how far it reaches.
 *
 * **Stop 0 is at the focus and the last stop at the edge**, which is the same
 * way round as SVG's `radialGradient`, so nothing is reversed on the way out.
 * The plan said path gradients reverse stop order; measured, they do not - a
 * black-to-white circle from the centre paints black in the middle.
 *
 * **`a:fillToRect` is insets, and only its centre matters.** A focus rect
 * covering the middle half of the shape paints byte-for-byte the same picture as
 * a degenerate point at the centre, over all 169 sampled positions, so the
 * rectangle's extent is not a flat region and does not need modelling. Two
 * spellings are not the centre: four zero insets are the **top-left corner**,
 * and no `a:fillToRect` element at all is the **centre** - which is why
 * `fillToRect` is nullable rather than defaulted.
 *
 * **`path="circle"` is a circle, not an ellipse, and an off-centre focus does
 * not move it.** The circle the last stop lands on is the shape's own - centred,
 * with the half-diagonal for a radius - and the focus only says where the ramp
 * begins, exactly as SVG's `fx`/`fy` do. Sixteen probes over eight foci and
 * three aspect ratios fit this to within half a byte; concentric circles about
 * the focus, however they are scaled, are out by up to 49, and the same
 * construction in a space where the shape is a unit square by up to 39.
 *
 * **`path="rect"`, `path="shape"` on a rectangle, and an absent `@path` are the
 * same** - Chebyshev, each axis normalised to the distance from the focus to the
 * edge it is heading for, over nine probes to within 1.6 bytes. `shape` differs
 * on anything that is not a rectangle: on an ellipse it follows the outline.
 */
export function pathGeometry(shade: PathShade, w: number, h: number): PathGeometry {
  const rect = shade.fillToRect;
  const inset = insetRect(rect ?? { l: 0, t: 0, r: 0, b: 0 });
  const corner = rect !== null && rect.l === 0 && rect.t === 0 && rect.r === 0 && rect.b === 0;
  const fx = rect === null ? 0.5 : corner ? 0 : (inset.l + inset.r) / 2;
  const fy = rect === null ? 0.5 : corner ? 0 : (inset.t + inset.b) / 2;

  return {
    fx: fx * w,
    fy: fy * h,
    metric: shade.path === 'circle' ? 'radial' : shade.path === 'rect' ? 'box' : 'outline',
    cx: w / 2,
    cy: h / 2,
    radius: Math.hypot(w, h) / 2,
    w,
    h,
  };
}

/**
 * How far along the ramp the point `(x, y)` is, for a path gradient.
 *
 * In shape units, clamped to 0..1. `outline` is treated as `box`, because
 * without the shape's geometry - which lives in `@pptx-studio/geometry` and is
 * deliberately not a dependency of this package - a rectangle is the only
 * outline available.
 */
export function pathPositionAt(geometry: PathGeometry, x: number, y: number): number {
  const dx = x - geometry.fx;
  const dy = y - geometry.fy;
  let t: number;
  if (geometry.metric === 'radial') {
    // The contour through (x, y) is the circle that has grown from the focus
    // towards the shape's own circle: centre at F + s(C - F), radius sR. Solve
    // |P - F - s(C - F)| = sR for s, which is what SVG's fx/fy paint.
    const ex = geometry.cx - geometry.fx;
    const ey = geometry.cy - geometry.fy;
    const a = ex * ex + ey * ey - geometry.radius * geometry.radius;
    const b = -2 * (dx * ex + dy * ey);
    const c = dx * dx + dy * dy;
    const disc = b * b - 4 * a * c;
    if (disc < 0) {
      t = 1;
    } else {
      // The stable pair, because `a` goes to zero whenever the focus lands on
      // the outer circle - which a corner focus on a square shape does exactly,
      // and which the textbook formula turns into a division by almost nothing.
      const q = -0.5 * (b + (b >= 0 ? Math.sqrt(disc) : -Math.sqrt(disc)));
      const roots: number[] = [];
      if (a !== 0) roots.push(q / a);
      if (q !== 0) roots.push(c / q);
      // The focus is inside the outer circle, so `a` is never positive and the
      // two roots straddle zero: exactly one survives the filter.
      const usable = roots.filter((r) => r >= 0 && Number.isFinite(r));
      t = c === 0 ? 0 : usable.length === 0 ? 1 : Math.min(...usable);
    }
  } else {
    // Chebyshev, each axis normalised to the distance from the focus to the edge
    // it is heading for - which reduces to half the extent for a centred focus
    // and reproduces the corner focus PowerPoint's own from-corner fill writes.
    const rx = dx >= 0 ? geometry.w - geometry.fx : geometry.fx;
    const ry = dy >= 0 ? geometry.h - geometry.fy : geometry.fy;
    t = Math.max(rx === 0 ? 0 : Math.abs(dx) / rx, ry === 0 ? 0 : Math.abs(dy) / ry);
  }
  return t < 0 ? 0 : t > 1 ? 1 : t;
}

/* -------------------------------------------------------------------------- */
/* emitting                                                                   */
/* -------------------------------------------------------------------------- */

/** `#RRGGBB` when opaque, `rgba(...)` when not - the same rule as `toCss`. */
function css(color: Rgba): string {
  const r = toByte(color.r);
  const g = toByte(color.g);
  const b = toByte(color.b);
  const hex = [r, g, b].map((c) => c.toString(16).padStart(2, '0').toUpperCase()).join('');
  return color.a >= 1
    ? `#${hex}`
    : `rgba(${String(r)}, ${String(g)}, ${String(b)}, ${String(Math.round(color.a * 1000) / 1000)})`;
}

/** One `<stop>`'s attributes, ready for either renderer. */
export interface SvgStop {
  readonly offset: number;
  readonly color: string;
  readonly opacity: number;
}

/**
 * The stop list an SVG `<linearGradient>` or `<radialGradient>` needs.
 *
 * Opacity is emitted as `stop-opacity` rather than folded into an `rgba()`
 * colour, because SVG interpolates stop colour and stop opacity separately and
 * premultiplied - a fade from an opaque colour to `rgba(0,0,0,0)` darkens
 * through the middle, and keeping the colour constant while only the opacity
 * moves is the only way to get PowerPoint's fade.
 */
export function svgStops(fill: GradientFill, ctx: ColorContext = {}): readonly SvgStop[] {
  return resolveGradientStops(fill, ctx).map((stop) => ({
    offset: stop.pos,
    color: css({ ...stop.color, a: 1 }),
    opacity: Math.round(stop.color.a * 1000) / 1000,
  }));
}

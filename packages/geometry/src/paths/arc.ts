import { FULL_CIRCLE, angleToRadians } from './formula.js';
import type { Point } from './types.js';

/**
 * `a:arcTo`, the one drawing command that is not what it looks like.
 *
 * Three things about it are counter-intuitive, and each of them breaks a
 * different set of shapes when it is got wrong.
 *
 * ## 1. The angles are not parametric
 *
 * `stAng` is not the `t` in `(cx + wR cos t, cy + hR sin t)`. It is the angle of
 * a **ray from the centre**, and on an ellipse a ray at 30 degrees does not meet
 * the curve at parameter 30 degrees. The two agree only on a circle and at the
 * four cardinal angles - which is exactly why the mistake survives casual
 * testing. Converting one to the other is `unskewAngle`.
 *
 * This is not read off a specification: the preset data states it directly.
 * `pie` computes its own arc start point as
 *
 *     wt1 = sin wd2 stAng      ->  wd2 * sin(stAng)
 *     ht1 = cos hd2 stAng      ->  hd2 * cos(stAng)
 *     dx1 = cat2 wd2 ht1 wt1   ->  wd2 * cos(atan2(wt1, ht1))
 *     dy1 = sat2 hd2 ht1 wt1   ->  hd2 * sin(atan2(wt1, ht1))
 *     x1  = hc + dx1,  y1 = vc + dy1
 *
 * which is `centre + (wR cos t, hR sin t)` with `t = atan2(wR sin a, hR cos a)`,
 * spelled out in guide arithmetic. `pie`, `chord`, `arc` and `blockArc` all do
 * it, so the shapes carry their own statement of the formula and their own
 * independently computed answer to check against.
 *
 * ## 2. The arc is placed by its start point, not by a centre
 *
 * `a:arcTo` names no centre and no destination. It begins at the current point,
 * and the centre is whatever puts the point at `stAng` there:
 *
 *     centre = current - (wR cos t0, hR sin t0)
 *
 * The end point then falls out of `stAng + swAng`. So an arc cannot be read on
 * its own; it needs the point the previous command left behind. Across the 187
 * presets there are 393 arcs and **not one** of them opens a path, so that
 * current point always exists.
 *
 * ## 3. A negative sweep really does mean backwards
 *
 * 81 of the 393 arcs have a literal negative `swAng` and `donut` is the reason
 * it matters: its outer ring is four arcs of +90 degrees and its inner ring is
 * four of -90. The rings wind in opposite directions, and under the nonzero
 * fill rule that opposite winding **is** the hole. Take the absolute value, or
 * normalise -90 to 270, and the doughnut fills in solid.
 *
 * ## Angles increase clockwise
 *
 * y grows downward, so `cd4` (90 degrees) is the bottom of the shape and
 * `3cd4` the top. `ellipse` traces left, top, right, bottom with four sweeps of
 * +90 starting from `cd2`, which pins the direction without appeal to anything
 * outside the data.
 */

const TWO_PI = Math.PI * 2;

/** A full turn of sweep, in 60000ths of a degree. `swAng` is clamped to plus or minus this. */
export const ARC_MAX_SWEEP = FULL_CIRCLE;

/** One `a:arcTo`, after its four operands have been evaluated. */
export interface ArcParameters {
  /** Horizontal radius. Note this is a radius, not a diameter: presets pass `wd2`. */
  readonly wR: number;
  readonly hR: number;
  /** Start angle in 60000ths of a degree, as a ray from the centre. Not a parameter. */
  readonly stAng: number;
  /** Swing angle. Signed: negative sweeps anticlockwise on screen. */
  readonly swAng: number;
}

/** Everything an arc turns out to be, once its start point is known. */
export interface ArcGeometry {
  readonly centre: Point;
  readonly start: Point;
  readonly end: Point;
  /** The radii actually used - the absolute values of `wR` and `hR`. */
  readonly rx: number;
  readonly ry: number;
  /** Parametric start angle in radians. */
  readonly startT: number;
  /**
   * Parametric end angle in radians, lifted rather than wrapped: it is
   * `startT + delta` for the true signed sweep, so it may lie outside
   * `[-pi, pi]` and its difference from `startT` is the parametric sweep.
   */
  readonly endT: number;
  /** The swing angle actually used, after clamping. Same units as `swAng`. */
  readonly sweptAngle: number;
  /**
   * True when the arc has no extent to draw: a zero radius, or a non-finite
   * operand. Start, end and centre are still filled in as far as they can be.
   */
  readonly degenerate: boolean;
}

/**
 * One arc an SVG `A` command can actually draw.
 *
 * SVG's elliptical arc is defined by its two endpoints, so it cannot express a
 * closed loop: a command whose destination equals its origin draws nothing at
 * all. Twenty of the presets' arcs are exactly that - `smileyFace`'s outline
 * and both its eyes, `sun`, `actionButtonInformation`, `funnel`'s hole - so the
 * split is not a theoretical nicety.
 */
export interface ArcSegment {
  readonly from: Point;
  readonly to: Point;
  readonly rx: number;
  readonly ry: number;
  /** SVG `large-arc-flag`: this piece spans more than half a turn. */
  readonly largeArc: boolean;
  /** SVG `sweep-flag`: true when the sweep is positive, which on screen is clockwise. */
  readonly sweep: boolean;
}

/**
 * `swAng` clamped to a single turn in either direction, which is what Office
 * does.
 *
 * Clamping is not the same as normalising and the difference is visible: 400
 * degrees clamped is 360 and ends where it began, normalised is 40 and ends
 * somewhere else entirely. The distinction only matters for the commands that
 * follow the arc, and no preset reaches it - the largest swing in the corpus is
 * exactly one turn - so this follows the documented Office behaviour and is
 * pinned by a test rather than by evidence from a real file.
 */
export function clampSweep(swAng: number): number {
  if (Number.isNaN(swAng)) return swAng;
  if (swAng > ARC_MAX_SWEEP) return ARC_MAX_SWEEP;
  if (swAng < -ARC_MAX_SWEEP) return -ARC_MAX_SWEEP;
  return swAng;
}

/**
 * Turn a DrawingML ray angle into the ellipse parameter that reaches the same
 * point: `t = atan2(wR sin a, hR cos a)`.
 *
 * Lifted rather than wrapped. `atan2` alone would fold every angle back into a
 * half-open turn, which loses the winding an arc needs; here a whole number of
 * turns is taken out first and added back afterwards, so the result is
 * continuous and strictly increasing in `angle` across every turn boundary.
 * That is what lets the parametric sweep be read off as a plain subtraction,
 * including for the sweeps that exceed half a turn.
 *
 * The four cardinal angles are fixed points, whatever the radii - which is
 * precisely why an implementation that skips this step looks correct on
 * `ellipse`, `donut` and every rounded rectangle, and is wrong on `moon`.
 */
export function unskewAngle(angle: number, wR: number, hR: number): number {
  const turns = Math.floor((angle + FULL_CIRCLE / 2) / FULL_CIRCLE);
  if (!Number.isFinite(turns)) return Number.NaN;
  const base = angleToRadians(angle - turns * FULL_CIRCLE);
  return Math.atan2(wR * Math.sin(base), hR * Math.cos(base)) + turns * TWO_PI;
}

/** The point at parameter `t` on the ellipse centred at `centre`. */
function ellipsePoint(centre: Point, rx: number, ry: number, t: number): Point {
  return { x: centre.x + rx * Math.cos(t), y: centre.y + ry * Math.sin(t) };
}

/**
 * Resolve one arc against the point the path has reached.
 *
 * The radii are used as absolute values. A negative one describes the same
 * curve as its positive twin - it only mirrors the parameterisation - and SVG
 * requires the absolute value, so taking it here keeps the two agreeing. No
 * preset produces a negative radius at any size, which a test holds to.
 */
export function arcGeometry(from: Point, arc: ArcParameters): ArcGeometry {
  const rx = Math.abs(arc.wR);
  const ry = Math.abs(arc.hR);
  const sweptAngle = clampSweep(arc.swAng);

  const startT = unskewAngle(arc.stAng, rx, ry);
  const endT = unskewAngle(arc.stAng + sweptAngle, rx, ry);

  const degenerate =
    !Number.isFinite(rx) ||
    !Number.isFinite(ry) ||
    !Number.isFinite(startT) ||
    !Number.isFinite(endT) ||
    rx === 0 ||
    ry === 0;

  const centre: Point = {
    x: from.x - rx * Math.cos(startT),
    y: from.y - ry * Math.sin(startT),
  };

  return {
    centre,
    start: ellipsePoint(centre, rx, ry, startT),
    end: ellipsePoint(centre, rx, ry, endT),
    rx,
    ry,
    startT,
    endT,
    sweptAngle,
    degenerate,
  };
}

/** Where an arc leaves the pen. The only thing a path walker needs from it. */
export function arcEnd(from: Point, arc: ArcParameters): Point {
  return arcGeometry(from, arc).end;
}

/**
 * The point a given fraction of the way along the arc, `0` being the start and
 * `1` the end. Fractions outside that range extrapolate along the same ellipse.
 */
export function arcPointAt(geometry: ArcGeometry, fraction: number): Point {
  const t = geometry.startT + (geometry.endT - geometry.startT) * fraction;
  return ellipsePoint(geometry.centre, geometry.rx, geometry.ry, t);
}

/**
 * How close to a whole turn still counts as one, in radians.
 *
 * Generous by floating-point standards because the input is not a float: the
 * angles arrive as integer 60000ths of a degree through several formulas, so a
 * sweep that means a whole turn arrives a few units of last place away from one
 * and there is nothing between "a whole turn" and "a whole turn less a
 * thousandth of a degree" that any file means on purpose.
 */
const CLOSED_EPSILON = 1e-9;

/**
 * Break an arc into pieces an SVG `A` command can draw, in order.
 *
 * One piece for anything short of a whole turn. A whole turn becomes two half
 * turns, because an `A` whose destination is its origin is a no-op in SVG - it
 * is not drawn as a circle, it is not drawn at all - and `smileyFace` would
 * lose its outline and both eyes to it.
 *
 * Empty for a degenerate arc. A caller that wants a straight line there should
 * say so itself; this returns what there is to draw, which is nothing.
 */
export function arcSegments(from: Point, arc: ArcParameters): ArcSegment[] {
  const geometry = arcGeometry(from, arc);
  if (geometry.degenerate) return [];

  const delta = geometry.endT - geometry.startT;
  if (delta === 0) return [];

  const span = Math.abs(delta);
  const pieces = span >= TWO_PI - CLOSED_EPSILON ? Math.ceil(span / Math.PI) : 1;
  const step = delta / pieces;
  const sweep = delta > 0;
  const largeArc = Math.abs(step) > Math.PI;

  const segments: ArcSegment[] = [];
  let cursor = geometry.start;
  for (let i = 1; i <= pieces; i += 1) {
    const to =
      i === pieces
        ? geometry.end
        : ellipsePoint(geometry.centre, geometry.rx, geometry.ry, geometry.startT + step * i);
    segments.push({ from: cursor, to, rx: geometry.rx, ry: geometry.ry, largeArc, sweep });
    cursor = to;
  }
  return segments;
}

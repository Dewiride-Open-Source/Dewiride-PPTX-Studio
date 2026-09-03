import {
  evaluateGuides,
  resolveOperand,
  type EvaluateOptions,
  type ShapeSize,
} from './evaluate.js';
import { FULL_CIRCLE, type FormulaSite } from './formula.js';
import type { Geometry, Point, PresetAdjustHandle, PresetGuide } from './types.js';

/**
 * The yellow handles, and what happens when one is dragged.
 *
 * Everything up to 2.4 ran forwards: a definition and a size produce numbers,
 * and the numbers produce a path. A handle runs the other way. The user grabs a
 * dot, moves it to a point, and that point has to become a new value for one of
 * the shape's adjust guides - which is then written into `a:prstGeom/a:avLst`
 * and read back by PowerPoint.
 *
 * ## What a handle is, once the data is counted
 *
 * 120 of the 187 presets carry one; there are 243 handles between them, 220
 * `a:ahXY` and 23 `a:ahPolar` across ten shapes, and 287 adjustable axes. Every
 * one of those 287 names an adjust value - `adj`, or `adj1` through `adj8` -
 * and every one of those names is declared in that preset's own `avLst`. Not
 * one names a computed guide. So the thing a drag writes is always an adjust
 * value, which is the only thing `a:avLst` can hold.
 *
 * ## `minX` and `maxX` bound the guide, not the position
 *
 * They are `ST_AdjCoordinate`, so they can be a guide name, and 92 of the 287
 * axes use one. But they are read against the guide's own units rather than
 * against the shape: `arc` declares `maxAng="21599999"`, which is one turn
 * short of a full circle in 60000ths of a degree and is not a coordinate, and
 * the four callout families declare `maxX="2147483647"`, which is the largest
 * `ST_Coordinate` and means *unbounded* rather than a point 2.1 billion units
 * to the right.
 *
 * ## An axis is a scalar readout of the handle's position
 *
 * This is the decision the module turns on, and it is what lets one routine
 * serve all 287 axes.
 *
 * `a:ahXY` names its axes by coordinate: `gdRefX` moves the handle's x and
 * `gdRefY` its y. `a:ahPolar` names them by radius and angle, and gives no
 * origin to measure either from. So there are four readouts of one position -
 * `x`, `y`, `r`, `ang` - and an axis is one readout paired with one adjust
 * guide. Dragging is solving `readout(v) = target` for `v`.
 *
 * The two axes of a handle are independent: across all 220 `a:ahXY` handles,
 * sweeping the guide named by `gdRefY` through its whole range never moves
 * `pos.x` at all. So two axes are two 1-D solves and never a 2-D one.
 *
 * ## The polar origin is the shape centre, and the presets say so
 *
 * `a:ahPolar` does not name one, and it matters: a radius and an angle are
 * meaningless without it. The preset data answers it, because ten of the
 * eighteen angle axes place `pos` exactly on the ray at their own angle.
 *
 * Sweep `arc`, `blockArc`, `chord`, `pie`, `circularArrow` or `leftCircularArrow`
 * through 49 angles at six aspect ratios - 200x100, 100x100, 100x400,
 * 914400x685800, 40x900 and 3x1 - and compute `atan2(pos.y - h/2, pos.x - w/2)`.
 * It reproduces the guide named by `gdRefAng` to within **2.2e-8 of a 60000th
 * of a degree**, which is one unit in the last place of the arctangent and not
 * a modelling error. At 200x100 the locus is an ellipse, and the radius runs
 * from 50 to 100 across the sweep - yet the angle still lands exactly. That is
 * 2.3's finding restated from the other side: an angle here names a ray out of
 * the centre, not a parameter of an ellipse.
 *
 * So the origin is `(hc, vc)`. Measured, not assumed.
 *
 * ## But the ray identity is not a licence to read the angle off the mouse
 *
 * The other eight angle axes are not on their own ray. `circularArrow`'s
 * innermost handle is out by up to 156 degrees, its stub handle by a constant
 * 19.04, and `mathNotEqual`'s angle handle slides along the shape's top edge
 * and is out by 171. Setting the guide to the mouse's angle would be exactly
 * right for ten axes and visibly wrong for eight.
 *
 * Inverting the readout is right for all eighteen, and for the ten it costs one
 * extra step of arithmetic. So the readout is always inverted and never assumed.
 *
 * ## What the other implementations do, and why neither approach was taken
 *
 * Nobody inverts the formula symbolically and nobody bisects.
 *
 * LibreOffice derives the inverse by hand, per preset, dispatching on the shape
 * name, with Newton's method for four ellipse-inset radius handles and a
 * catch-all that assumes the guide is the coordinate over the width or height
 * times 100000. Its own source and design notes say there is no general method
 * and that the approach cannot work for a custom geometry, because it needs the
 * formulas in advance.
 *
 * ONLYOFFICE calibrates instead: it evaluates the shape at the referenced
 * guide's declared minimum and again at its maximum, reads where the handle
 * landed each time, and fits one linear coefficient. That is exact for the
 * affine majority, silently wrong for the rest, and it never checks that the
 * handle ended up under the pointer. It carries exactly one shape-specific
 * branch, for `mathNotEqual`, which it corrects by half a turn - the same
 * anomaly the ray measurement above puts at 171 degrees.
 *
 * Both take a polar angle straight from an arctangent about the shape centre,
 * which is the origin above arrived at from the implementer's side rather than
 * from the data. Apache POI models an adjust handle as an empty marker
 * interface and has no inverse at all; python-pptx exposes adjust values with
 * no clamping and no knowledge that `ahLst` exists.
 *
 * The method below is neither. It needs no formulas in advance, so a
 * `custGeom` works, and it verifies rather than assuming, so the shapes
 * calibration gets wrong are the ones it spends its extra probes on.
 *
 * ## One method, because false position is the analytic answer
 *
 * The map from an adjust value to a readout, composed through the whole guide
 * chain, is **affine at every size on 243 of the 287 axes** and monotone on
 * 282. That is the shape of the plan's "analytic for the dominant forms,
 * bisection for the rest" - but it does not need two methods, and it does not
 * need to look at the formula text at all.
 *
 * Sample the readout across the axis's range, take the sub-interval that
 * straddles the target, and refine with false position. On an affine map the
 * first false-position step lands on the root exactly, so the analytic case is
 * covered by being a special case rather than by being a separate code path.
 * Whether a map is affine is a property of the whole composed chain and not of
 * any one `fmla`, so pattern-matching the multiply-divide form that dominates
 * the guide lists would have found some of the 243 and missed the rest.
 * (Spelled out rather than quoted, as in `types.ts`: that operator's own two
 * characters end a block comment.)
 *
 * ## Why a scan, rather than bracketing the endpoints
 *
 * Because five of the 1435 axis-by-size combinations genuinely fold back.
 * `curvedLeftArrow` and `curvedRightArrow` at 100x100 run their handle down by
 * 0.615 units a step for thirty-nine steps and then back up by 0.179 on the
 * fortieth; `circularArrow`, `leftCircularArrow` and `leftRightCircularArrow`
 * swing their innermost handle 147 degrees backwards once, at 40x900, when it
 * crosses the centre of a shape 40 units wide.
 *
 * A fold means more than one adjust value can put the handle where the mouse
 * is. The scan finds every candidate interval and the one nearest the value the
 * shape currently holds wins, so a handle moves continuously under the pointer
 * instead of teleporting to the far branch. When nothing straddles the target -
 * the drag went past the end of the range - the nearest reachable value wins,
 * which for a monotone axis is exactly `min` or `max`.
 *
 * ## The readout matters more than the method
 *
 * Read a polar handle's angle axis as `pos.y` instead of as an angle and the
 * numbers collapse: 58 of the classifications turn non-monotone rather than 5,
 * and `mathNotEqual`'s angle handle becomes literally constant, because its
 * `pos.y` is `t` and never moves. Twelve shapes stop being invertible at all.
 * Choosing the right scalar is the whole of the problem; the root-finder is the
 * easy half.
 */

/** Which scalar of a handle's position an axis moves. */
export type HandleAxisKind = 'x' | 'y' | 'r' | 'ang';

/** How many points the readout is sampled at before a bracket is chosen. */
export const HANDLE_SAMPLES = 33;

/**
 * When the search stops, in adjust units.
 *
 * Half a unit, because an adjust value is written to the file as an integer and
 * half a unit is below the resolution of anything that can be stored. One unit
 * is a thousandth of a percent on a proportional axis and a 60000th of a degree
 * on an angular one; both are far finer than a screen pixel at any zoom.
 */
export const HANDLE_TOLERANCE = 0.5;

/** A hard stop on refinement. Reached only if the readout is pathological. */
const MAX_REFINEMENTS = 64;

/** One adjustable axis of one handle, resolved at a size. */
export interface ResolvedHandleAxis {
  readonly kind: HandleAxisKind;
  /** The adjust value this axis writes: `gdRefX`, `gdRefY`, `gdRefR` or `gdRefAng`. */
  readonly guide: string;
  /** What that guide holds as the shape currently stands. */
  readonly value: number;
  /**
   * The bounds the search and the clamp use, ordered so `min <= max`.
   *
   * The declared `minX`/`maxX` pair, widened where it had to be - see `widened`.
   */
  readonly min: number;
  readonly max: number;
  /**
   * True when the declared bounds did not contain the shape's own value and
   * were widened to admit it. 16 of the 1435 axis-and-size combinations.
   */
  readonly widened: boolean;
  /** The readout at `value`: a coordinate, a radius, or an angle in 60000ths of a degree. */
  readonly readout: number;
}

/** One `a:ahXY` or `a:ahPolar`, resolved at a size. */
export interface ResolvedHandle {
  readonly kind: 'xy' | 'polar';
  /** Where to draw it, in shape space. */
  readonly pos: Point;
  /** One entry per declared `gdRef`. Zero, one or two - either axis may be absent. */
  readonly axes: readonly ResolvedHandleAxis[];
  /** False if the position or a bound is not finite, which a zero-sized shape reaches. */
  readonly finite: boolean;
}

export interface ResolveHandleOptions extends EvaluateOptions {
  /** The preset's name, for the message when an operand does not resolve. */
  readonly name?: string | null;
}

/**
 * The origin a polar handle's radius and angle are measured from.
 *
 * `(hc, vc)`. See the note above: ten of the eighteen angle axes put `pos` on
 * the ray at their own angle to within one unit in the last place, at every
 * aspect ratio tried.
 */
export function handleCentre(size: ShapeSize): Point {
  return { x: size.w / 2, y: size.h / 2 };
}

/** Read one scalar off a handle position. Angles come back in 60000ths of a degree. */
export function readHandleAxis(kind: HandleAxisKind, pos: Point, centre: Point): number {
  switch (kind) {
    case 'x':
      return pos.x;
    case 'y':
      return pos.y;
    case 'r':
      return Math.hypot(pos.x - centre.x, pos.y - centre.y);
    case 'ang':
      return (Math.atan2(pos.y - centre.y, pos.x - centre.x) * (FULL_CIRCLE / 2)) / Math.PI;
  }
}

interface AxisSpec {
  readonly kind: HandleAxisKind;
  readonly guide: string;
  readonly min: string | null;
  readonly max: string | null;
}

/**
 * The axes a handle declares, in a form that does not care which element it came from.
 *
 * Either axis may be absent - a handle that only moves horizontally has a
 * `gdRefX` and no `gdRefY` - and a missing `gdRef` means that axis is not
 * adjustable, which is different from adjustable through a range of zero.
 */
function axesOf(handle: PresetAdjustHandle): AxisSpec[] {
  const list: AxisSpec[] = [];
  if (handle.kind === 'xy') {
    if (handle.gdRefX !== null) {
      list.push({ kind: 'x', guide: handle.gdRefX, min: handle.minX, max: handle.maxX });
    }
    if (handle.gdRefY !== null) {
      list.push({ kind: 'y', guide: handle.gdRefY, min: handle.minY, max: handle.maxY });
    }
  } else {
    // Angle before radius, and the order is load-bearing. See `dragHandle`.
    if (handle.gdRefAng !== null) {
      list.push({ kind: 'ang', guide: handle.gdRefAng, min: handle.minAng, max: handle.maxAng });
    }
    if (handle.gdRefR !== null) {
      list.push({ kind: 'r', guide: handle.gdRefR, min: handle.minR, max: handle.maxR });
    }
  }
  return list;
}

/**
 * The caller's adjust values, reduced to numbers.
 *
 * `AdjustOverrides` may be a guide list carrying arbitrary formulas, and a
 * search that moves one adjust value has to hold the others still. Running
 * their formulas again on every probe would be both slower and, if one of them
 * referenced the value being moved, wrong.
 */
function settleAdjust(
  geometry: Geometry,
  size: ShapeSize,
  options: ResolveHandleOptions,
): Record<string, number> {
  const settled = evaluateGuides(geometry, size, options);
  const base: Record<string, number> = {};
  const hold = (name: string): void => {
    const value = settled.get(name);
    if (value !== undefined) base[name] = value;
  };
  for (const gd of geometry.avLst) hold(gd.name);
  const adjust = options.adjust;
  if (Array.isArray(adjust)) {
    for (const gd of adjust as readonly PresetGuide[]) hold(gd.name);
  } else if (adjust !== undefined) {
    for (const key of Object.keys(adjust)) hold(key);
  }
  return base;
}

/**
 * The bounds of an axis, resolved, ordered, and widened to admit the shape's
 * own value.
 *
 * ## An absent bound freezes the axis
 *
 * It does not free it. ISO 29500-1 says it per attribute - if `maxX` is
 * omitted "it is assumed that this adjust handle cannot move in the x
 * direction. That is the maxX and minX are equal" - and it never says equal to
 * what, so the value the shape holds is the only answer available.
 *
 * That is the opposite of the obvious reading, and it is unreachable from the
 * presets: all 243 records state both bounds or neither `gdRef` at all. It is
 * reachable from a `custGeom`, where freeing an axis the file froze would let a
 * drag write a value the author said could not change.
 *
 * A bound of 2147483647 written out in full is a different thing entirely - a
 * stated bound that happens to be the widest `ST_Coordinate` there is, which is
 * how the four callout families say "anywhere". That one is honoured as stated.
 *
 * The two are swapped if the file states them the other way round; no preset
 * does, and a `custGeom` may.
 *
 * ## Why a declared bound is sometimes not a bound
 *
 * A range that excludes the value the shape actually holds is not a range for
 * that value, so the smallest correction is to admit it. That happens on 16 of
 * the 1435 axis-and-size combinations, in two quite different ways.
 *
 * `mathDivide` ships `adj2="5880"` while its own `maxAdj2` works out to 2930,
 * and `leftRightRibbon` ships `adj2="50000"` against a `maxAdj2` of 46875. Both
 * already clamp themselves - `a2` is `pin 0 adj2 maxAdj2` - so the shape draws
 * identically either way and the declared bound is honest. Widening here is
 * what stops merely touching a handle from rewriting a deck's own number to a
 * different one that draws the same, which in a package built on preserving
 * what it did not edit is the whole point.
 *
 * `star24` and `star32` are a defect in the shape definitions. Both declare
 * `maxY="ssd2"` - half the shorter side, a coordinate - for a handle whose
 * `gdRefY` is an adjust value in thousandths of a percent, and whose own
 * `gdLst` clamps it with `pin 0 adj 50000`. The other five stars in the family
 * declare `maxY="50000"`.
 *
 * Widening does not fix that, and it is worth being exact about why. Because
 * `ssd2` scales with the shape, the damage depends on the unit the caller is
 * working in - which is the one thing this package otherwise never lets happen.
 * On a shape 100 units tall the bound is 50 against a guide that runs to 50000,
 * so the declared travel is a two-thousandth of the real one and widening is
 * what keeps the handle moving at all. On the same shape in EMU it is 342900,
 * nearly seven times the real travel, and most of the range is a flat region
 * above the shape's own `pin`. Usable in both, and differently in each.
 *
 * Whether Office enforces `ahLst` bounds at all is not answerable from the
 * corpus and is written down as an experiment.
 */
function bounds(
  spec: AxisSpec,
  guides: ReadonlyMap<string, number>,
  site: FormulaSite,
  value: number,
): { min: number; max: number; widened: boolean } {
  if (spec.min === null || spec.max === null) return { min: value, max: value, widened: false };
  const lo = resolveOperand(spec.min, guides, site);
  const hi = resolveOperand(spec.max, guides, site);
  const min = Math.min(lo, hi);
  const max = Math.max(lo, hi);
  if (!Number.isFinite(value) || (value >= min && value <= max)) {
    return { min, max, widened: false };
  }
  return { min: Math.min(min, value), max: Math.max(max, value), widened: true };
}

/**
 * Every handle of a geometry, resolved at a size.
 *
 * Separate from `resolveGeometry` on purpose. A slide draws every shape and
 * edits one, so the handles of the 199 shapes nobody has selected are work
 * nobody asked for.
 */
export function resolveHandles(
  geometry: Geometry,
  size: ShapeSize,
  options: ResolveHandleOptions = {},
): readonly ResolvedHandle[] {
  const name = options.name ?? geometry.name;
  const site: FormulaSite = { preset: name, guide: null };
  const guides = evaluateGuides(geometry, size, options);
  const centre = handleCentre(size);

  return geometry.ahLst.map((handle) => {
    const pos: Point = {
      x: resolveOperand(handle.pos.x, guides, site),
      y: resolveOperand(handle.pos.y, guides, site),
    };
    let finite = Number.isFinite(pos.x) && Number.isFinite(pos.y);
    const axes = axesOf(handle).map((spec) => {
      const value = resolveOperand(spec.guide, guides, site);
      const { min, max, widened } = bounds(spec, guides, site, value);
      if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(value)) finite = false;
      return {
        kind: spec.kind,
        guide: spec.guide,
        value,
        min,
        max,
        widened,
        readout: readHandleAxis(spec.kind, pos, centre),
      } satisfies ResolvedHandleAxis;
    });
    return { kind: handle.kind, pos, axes, finite };
  });
}

/**
 * Lift the sampled angles onto one continuous branch.
 *
 * `atan2` returns a value in a half-open turn, so a handle swept through a full
 * circle comes back as a sawtooth. Undoing that is what makes the straddle test
 * below mean anything for an angle.
 */
function unwrap(values: number[]): void {
  let turns = 0;
  for (let i = 1; i < values.length; i++) {
    const previous = (values[i - 1] as number) - turns * FULL_CIRCLE;
    const raw = values[i] as number;
    if (raw - previous > FULL_CIRCLE / 2) turns--;
    else if (raw - previous < -FULL_CIRCLE / 2) turns++;
    values[i] = raw + turns * FULL_CIRCLE;
  }
}

/** An angular difference, folded into half a turn either side of zero. */
function foldAngle(difference: number): number {
  let d = difference % FULL_CIRCLE;
  if (d > FULL_CIRCLE / 2) d -= FULL_CIRCLE;
  if (d <= -FULL_CIRCLE / 2) d += FULL_CIRCLE;
  return d;
}

/**
 * False position with the Illinois correction, on a bracket known to straddle.
 *
 * The first step is exact whenever the readout is affine on the bracket, which
 * is 243 of the 287 axes at every size. The Illinois halving is what stops the
 * other 43 from creeping towards a stuck endpoint one side at a time.
 */
function refine(
  f: (v: number) => number,
  a0: number,
  b0: number,
  fa0: number,
  fb0: number,
): number {
  let a = a0;
  let b = b0;
  let fa = fa0;
  let fb = fb0;
  let side = 0;

  // Close enough in readout terms to stop, relative to how far the bracket's
  // ends are from the target. An affine map lands on the root on the first step
  // but rarely to the last bit, and without this the search would then spend a
  // dozen probes bisecting a residual of 1e-15.
  const settled = 1e-12 * (Math.abs(fa0) + Math.abs(fb0));

  for (let i = 0; i < MAX_REFINEMENTS; i++) {
    if (b - a <= HANDLE_TOLERANCE) break;
    const denominator = fb - fa;
    let c = denominator === 0 ? (a + b) / 2 : (fb * a - fa * b) / denominator;
    if (!Number.isFinite(c) || c <= a || c >= b) c = (a + b) / 2;
    const fc = f(c);
    if (!Number.isFinite(fc) || Math.abs(fc) <= settled) return c;
    // Keep the end whose sign the root is not on. Halving the *retained* end's
    // value is the Illinois correction, and it is what stops false position
    // creeping towards a stuck endpoint one probe at a time.
    if (fc < 0 === fa < 0) {
      a = c;
      fa = fc;
      if (side === 1) fb /= 2;
      side = 1;
    } else {
      b = c;
      fb = fc;
      if (side === -1) fa /= 2;
      side = -1;
    }
  }
  const denominator = fb - fa;
  const linear = denominator === 0 ? (a + b) / 2 : (fb * a - fa * b) / denominator;
  return linear >= a && linear <= b ? linear : (a + b) / 2;
}

export interface InvertAxisOptions extends ResolveHandleOptions {
  /**
   * The value to stay near when a fold makes more than one answer possible.
   * Defaults to the axis's current value, which is what a drag wants.
   */
  readonly near?: number;
}

/**
 * The adjust value that puts this axis's readout at `target`.
 *
 * Not rounded and not clamped to the integers a file can hold - see
 * `dragHandle` for that. Clamped to the axis's own `min` and `max`, because
 * those are what the shape says the handle may do.
 *
 * `target` is a coordinate for an `x` or `y` axis, a distance from the centre
 * for an `r` axis, and an angle in 60000ths of a degree for an `ang` axis. An
 * angle need not be normalised: it is lifted by whole turns onto the branch
 * nearest the axis's current value, so a drag winding past twelve o'clock
 * carries on rather than snapping back.
 */
export function invertAxis(
  geometry: Geometry,
  size: ShapeSize,
  handle: PresetAdjustHandle,
  axis: ResolvedHandleAxis,
  target: number,
  options: InvertAxisOptions = {},
): number {
  const name = options.name ?? geometry.name;
  const site: FormulaSite = { preset: name, guide: null };
  const centre = handleCentre(size);
  const near = options.near ?? axis.value;

  const base = settleAdjust(geometry, size, options);

  const readout = (v: number): number => {
    const guides = evaluateGuides(geometry, size, { adjust: { ...base, [axis.guide]: v } });
    const pos: Point = {
      x: resolveOperand(handle.pos.x, guides, site),
      y: resolveOperand(handle.pos.y, guides, site),
    };
    return readHandleAxis(axis.kind, pos, centre);
  };

  const { min, max } = axis;
  if (!(max > min)) return min;

  const vs: number[] = [];
  const ys: number[] = [];
  for (let k = 0; k < HANDLE_SAMPLES; k++) {
    const v = min + ((max - min) * k) / (HANDLE_SAMPLES - 1);
    const y = readout(v);
    if (!Number.isFinite(y)) continue;
    vs.push(v);
    ys.push(y);
  }
  if (vs.length < 2) return near;
  if (axis.kind === 'ang') unwrap(ys);

  // For an angle the target names a direction, so every whole turn of it is the
  // same direction and each is a candidate. For everything else there is one.
  const targets: number[] = [target];
  if (axis.kind === 'ang') {
    const low = Math.min(...ys);
    const high = Math.max(...ys);
    const first = Math.floor((low - target) / FULL_CIRCLE);
    const last = Math.ceil((high - target) / FULL_CIRCLE);
    targets.length = 0;
    for (let n = first; n <= last && targets.length < 8; n++)
      targets.push(target + n * FULL_CIRCLE);
    if (targets.length === 0) targets.push(target);
  }

  let best: {
    a: number;
    b: number;
    fa: number;
    fb: number;
    target: number;
    anchor: number;
    distance: number;
  } | null = null;
  let fallback: { v: number; error: number } | null = null;

  // A value that already puts the handle on the target stays put.
  //
  // Flat regions are real and common: a preset's own `gdLst` clamps its adjust
  // values with `pin`, so `mathDivide` draws identically for every `adj2` from
  // 2930 to its declared default of 5880. Solving inside such a region would
  // pick some arbitrary member of it and rewrite a deck's own number to another
  // one that looks the same, which in a package whose first principle is
  // preserving what it did not edit is the wrong answer.
  const spread = Math.max(...ys) - Math.min(...ys);
  // An axis whose readout does not move is not adjustable, whatever the file
  // says. No preset reaches this - in none of the 220 `a:ahXY` handles does
  // `pos` name the guide in `gdRefX` directly - but a `custGeom` can name a
  // guide nothing downstream reads, and the alternative is a drag that silently
  // writes the range minimum every time.
  if (spread === 0) return near;
  const still = 1e-9 * spread;
  const held = readout(Math.min(max, Math.max(min, near)));
  if (Number.isFinite(held)) {
    for (const t of targets) {
      const error = axis.kind === 'ang' ? Math.abs(foldAngle(held - t)) : Math.abs(held - t);
      if (error <= still) return near;
    }
  }

  for (const t of targets) {
    for (let k = 0; k < vs.length; k++) {
      const error = Math.abs((ys[k] as number) - t);
      if (fallback === null || error < fallback.error) fallback = { v: vs[k] as number, error };
    }
    for (let k = 0; k + 1 < vs.length; k++) {
      const fa = (ys[k] as number) - t;
      const fb = (ys[k + 1] as number) - t;
      // A sample that lands exactly on the target is not short-circuited: zero
      // is not greater than zero, so it becomes a bracket like any other and
      // the tie-break below still gets to choose which one.
      if (fa * fb > 0) continue;
      const a = vs[k] as number;
      const b = vs[k + 1] as number;
      const distance = near < a ? a - near : near > b ? near - b : 0;
      if (best === null || distance < best.distance) {
        best = { a, b, fa, fb, target: t, anchor: ys[k] as number, distance };
      }
    }
  }

  if (best === null) return fallback === null ? near : fallback.v;

  // Inside the bracket the readout moves by at most one sample step, so an
  // angle can be lifted back onto the bracket's own branch by folding against
  // its left end. Without that, refinement would compare an `atan2` result
  // against a target that has been wound on by whole turns.
  const { anchor, target: aim } = best;
  const residual =
    axis.kind === 'ang'
      ? (v: number): number => anchor + foldAngle(readout(v) - anchor) - aim
      : (v: number): number => readout(v) - aim;
  const solved = refine(residual, best.a, best.b, best.fa, best.fb);
  return Math.min(max, Math.max(min, solved));
}

/**
 * Round an adjust value to something a file can hold.
 *
 * `ST_AdjCoordinate` is a guide name or an integer, so the value written is an
 * integer. Rounding happens after the clamp and is then pulled back inside the
 * bounds, because a bound may be fractional - `trapezoid`'s `maxAdj` is
 * `50000*w/ss`, which is 66666.67 on a 914400x685800 shape - and rounding a
 * clamped value can step past it.
 */
export function roundAdjust(value: number, min: number, max: number): number {
  const clamped = Math.min(max, Math.max(min, value));
  const rounded = Math.round(clamped);
  if (rounded < min) return Math.min(Math.ceil(min), Math.floor(max));
  if (rounded > max) return Math.max(Math.floor(max), Math.ceil(min));
  return rounded;
}

/**
 * Drag one handle to a point, and get back the adjust values to write.
 *
 * The result holds one entry per adjustable axis - at most two - and nothing
 * else, so a caller can merge it into the shape's `a:avLst` without disturbing
 * the adjust values this handle does not control.
 *
 * `to` is in shape space, the same coordinates `pos` comes back in. A polar
 * handle's radius and angle are both measured from the shape centre.
 *
 * ## The axes are solved one after another, and the order matters
 *
 * Each axis is solved against the values already solved, not against the shape
 * as it stood before the drag.
 *
 * For an `a:ahXY` that changes nothing, and it was measured rather than
 * assumed: across all 220 of them, sweeping the guide named by `gdRefY` through
 * its whole range never moves `pos.x`.
 *
 * The four two-axis polar handles are not decoupled, and `blockArc` shows why.
 * Its handle sits at `dx2 = iwd2*cos(u)`, `dy2 = ihd2*sin(u)` where `u` is the
 * unskewed form of the angle, so the ray out of the centre is exactly the angle
 * guide and does not depend on the radius guide at all - while the radius
 * depends on both. Solve them simultaneously against the pre-drag state and the
 * handle lands **21% of the shape** away from the pointer at 200x100 and 14.8%
 * at 100x400. Solve the angle first and the radius against the answer, and it
 * lands within 0.0002%.
 *
 * On a square shape `iwd2` equals `ihd2` and the error vanishes entirely, which
 * is exactly why a square-only test would never have found it. The three
 * circular arrows are not exactly solvable this way - their handle is confined
 * to a curve the pointer can leave - but sequential beats simultaneous on every
 * sample there too, by up to a factor of seven.
 */
export function dragHandle(
  geometry: Geometry,
  size: ShapeSize,
  handle: PresetAdjustHandle,
  to: Point,
  options: ResolveHandleOptions = {},
): Record<string, number> {
  const name = options.name ?? geometry.name;
  const site: FormulaSite = { preset: name, guide: null };
  const centre = handleCentre(size);
  const base = settleAdjust(geometry, size, options);

  const result: Record<string, number> = {};
  for (const spec of axesOf(handle)) {
    const adjust = { ...base, ...result };
    const guides = evaluateGuides(geometry, size, { adjust });
    const pos: Point = {
      x: resolveOperand(handle.pos.x, guides, site),
      y: resolveOperand(handle.pos.y, guides, site),
    };
    const value = resolveOperand(spec.guide, guides, site);
    const { min, max, widened } = bounds(spec, guides, site, value);
    if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(value)) continue;
    const axis: ResolvedHandleAxis = {
      kind: spec.kind,
      guide: spec.guide,
      value,
      min,
      max,
      widened,
      readout: readHandleAxis(spec.kind, pos, centre),
    };
    const target = readHandleAxis(spec.kind, to, centre);
    const solved = invertAxis(geometry, size, handle, axis, target, { ...options, adjust, name });
    // Rounded before the next axis sees it, because the rounded value is what
    // the file will hold and what the shape will be drawn from.
    result[spec.guide] = roundAdjust(solved, min, max);
  }
  return result;
}

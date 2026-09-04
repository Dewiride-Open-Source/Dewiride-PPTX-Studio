/**
 * Where a shape ends up, and which way round it is.
 *
 * Every rule here was measured against Microsoft PowerPoint in experiment C6
 * (`tools/ground-truth/`, fixture `corpus/ground-truth/transforms.json`), and
 * the fixture carries the raw cases so `render.test.ts` re-derives them rather
 * than trusting this comment.
 *
 * ## Flip happens before rotate
 *
 * `a:xfrm` carries `@rot`, `@flipH` and `@flipV` and says nothing about the
 * order they compose in. The two readings are not equivalent: for any
 * reflection `F`, `F R(t) F = R(-t)`, so `R(t) F` and `F R(t)` differ by the
 * sign of the angle.
 *
 * PowerPoint answered it in writing. Handed a shape already rotated 30 degrees
 * and asked to mirror it, it wrote `rot="19800000" flipH="1"` - **minus** 30
 * with a flip - and did the same at 45, 120 and 200 degrees and on both axes.
 * Only a renderer that flips first has to negate the angle to reproduce a
 * mirrored figure. Both flips at once left the angle alone, which is the same
 * fact from the other side: `flipH flipV` is a half turn, and rotations commute.
 *
 * So a shape's own transform is `translate . rotate . flip`, about the centre of
 * its `a:ext` box, and the same order governs where a group puts its children:
 * of the 42 composed positions C6 read back, the flip-first model gets 42 and
 * the rotate-first model 41 - the one it misses being the only probe with a
 * group that is *both* turned and mirrored, which is the only case where the two
 * can disagree.
 *
 * ## A group's children are written in their own coordinate system
 *
 * `a:chOff` and `a:chExt` declare it, and the map is
 *
 * ```text
 *   x_parent = off.x + (x_child - chOff.x) * (ext.cx / chExt.cx)
 * ```
 *
 * with the two axes independent. Scored over the same 42 cases: this rule 42,
 * dropping the scale factor 32, ignoring `chOff` 3. Dropping the scale is the
 * one to fear - it is right on three quarters of the probes and on every group
 * nobody has resized, which is every group until a user drags one.
 *
 * Two zeros, and they are not the same zero. **A `chExt` of zero means no
 * scaling on that axis**, not a collapse - measured, and a divisor would be a
 * division by zero anyway. An `ext` of zero really is zero: a group with no
 * width renders its children with no width. And a group with **no `a:chOff` or
 * `a:chExt` at all** uses `chOff = 0` with no scaling, so its children's
 * coordinates are offsets from the group's own corner - not, as one might
 * assume, `chOff = off`.
 */

import type { Xfrm } from '@pptx-studio/model';

/** 914400 EMU to the inch, 72 points to the inch. */
export const EMU_PER_POINT = 12700;

/** Sixtieths of a degree: the unit of `a:xfrm/@rot`. */
export const ANGLE_UNITS_PER_DEGREE = 60000;

/** An axis-aligned rectangle in EMU. */
export interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * A placed shape's rectangle and orientation, in slide EMU.
 *
 * `cx` and `cy` are the **unrotated** extent, which is what PowerPoint's own
 * object model reports and what the geometry is evaluated at.
 */
export interface Frame extends Box {
  /** Degrees, clockwise on screen. Normalised to `[0, 360)`. */
  readonly rot: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

export const IDENTITY_FRAME: Frame = {
  x: 0,
  y: 0,
  cx: 0,
  cy: 0,
  rot: 0,
  flipH: false,
  flipV: false,
};

function normalizeDegrees(value: number): number {
  const wrapped = value % 360;
  return wrapped < 0 ? wrapped + 360 : wrapped;
}

/** An `a:xfrm` as a frame, with `@rot` converted out of sixtieths of a degree. */
export function frameOf(xfrm: Xfrm): Frame {
  return {
    x: xfrm.x,
    y: xfrm.y,
    cx: xfrm.cx,
    cy: xfrm.cy,
    rot: normalizeDegrees(xfrm.rot / ANGLE_UNITS_PER_DEGREE),
    flipH: xfrm.flipH,
    flipV: xfrm.flipV,
  };
}

/* -------------------------------------------------------------------------- */
/* the child coordinate system                                                */
/* -------------------------------------------------------------------------- */

/** The affine part of a group's child map: a scale and an origin, per axis. */
export interface ChildSpace {
  readonly sx: number;
  readonly sy: number;
  readonly ox: number;
  readonly oy: number;
}

export const UNIT_CHILD_SPACE: ChildSpace = { sx: 1, sy: 1, ox: 0, oy: 0 };

/**
 * The child map of a group, given the rectangle the group actually occupies.
 *
 * `frame` is the group's **composed** frame, so a group inside a group already
 * carries its parent's scaling and the factors multiply on their own.
 */
export function childSpace(frame: Box, xfrm: Xfrm): ChildSpace {
  const child = xfrm.child;
  // No `a:chOff`/`a:chExt` at all: the children are offsets from the group's
  // own corner, at scale one. Measured - `chOff = off` is the plausible reading
  // and it is wrong by exactly `off`.
  if (child === null) return UNIT_CHILD_SPACE;
  return {
    // A `chExt` of zero means the axis is not scaled.
    sx: child.cx === 0 ? 1 : frame.cx / child.cx,
    sy: child.cy === 0 ? 1 : frame.cy / child.cy,
    ox: child.x,
    oy: child.y,
  };
}

/* -------------------------------------------------------------------------- */
/* the extent swap                                                            */
/* -------------------------------------------------------------------------- */

/**
 * Whether a group's two scale factors change places for a child at this angle.
 *
 * A rotated rectangle under a map that scales the two axes differently is a
 * parallelogram, and `a:xfrm` has nowhere to put one - an offset, an extent, an
 * angle and two mirrors, and no shear. So PowerPoint approximates, and what it
 * does is snap the child's angle to the nearest quadrant and, when that quadrant
 * is a quarter or three quarters of a turn, give the child's width the group's
 * *vertical* factor and its height the horizontal one.
 *
 * Measured at eighteen angles. The rounding is the part worth having: at 45
 * degrees the swap happens and at 135 it does not, so the tie goes upward both
 * times and `|sin| > |cos|` - the form anyone writes first - is wrong at one end
 * or the other. Scored: this rule 18/18, `|sin| > |cos|` 14/18, `|sin| >= |cos|`
 * 17/18.
 *
 * Harmless when the scale is uniform, where swapping two equal factors changes
 * nothing, so it is applied unconditionally rather than guarded.
 */
export function swapsExtents(degrees: number): boolean {
  return Math.floor((normalizeDegrees(degrees) + 45) / 90) % 2 === 1;
}

/* -------------------------------------------------------------------------- */
/* composition                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Compose two turns.
 *
 * Writing a turn as `R(a) F` where `F` is one of `I`, `X`, `Y` or `XY`:
 *
 * ```text
 *   R(a) F1 R(b) F2  =  R(a + s1 b) (F1 F2),   s1 = -1 when F1 mirrors one axis
 * ```
 *
 * because `X R(b) = R(-b) X` and `X Y = R(180)`, which commutes. So the outer
 * turn's rotation and the inner's **subtract** whenever the outer mirrors
 * exactly one axis - which is why a group flipped horizontally and holding a
 * child at 40 degrees reads back as 320, and why summing the angles gets 9 of
 * C6's 13 orientation probes rather than 13.
 */
export function composeTurn(
  outer: { rot: number; flipH: boolean; flipV: boolean },
  inner: { rot: number; flipH: boolean; flipV: boolean },
): { rot: number; flipH: boolean; flipV: boolean } {
  const mirrorsOneAxis = outer.flipH !== outer.flipV;
  return {
    rot: normalizeDegrees(outer.rot + (mirrorsOneAxis ? -inner.rot : inner.rot)),
    flipH: outer.flipH !== inner.flipH,
    flipV: outer.flipV !== inner.flipV,
  };
}

/** A point turned about the origin: mirrored first, then rotated. */
export function turnVector(
  turn: { rot: number; flipH: boolean; flipV: boolean },
  x: number,
  y: number,
): { x: number; y: number } {
  const fx = turn.flipH ? -x : x;
  const fy = turn.flipV ? -y : y;
  const radians = (turn.rot * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return { x: fx * cos - fy * sin, y: fx * sin + fy * cos };
}

/**
 * Where a group puts one of its children.
 *
 * `group` is the group's composed frame in slide EMU, `space` its child map, and
 * `child` the child's own `a:xfrm` as written. Four measured steps, in order:
 *
 * 1. map the rectangle through the child coordinate system;
 * 2. scale the extents, exchanging the two factors when the child's angle says
 *    to - the centre stays where step 1 put it, which was true of all eighteen
 *    shear probes;
 * 3. turn the centre about the group's centre, mirroring before rotating;
 * 4. compose the two orientations.
 */
export function placeChild(group: Frame, space: ChildSpace, child: Frame): Frame {
  const x = group.x + (child.x - space.ox) * space.sx;
  const y = group.y + (child.y - space.oy) * space.sy;
  const swap = swapsExtents(child.rot);
  const cx = child.cx * (swap ? space.sy : space.sx);
  const cy = child.cy * (swap ? space.sx : space.sy);

  // The centre of the naively mapped rectangle, which is where it stays.
  const centreX = x + (child.cx * space.sx) / 2;
  const centreY = y + (child.cy * space.sy) / 2;

  const groupCentreX = group.x + group.cx / 2;
  const groupCentreY = group.y + group.cy / 2;
  const turned = turnVector(group, centreX - groupCentreX, centreY - groupCentreY);

  const turn = composeTurn(group, child);
  return {
    x: groupCentreX + turned.x - cx / 2,
    y: groupCentreY + turned.y - cy / 2,
    cx,
    cy,
    rot: turn.rot,
    flipH: turn.flipH,
    flipV: turn.flipV,
  };
}

/* -------------------------------------------------------------------------- */
/* emission                                                                   */
/* -------------------------------------------------------------------------- */

function trim(value: number): string {
  if (!Number.isFinite(value)) return '0';
  const rounded = Math.round(value * 1000) / 1000;
  return Object.is(rounded, -0) ? '0' : String(rounded);
}

/**
 * The SVG `transform` that puts a shape's local space on the slide.
 *
 * The geometry is evaluated in a box from `(0, 0)` to `(cx, cy)`, so the
 * transform has to move that box to the frame, turn it about its own centre,
 * and mirror it *inside* the rotation. SVG applies a transform list right to
 * left, so the order written here is the reverse of the order applied, and the
 * mirror being rightmost is exactly the measured `rotate . flip`.
 *
 * Returns an empty string for a frame that needs no transform at all, so an
 * un-turned shape at the origin emits no attribute.
 */
export function frameTransform(frame: Frame): string {
  // Nothing is turned, so the pair of translations about the centre cancels and
  // what is left is where the shape is. Worth the branch: it is the overwhelming
  // majority of shapes on any slide, and it keeps the emitted markup readable.
  if (frame.rot === 0 && !frame.flipH && !frame.flipV) {
    return frame.x === 0 && frame.y === 0 ? '' : `translate(${trim(frame.x)} ${trim(frame.y)})`;
  }

  const parts: string[] = [];
  const centreX = frame.x + frame.cx / 2;
  const centreY = frame.y + frame.cy / 2;
  if (centreX !== 0 || centreY !== 0) {
    parts.push(`translate(${trim(centreX)} ${trim(centreY)})`);
  }
  if (frame.rot !== 0) parts.push(`rotate(${trim(frame.rot)})`);
  if (frame.flipH || frame.flipV) {
    parts.push(`scale(${frame.flipH ? '-1' : '1'} ${frame.flipV ? '-1' : '1'})`);
  }
  if (frame.cx !== 0 || frame.cy !== 0) {
    parts.push(`translate(${trim(-frame.cx / 2)} ${trim(-frame.cy / 2)})`);
  }
  return parts.join(' ');
}

/** The union of two boxes. */
export function unionBox(a: Box | null, b: Box | null): Box | null {
  if (a === null) return b;
  if (b === null) return a;
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  return {
    x,
    y,
    cx: Math.max(a.x + a.cx, b.x + b.cx) - x,
    cy: Math.max(a.y + a.cy, b.y + b.cy) - y,
  };
}

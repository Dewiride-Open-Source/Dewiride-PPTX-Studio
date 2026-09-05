import { FULL_CIRCLE } from './formula.js';

/**
 * The guides every shape has without declaring them.
 *
 * A preset formula freely references `w`, `hc`, `ss`, `wd2`, `ssd6`, `3cd4` and
 * a couple of dozen others that appear in no `gdLst` anywhere. They are seeded
 * before evaluation begins, and they depend on nothing but the shape's width
 * and height.
 *
 * ## The table is derived from the naming rule, not from a copied list
 *
 * The names are systematic. `wdN` is the width divided by N, `hdN` the height,
 * `ssdN` the shorter side; `cdN` is a full circle divided by N and `McdN` is M
 * of those. Writing the divisors out and applying the rule is both shorter than
 * an enumeration and harder to get individually wrong.
 *
 * It is also the only version that survives contact with the data. The 187
 * presets reference `wd3`, `wd12`, `hd3`, `hd10`, `ssd16` and `cd3`, and
 * several published enumerations of "the built-in guides" omit some of those -
 * an enumeration transcribed from one of them would be missing a guide that
 * `rtTriangle` and `curvedLeftArrow` need, and the failure would be a shape
 * that throws rather than one that looks wrong.
 *
 * `builtins.test.ts` closes the loop from the other side: it holds the list of
 * every name the 187 presets actually reference, read off the committed
 * buckets, and asserts this table covers all of it. Two sources, and a
 * disagreement between them is a finding rather than something to paper over.
 *
 * ## What is NOT here
 *
 * Adjust values. `adj`, `adj1`..`adj8`, `hf` and `vf` are declared by each
 * preset's own `avLst` with its own defaults, so they are not built in even
 * though they look like they might be. See `evaluate.ts`.
 */

/**
 * Divisors that appear as `wdN`. The presets need 2, 3, 4, 5, 6, 8, 10, 12 and
 * 32; the width and height families are given the same divisors because there
 * is no reason for them to differ and a shape is free to be authored either way
 * round.
 */
const WIDTH_DIVISORS = [2, 3, 4, 5, 6, 8, 10, 12, 32] as const;
const HEIGHT_DIVISORS = [2, 3, 4, 5, 6, 8, 10, 12, 32] as const;

/** Divisors that appear as `ssdN`. The presets need 2, 6, 8, 16 and 32. */
const SHORT_SIDE_DIVISORS = [2, 4, 6, 8, 16, 32] as const;

/** `cdN` is one full turn divided by N. The presets need 2, 3 and 4. */
const CIRCLE_DIVISORS = [2, 3, 4, 8] as const;

/**
 * `McdN` is M turns of a full circle divided by N - so `3cd4` is three quarter
 * turns, 270 degrees, 16200000 angle units. The presets need only `3cd4`; the
 * eighths are the rest of the family and are documented alongside it.
 */
const CIRCLE_MULTIPLES = [
  [3, 4],
  [3, 8],
  [5, 8],
  [7, 8],
] as const;

/**
 * Every built-in guide for a shape of this size.
 *
 * Angles are in 60000ths of a degree and are independent of the size, but they
 * are returned in the same map because a formula operand cannot tell an angle
 * from a coordinate - only the operator consuming it can.
 *
 * A fresh map per call, deliberately: the values depend on `w` and `h`, and
 * caching them would mean keying a cache on a pair of floats to save forty
 * divisions.
 */
export function builtinGuides(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  const guides = new Map<string, number>([
    // The edges. `l` and `t` are zero because preset geometry is expressed in
    // the shape's own space, not the slide's - the offset onto the slide is
    // `a:xfrm/@off` and is applied much later, in 2.10.
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],

    ['w', w],
    ['h', h],
    ['hc', w / 2],
    ['vc', h / 2],

    // The shorter side - what keeps a rounded corner round on a shape that is
    // not square. 84 of the 187 presets depend on it.
    ['ss', ss],
    // The longer side. No preset references it; it is here because custom
    // geometry may, and a missing built-in is a throw rather than a wrong
    // number.
    ['ls', Math.max(w, h)],
  ]);

  for (const n of WIDTH_DIVISORS) guides.set(`wd${String(n)}`, w / n);
  for (const n of HEIGHT_DIVISORS) guides.set(`hd${String(n)}`, h / n);
  for (const n of SHORT_SIDE_DIVISORS) guides.set(`ssd${String(n)}`, ss / n);
  for (const n of CIRCLE_DIVISORS) guides.set(`cd${String(n)}`, FULL_CIRCLE / n);
  for (const [m, n] of CIRCLE_MULTIPLES) {
    guides.set(`${String(m)}cd${String(n)}`, (m * FULL_CIRCLE) / n);
  }

  return guides;
}

/**
 * The names of every built-in guide, sorted by code unit.
 *
 * Computed from a unit-sized shape rather than written out again, so it cannot
 * drift from `builtinGuides`. Sorted the same way as `presetNames` and for the
 * same reason - `localeCompare` is collation and depends on the engine's ICU
 * build.
 */
export function builtinGuideNames(): readonly string[] {
  return [...builtinGuides(1, 1).keys()].sort();
}

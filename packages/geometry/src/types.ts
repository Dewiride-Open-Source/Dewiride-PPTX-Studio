/**
 * The shape of a preset geometry, before any of it is evaluated.
 *
 * Everything here is a *string*, deliberately. A preset definition is a program
 * whose operands are guide names, built-in names and literals all mixed
 * together - the multiply-divide operator takes the shape width, a user guide
 * and the literal 200000 in one formula - and nothing can tell them apart until
 * the evaluator has a shape size and an `avLst` to resolve against. Parsing
 * that 200000 to a number here would mean deciding, at build time, which
 * operands are numbers, which is 2.2's job and not the transcoder's.
 *
 * (Spelled out rather than quoted, because the operator's own two characters
 * end a block comment.)
 *
 * The one exception is `PresetPath`'s own attributes, which are fixed values
 * with no formula behind them.
 */

/** Which generated bucket a preset lives in. A packaging fact, not a semantic one. */
export type PresetBucketName = 'basic' | 'arrows' | 'callouts' | 'flowchart' | 'misc' | 'stars';

/** One `a:gd`: a named value and the formula that produces it. */
export interface PresetGuide {
  readonly name: string;
  /**
   * The `fmla` attribute, split on whitespace: the operator first, then its
   * operands. Not a fixed arity, because the source data does not have one -
   * see `OVER_LONG_FORMULAS` in the generated buckets.
   */
  readonly fmla: readonly string[];
}

/** A point in guide space. Both coordinates are formula operands, not numbers. */
export interface PresetPoint {
  readonly x: string;
  readonly y: string;
}

/**
 * A point after evaluation, in shape coordinates.
 *
 * The same units the caller passed as the shape size - this package never
 * decides whether that is EMU, points or pixels.
 */
export interface Point {
  readonly x: number;
  readonly y: number;
}

/**
 * One drawing command.
 *
 * `arcTo` carries no destination point: the arc starts at the current point and
 * its end is computed from the radii and the two angles. That asymmetry is in
 * ECMA, and 2.3 is where it stops being convenient - the angles are not
 * parametric, so `stAng` has to be unskewed before an SVG ellipse can be asked
 * to draw the same curve.
 */
export type PresetCommand =
  | { readonly kind: 'moveTo'; readonly to: PresetPoint }
  | { readonly kind: 'lnTo'; readonly to: PresetPoint }
  | { readonly kind: 'quadBezTo'; readonly c1: PresetPoint; readonly to: PresetPoint }
  | {
      readonly kind: 'cubicBezTo';
      readonly c1: PresetPoint;
      readonly c2: PresetPoint;
      readonly to: PresetPoint;
    }
  | {
      readonly kind: 'arcTo';
      readonly wR: string;
      readonly hR: string;
      readonly stAng: string;
      readonly swAng: string;
    }
  | { readonly kind: 'close' };

/** `ST_PathFillMode`. `norm` is the default and is never written in the source. */
export type PresetPathFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

export interface PresetPath {
  /**
   * Path-space width and height. **Zero means no scaling**, not a degenerate
   * path: an unscaled path is already in shape coordinates. Reading these as a
   * divisor without the zero check is a divide-by-zero on 270 of the 320 paths.
   */
  readonly w: number;
  readonly h: number;
  readonly fill: PresetPathFill;
  readonly stroke: boolean;
  readonly extrusionOk: boolean;
  readonly commands: readonly PresetCommand[];
}

/**
 * An adjust handle, in one of its two coordinate systems.
 *
 * Either axis may be absent: a handle that only moves horizontally has a
 * `gdRefX` and no `gdRefY`, and a null there means "this axis is not
 * adjustable", which is different from "adjustable through a range of zero".
 */
export interface PresetAdjustHandleXY {
  readonly kind: 'xy';
  readonly gdRefX: string | null;
  readonly minX: string | null;
  readonly maxX: string | null;
  readonly gdRefY: string | null;
  readonly minY: string | null;
  readonly maxY: string | null;
  readonly pos: PresetPoint;
}

export interface PresetAdjustHandlePolar {
  readonly kind: 'polar';
  readonly gdRefR: string | null;
  readonly minR: string | null;
  readonly maxR: string | null;
  readonly gdRefAng: string | null;
  readonly minAng: string | null;
  readonly maxAng: string | null;
  readonly pos: PresetPoint;
}

export type PresetAdjustHandle = PresetAdjustHandleXY | PresetAdjustHandlePolar;

/** A connection site: where a connector may attach, and the angle it leaves at. */
export interface PresetConnectionSite {
  readonly ang: string;
  readonly pos: PresetPoint;
}

/** The text rectangle. Four formula operands, not a box. */
export interface PresetTextRect {
  readonly l: string;
  readonly t: string;
  readonly r: string;
  readonly b: string;
}

/**
 * A geometry definition: one of the 187 presets, or an `a:custGeom` read from a
 * slide.
 *
 * One type, because they are one thing. `a:custGeom` carries `avLst`, `gdLst`,
 * `ahLst`, `cxnLst`, `rect` and `pathLst` - the same six children, with the
 * same meanings - so a custom geometry is a preset definition written inline in
 * the shape instead of being looked up by name. The only difference is the
 * name: `a:prstGeom` has one and `a:custGeom` does not.
 *
 * Everything downstream takes this type rather than `PresetShape`, so the two
 * cannot drift into separate code paths. `path.ts` has the conformance test
 * that holds them to it.
 */
export interface Geometry {
  /** The `ST_ShapeType` value from `a:prstGeom/@prst`, or `null` for an `a:custGeom`. */
  readonly name: string | null;
  /** Adjust values, with their defaults. Overridden by name from the shape's own `avLst`. */
  readonly avLst: readonly PresetGuide[];
  /** Computed guides, in dependency order - the source data is already ordered. */
  readonly gdLst: readonly PresetGuide[];
  readonly ahLst: readonly PresetAdjustHandle[];
  readonly cxnLst: readonly PresetConnectionSite[];
  /**
   * The text rectangle, or `null` when there is none - five presets, and any
   * `custGeom` that omits `a:rect`. A caller that defaults a missing one to the
   * shape bounds is choosing a policy, which is why this is null rather than
   * pre-filled.
   */
  readonly rect: PresetTextRect | null;
  readonly pathLst: readonly PresetPath[];
}

/**
 * A preset: a `Geometry` that is guaranteed to have a name.
 *
 * The five presets with no text rectangle are `line`, `lineInv`, `chartPlus`,
 * `chartStar` and `chartX`.
 */
export interface PresetShape extends Geometry {
  readonly name: string;
}

/**
 * Build a `Geometry` from the parts an `a:custGeom` actually declared.
 *
 * Every child of `a:custGeom` except `a:pathLst` is optional in the schema, and
 * most real ones carry only `avLst`, `gdLst` and `pathLst`. This fills the rest
 * in as empty so a reader does not have to, and so nothing downstream has to
 * ask whether a list is missing or merely empty.
 */
export function custGeom(parts: Partial<Geometry> = {}): Geometry {
  return {
    name: parts.name ?? null,
    avLst: parts.avLst ?? [],
    gdLst: parts.gdLst ?? [],
    ahLst: parts.ahLst ?? [],
    cxnLst: parts.cxnLst ?? [],
    rect: parts.rect ?? null,
    pathLst: parts.pathLst ?? [],
  };
}

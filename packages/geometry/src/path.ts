import { arcEnd, arcSegments, type ArcParameters, type ArcSegment } from './arc.js';
import {
  evaluateGuides,
  resolveOperand,
  resolvePoint,
  type EvaluateOptions,
  type ShapeSize,
} from './evaluate.js';
import type { FormulaSite } from './formula.js';
import type { Geometry, Point, PresetPath, PresetPathFill } from './types.js';

/**
 * From a geometry definition to something that can be drawn.
 *
 * This closes the chain 2.1 to 2.4: 2.1 read the shapes, 2.2 gave them numbers,
 * 2.3 resolved the arcs, and here a `pathLst` becomes a list of segments and an
 * SVG `d` string.
 *
 * ## One `a:path`, one `<path>`
 *
 * The 187 presets hold 320 paths and 63 of them declare more than one. That
 * alone would not force separate output - the paths could be concatenated. What
 * forces it is that **60 of the 187 have paths that disagree about fill or
 * stroke**: 96 paths are `fill="none"` and 95 of those are stroked, which is
 * how `smileyFace` draws a filled head and an unfilled mouth, and how the 17
 * action buttons draw a filled plate with an unfilled glyph on it. A renderer
 * that merges a shape's paths into one `d` cannot render any of those 60.
 *
 * So one `ResolvedPath` per `a:path`, each carrying its own `fill` and
 * `stroke`. What those mean in paint is 2.6 to 2.8; they pass through here
 * untouched.
 *
 * ## Path space, and the zero that is not a divisor
 *
 * `a:path/@w` and `@h` declare the coordinate system the path is drawn in, and
 * a **zero means there is no such system** - the coordinates are already in
 * shape space. 270 of the 320 paths are unscaled, so reading these as a divisor
 * without the zero check is a division by zero on five paths in six.
 *
 * The other 50 use nine different spaces: 1x1, 2x2, 5x5, 6x6, 8x8, 10x10,
 * 20x20, 21600x21600 and 43200x43200. `flowChartCollate` really is drawn with
 * the coordinates 0, 1 and 2 in a two-unit square.
 *
 * The axes are independent - `@w` of zero with a non-zero `@h` scales y and not
 * x - which is what the two attributes say and is not reachable from the
 * presets, where all 50 scaled paths declare both and every one is square.
 *
 * ## What a guide means inside a scaled path
 *
 * Nothing in the corpus answers this, and the measurement is worth stating
 * precisely: across all 50 scaled paths **every operand in an x, y, wR or hR
 * position is an integer literal**. The only names appearing anywhere in them
 * are `cd2`, `cd4` and `3cd4`, and only in `stAng` and `swAng` positions, where
 * they are angles and belong to no coordinate system at all.
 *
 * So this is a decision rather than a finding: a guide resolves to a number,
 * and a number in a path declaring `@w` is in path space like every other
 * number there. Resolving an operand and choosing a coordinate system are
 * separate steps, and the path is what says which system it is in.
 *
 * ## An arc is resolved in path space, and `cloud` proves it
 *
 * Scaling an ellipse arc by different factors on the two axes gives another
 * ellipse arc, so two orders are available: resolve the arc in path space and
 * scale the result, or scale the radii first and resolve in shape space. They
 * agree on a circle, at cardinal angles, and whenever the shape is square,
 * which between them covers every preset but one.
 *
 * `cloud` is the exception and it settles the question. Its outline is a ring
 * of eleven arcs at angles like -190.5 degrees, chained end to end in a square
 * 43200 path space. The ring is meant to close, and closes to within 0.32% -
 * the residue of the integer angles it was authored with. Resolved in path
 * space that residue is an affine image of a fixed figure and stays where it
 * is. Resolved in shape space the ring tears open:
 *
 * ```text
 *   size          path space   shape space
 *   100 x 100       0.32%        0.32%      agree - square
 *   200 x 100       0.18%        5.46%      31x worse
 *   400 x 100       0.11%        3.89%      34x worse
 *   100 x 800       0.31%        4.95%      16x worse
 * ```
 *
 * A 5% gap in the outline of a cloud is a visible tear. So the arc is resolved
 * against the path's own coordinate system and the figure is scaled afterwards,
 * which is the same order everything else in the path gets. That is the point:
 * a path is a drawing in its own space, and the space is stretched, not the
 * drawing rules.
 *
 * ## `close` closes a subpath, not the path
 *
 * 35 presets have a `close` that is not the last command of its path and carry
 * on drawing afterwards; 51 paths have no `close` at all; one path reaches
 * eleven subpaths. SVG's `Z` behaves the same way - the command after it starts
 * a new subpath at the closed one's first point - so `close` maps straight
 * through.
 */

/** Decimal places kept in an emitted `d` string. */
export const DEFAULT_PRECISION = 4;

/** One segment of a resolved path, in shape coordinates and ready to emit. */
export type PathSegment =
  | { readonly kind: 'move'; readonly to: Point }
  | { readonly kind: 'line'; readonly to: Point }
  | { readonly kind: 'quad'; readonly c1: Point; readonly to: Point }
  | { readonly kind: 'cubic'; readonly c1: Point; readonly c2: Point; readonly to: Point }
  | ({ readonly kind: 'arc' } & ArcSegment)
  | { readonly kind: 'close' };

export interface ResolvedPath {
  /** `ST_PathFillMode`, passed through. What it means in paint is 2.6 to 2.8. */
  readonly fill: PresetPathFill;
  readonly stroke: boolean;
  readonly extrusionOk: boolean;
  /** The factors path space was scaled by. Both 1 when the path declares none. */
  readonly scaleX: number;
  readonly scaleY: number;
  readonly segments: readonly PathSegment[];
  /**
   * SVG path data. **Empty when `finite` is false**, because a `d` holding a
   * `NaN` is not merely wrong, it is invalid: a browser drops the whole path
   * and says nothing. The numbers stay in `segments` for a caller that wants to
   * see what happened.
   */
  readonly d: string;
  /**
   * False when any operand, scale factor or emitted coordinate was not finite.
   *
   * A real case rather than a corrupt-file one. At zero width 11 presets reach
   * a non-finite number inside a path, at zero height 16, and 19 when both
   * dimensions collapse - see the note on division in `evaluate.ts`.
   */
  readonly finite: boolean;
}

/** The text rectangle, resolved. Four numbers in shape coordinates. */
export interface ResolvedRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/** A connection site, resolved. `ang` stays in 60000ths of a degree. */
export interface ResolvedConnectionSite {
  readonly ang: number;
  readonly pos: Point;
}

export interface ResolvedGeometry {
  /** The preset's name, or `null` for an `a:custGeom`, which has none. */
  readonly name: string | null;
  /** Built-ins, adjust values and computed guides, as `evaluateGuides` returns them. */
  readonly guides: ReadonlyMap<string, number>;
  readonly paths: readonly ResolvedPath[];
  /** `null` for the five presets with no `a:rect`. Defaulting one is a policy. */
  readonly textRect: ResolvedRect | null;
  readonly connectionSites: readonly ResolvedConnectionSite[];
}

export interface ResolvePathOptions {
  /** Decimal places in the `d` string. Default `DEFAULT_PRECISION`. */
  readonly precision?: number;
  /** The shape name, used only so a failed operand can say where it came from. */
  readonly name?: string | null;
}

export type ResolveGeometryOptions = EvaluateOptions & ResolvePathOptions;

/**
 * The factors that carry a path's own coordinate system into shape space.
 *
 * `1` on an axis the path does not scale, which is the case for 270 of the 320
 * preset paths and for every `custGeom` that omits the attribute.
 */
export function pathScale(path: PresetPath, size: ShapeSize): { x: number; y: number } {
  return {
    x: path.w === 0 ? 1 : size.w / path.w,
    y: path.h === 0 ? 1 : size.h / path.h,
  };
}

function format(value: number, precision: number): string {
  const fixed = value.toFixed(precision);
  const trimmed = fixed.includes('.') ? fixed.replace(/0+$/, '').replace(/\.$/, '') : fixed;
  return trimmed === '-0' ? '0' : trimmed;
}

/**
 * Segments to an SVG `d` string.
 *
 * Rounded to `precision` decimals with trailing zeros dropped, so two of them
 * can be compared for equality - which is what the `custGeom` conformance test
 * does across all 187 presets.
 */
export function pathData(
  segments: readonly PathSegment[],
  precision: number = DEFAULT_PRECISION,
): string {
  const n = (value: number): string => format(value, precision);
  let out = '';
  for (const segment of segments) {
    switch (segment.kind) {
      case 'move':
        out += `M${n(segment.to.x)} ${n(segment.to.y)}`;
        break;
      case 'line':
        out += `L${n(segment.to.x)} ${n(segment.to.y)}`;
        break;
      case 'quad':
        out += `Q${n(segment.c1.x)} ${n(segment.c1.y)} ${n(segment.to.x)} ${n(segment.to.y)}`;
        break;
      case 'cubic':
        out +=
          `C${n(segment.c1.x)} ${n(segment.c1.y)} ` +
          `${n(segment.c2.x)} ${n(segment.c2.y)} ` +
          `${n(segment.to.x)} ${n(segment.to.y)}`;
        break;
      case 'arc':
        out +=
          `A${n(segment.rx)} ${n(segment.ry)} 0 ` +
          `${segment.largeArc ? 1 : 0} ${segment.sweep ? 1 : 0} ` +
          `${n(segment.to.x)} ${n(segment.to.y)}`;
        break;
      case 'close':
        out += 'Z';
        break;
    }
  }
  return out;
}

/**
 * One arc piece, carried from path space into shape space.
 *
 * The radii scale per axis and the flags survive, because an axis-aligned scale
 * maps an axis-aligned ellipse arc onto another one over the same parameter
 * range. Radii are taken absolute, which SVG requires.
 */
function scaleArc(
  piece: ArcSegment,
  scale: { readonly x: number; readonly y: number },
  reflected: boolean,
): { kind: 'arc' } & ArcSegment {
  return {
    kind: 'arc',
    from: { x: piece.from.x * scale.x, y: piece.from.y * scale.y },
    to: { x: piece.to.x * scale.x, y: piece.to.y * scale.y },
    rx: piece.rx * Math.abs(scale.x),
    ry: piece.ry * Math.abs(scale.y),
    largeArc: piece.largeArc,
    sweep: reflected ? !piece.sweep : piece.sweep,
  };
}

/**
 * Resolve one `a:path` against a set of guides and a shape size.
 *
 * The pen walks in the path's own coordinate system and each result is mapped
 * into shape space as it is produced - see the note on `cloud` above for why
 * that order is not interchangeable with the other one.
 */
export function resolvePath(
  path: PresetPath,
  guides: ReadonlyMap<string, number>,
  size: ShapeSize,
  options: ResolvePathOptions = {},
): ResolvedPath {
  const site: FormulaSite = { preset: options.name ?? null, guide: null };
  const scale = pathScale(path, size);
  let finite = Number.isFinite(scale.x) && Number.isFinite(scale.y);

  const at = (token: string): number => {
    const value = resolveOperand(token, guides, site);
    if (!Number.isFinite(value)) finite = false;
    return value;
  };
  const point = (x: string, y: string): Point => ({ x: at(x), y: at(y) });
  const map = (p: Point): Point => {
    const out = { x: p.x * scale.x, y: p.y * scale.y };
    if (!Number.isFinite(out.x) || !Number.isFinite(out.y)) finite = false;
    return out;
  };

  // A reflection on exactly one axis reverses which way round an arc goes. Not
  // reachable from a slide - `ext/@cx` is a positive coordinate and mirroring
  // is `flipH` on the transform rather than a negative size - but the flag
  // costs one comparison and being quietly wrong here would be hard to find.
  const reflected = scale.x * scale.y < 0;

  const segments: PathSegment[] = [];
  let cursor: Point = { x: 0, y: 0 };
  let subpathStart: Point = cursor;

  for (const command of path.commands) {
    switch (command.kind) {
      case 'moveTo':
        cursor = point(command.to.x, command.to.y);
        subpathStart = cursor;
        segments.push({ kind: 'move', to: map(cursor) });
        break;
      case 'lnTo':
        cursor = point(command.to.x, command.to.y);
        segments.push({ kind: 'line', to: map(cursor) });
        break;
      case 'quadBezTo': {
        const c1 = point(command.c1.x, command.c1.y);
        cursor = point(command.to.x, command.to.y);
        segments.push({ kind: 'quad', c1: map(c1), to: map(cursor) });
        break;
      }
      case 'cubicBezTo': {
        const c1 = point(command.c1.x, command.c1.y);
        const c2 = point(command.c2.x, command.c2.y);
        cursor = point(command.to.x, command.to.y);
        segments.push({ kind: 'cubic', c1: map(c1), c2: map(c2), to: map(cursor) });
        break;
      }
      case 'arcTo': {
        const arc: ArcParameters = {
          wR: at(command.wR),
          hR: at(command.hR),
          stAng: at(command.stAng),
          swAng: at(command.swAng),
        };
        for (const piece of arcSegments(cursor, arc)) {
          segments.push(scaleArc(piece, scale, reflected));
        }
        cursor = arcEnd(cursor, arc);
        if (!Number.isFinite(cursor.x) || !Number.isFinite(cursor.y)) finite = false;
        break;
      }
      case 'close':
        cursor = subpathStart;
        segments.push({ kind: 'close' });
        break;
    }
  }

  return {
    fill: path.fill,
    stroke: path.stroke,
    extrusionOk: path.extrusionOk,
    scaleX: scale.x,
    scaleY: scale.y,
    segments,
    d: finite ? pathData(segments, options.precision ?? DEFAULT_PRECISION) : '',
    finite,
  };
}

/**
 * Evaluate a geometry and resolve everything a renderer needs from it.
 *
 * The same function for a preset and for an `a:custGeom`, because they are the
 * same thing: a `custGeom` is a preset definition written inline in the slide
 * instead of being looked up by name. The only difference the type carries is
 * that a `custGeom` has no `name`.
 */
export function resolveGeometry(
  geometry: Geometry,
  size: ShapeSize,
  options: ResolveGeometryOptions = {},
): ResolvedGeometry {
  const guides = evaluateGuides(geometry, size, options);
  const site: FormulaSite = { preset: geometry.name, guide: null };
  const pathOptions: ResolvePathOptions = {
    name: geometry.name,
    ...(options.precision === undefined ? {} : { precision: options.precision }),
  };

  return {
    name: geometry.name,
    guides,
    paths: geometry.pathLst.map((path) => resolvePath(path, guides, size, pathOptions)),
    textRect:
      geometry.rect === null
        ? null
        : {
            l: resolveOperand(geometry.rect.l, guides, site),
            t: resolveOperand(geometry.rect.t, guides, site),
            r: resolveOperand(geometry.rect.r, guides, site),
            b: resolveOperand(geometry.rect.b, guides, site),
          },
    connectionSites: geometry.cxnLst.map((cxn) => ({
      ang: resolveOperand(cxn.ang, guides, site),
      pos: resolvePoint(cxn.pos, guides, site),
    })),
  };
}

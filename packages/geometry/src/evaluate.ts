import { builtinGuides } from './builtins.js';
import { GeometryError } from './errors.js';
import { applyOperator, siteLabel, type FormulaSite } from './formula.js';
import type { Geometry, Point, PresetGuide, PresetPoint } from './types.js';

/**
 * Turning a geometry definition into numbers.
 *
 * This is the step 2.1 deliberately did not take. A `Geometry` is a program
 * whose operands are guide names, built-in names and literals mixed together;
 * given a size and a set of adjust values, this module runs it and produces
 * every guide the shape's points can refer to.
 *
 * It stops there. No path is emitted, no arc is unskewed, nothing is drawn -
 * those are `arc.ts` and `path.ts`. What comes out is a map of names to numbers.
 *
 * It takes a `Geometry` rather than a `PresetShape` because an `a:custGeom` is
 * the same program written inline in the slide; the only thing it lacks is a
 * name, and the only use this module has for a name is saying which shape an
 * operand failed in.
 */

/** The size of the shape being evaluated, in whatever unit the caller is using. */
export interface ShapeSize {
  readonly w: number;
  readonly h: number;
}

/**
 * A deck's own adjust values.
 *
 * Either the parsed `a:avLst` - which is a list of guides and may in principle
 * carry any formula, not just `val` - or the ergonomic form for a caller that
 * has already reduced them to numbers.
 */
export type AdjustOverrides = readonly PresetGuide[] | Readonly<Record<string, number>>;

export interface EvaluateOptions {
  /** Overrides from `a:prstGeom/a:avLst`, merged over the preset's defaults by NAME. */
  readonly adjust?: AdjustOverrides;
}

function toGuides(adjust: AdjustOverrides | undefined): Map<string, PresetGuide> {
  const map = new Map<string, PresetGuide>();
  if (adjust === undefined) return map;

  if (Array.isArray(adjust)) {
    for (const gd of adjust as readonly PresetGuide[]) map.set(gd.name, gd);
    return map;
  }

  for (const [name, value] of Object.entries(adjust as Readonly<Record<string, number>>)) {
    map.set(name, { name, fmla: ['val', String(value)] });
  }
  return map;
}

/** Only an integer. `ST_AdjCoordinate` is a guide name or an integer, and nothing else. */
const LITERAL = /^[-+]?\d+$/;

/**
 * Resolve one operand: a guide already in scope, or an integer literal.
 *
 * The two cases cannot be told apart by looking, which is the whole reason 2.1
 * kept every operand as a string. A name wins over a literal reading because a
 * name cannot be an integer and an integer cannot be a name, so there is no
 * ambiguity to resolve - only an order to state.
 */
export function resolveOperand(
  token: string,
  guides: ReadonlyMap<string, number>,
  site: FormulaSite | null = null,
): number {
  const known = guides.get(token);
  if (known !== undefined) return known;

  if (LITERAL.test(token)) return Number(token);

  throw new GeometryError(
    'FMLA_OPERAND',
    `${siteLabel(site)}operand "${token}" is neither an integer nor a guide in scope`,
    site?.preset ?? null,
  );
}

/** Resolve both coordinates of a point. Used by the text rectangle, connection sites and paths. */
export function resolvePoint(
  point: PresetPoint,
  guides: ReadonlyMap<string, number>,
  site: FormulaSite | null = null,
): Point {
  return {
    x: resolveOperand(point.x, guides, site),
    y: resolveOperand(point.y, guides, site),
  };
}

function runGuide(
  gd: PresetGuide,
  guides: ReadonlyMap<string, number>,
  preset: string | null,
): number {
  const site: FormulaSite = { preset, guide: gd.name };
  const op = gd.fmla[0];
  if (op === undefined) {
    throw new GeometryError(
      'FMLA_OPERATOR',
      `${preset === null ? '' : `${preset}.`}${gd.name}: empty formula`,
      preset,
    );
  }
  const args = gd.fmla.slice(1).map((token) => resolveOperand(token, guides, site));
  return applyOperator(op, args, site);
}

/**
 * Evaluate every guide of a preset at a given size.
 *
 * Returns the built-in guides, the adjust values and the computed guides in one
 * map, because an operand elsewhere in the shape - a connection site, the text
 * rectangle, a path coordinate - may name any of the three and cannot tell them
 * apart.
 *
 * ## Order, and why one forward pass is enough
 *
 * Built-ins, then `avLst`, then `gdLst` in document order, each formula seeing
 * only what was defined before it.
 *
 * That a single pass suffices is not an assumption. Across all 3622 computed
 * guides in the 187 presets there are **zero** forward references and zero self
 * references - every operand names something already defined by the time its
 * formula runs. Measured against the committed buckets, and re-measured by
 * `evaluate.test.ts` so that a future POI which reorders a `gdLst` fails here
 * rather than producing a shape built on an undefined guide.
 *
 * ## Adjust values merge by NAME
 *
 * Not by position and not by index. Eleven distinct adjust names appear across
 * the presets - `adj`, `adj1` through `adj8`, and `hf` and `vf` on the polygons
 * and stars - so a positional merge would quietly assign a deck's first adjust
 * to `hf` on a pentagon.
 *
 * An override naming an adjust the preset does not declare is still evaluated
 * and kept, because a deck that writes one is saying something and dropping it
 * silently is worse than carrying a value nothing reads. The exception is a
 * name that collides with a built-in: there the built-in wins, because letting
 * a deck redefine `w` would turn the shape's own width into a suggestion.
 *
 * ## Non-finite results are returned, not thrown
 *
 * 212 of the 1204 division formulas in the presets divide by a guide rather
 * than by a literal, and 36 shapes divide by `ss`. A shape with zero height is
 * completely ordinary - a flat line, a shape being dragged through zero - so
 * `x / 0` is reachable at legal sizes rather than being a corrupt-file case.
 *
 * Throwing there would mean a shape that vanishes from a slide mid-drag and
 * takes an error path with it. So the IEEE result is kept, and
 * `nonFiniteGuides` lets a caller ask. What a renderer should do about it is
 * 2.10's decision, made with a renderer in front of it; this module's job is to
 * report faithfully rather than to pick for it.
 */
export function evaluateGuides(
  shape: Geometry,
  size: ShapeSize,
  options: EvaluateOptions = {},
): Map<string, number> {
  const guides = builtinGuides(size.w, size.h);
  const overrides = toGuides(options.adjust);

  for (const gd of shape.avLst) {
    guides.set(gd.name, runGuide(overrides.get(gd.name) ?? gd, guides, shape.name));
  }

  for (const [name, gd] of overrides) {
    if (!guides.has(name)) guides.set(name, runGuide(gd, guides, shape.name));
  }

  for (const gd of shape.gdLst) {
    guides.set(gd.name, runGuide(gd, guides, shape.name));
  }

  return guides;
}

/**
 * The names of any guides that evaluated to something not finite, sorted.
 *
 * Empty for every preset at any size with a non-zero width and height. See the
 * note on division above for when it is not.
 */
export function nonFiniteGuides(guides: ReadonlyMap<string, number>): string[] {
  const bad: string[] = [];
  for (const [name, value] of guides) if (!Number.isFinite(value)) bad.push(name);
  return bad.sort();
}

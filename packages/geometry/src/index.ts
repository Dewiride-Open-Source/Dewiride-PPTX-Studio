/**
 * DrawingML geometry.
 *
 * Sub-phase 2.1 landed the preset definitions: 187 shapes, transcoded from
 * Apache POI's `presetShapeDefinitions.xml` and readable without evaluating
 * anything. 2.2 adds the evaluator - the seventeen `fmla` operators, the
 * built-in guides, and the merge of a deck's adjust values over a preset's
 * defaults - so a shape and a size now produce numbers.
 *
 * 2.3 adds the arcs: `a:arcTo` names no centre and no destination, its angles
 * are rays rather than ellipse parameters, and a negative swing genuinely runs
 * backwards. `arcGeometry` resolves one against the point a path has reached.
 *
 * 2.4 emits. `resolveGeometry` turns a definition and a size into one
 * `ResolvedPath` per `a:path` - segments, an SVG `d` string, and the fill and
 * stroke each path carries - plus the text rectangle and the connection sites.
 * It takes a `Geometry`, which is the preset type and the `a:custGeom` type
 * both, because they are the same thing written two ways.
 *
 * It stops at geometry. **Nothing is painted** - a `d` string is not a fill, a
 * stroke, a gradient or an effect, and those are 2.6 to 2.8.
 */
export { GeometryError, type GeometryErrorCode } from './errors.js';
export { PresetBucket, decodeShape } from './decode.js';
export { builtinGuideNames, builtinGuides } from './builtins.js';
export {
  ARC_MAX_SWEEP,
  arcEnd,
  arcGeometry,
  arcPointAt,
  arcSegments,
  clampSweep,
  unskewAngle,
  type ArcGeometry,
  type ArcParameters,
  type ArcSegment,
} from './arc.js';
export {
  ANGLE_UNITS_PER_DEGREE,
  FMLA_ARITY,
  FULL_CIRCLE,
  angleToRadians,
  applyOperator,
  isFmlaOperator,
  radiansToAngle,
  type FmlaOperator,
  type FormulaSite,
} from './formula.js';
export {
  evaluateGuides,
  nonFiniteGuides,
  resolveOperand,
  resolvePoint,
  type AdjustOverrides,
  type EvaluateOptions,
  type ShapeSize,
} from './evaluate.js';
export {
  DEFAULT_PRECISION,
  pathData,
  pathScale,
  resolveGeometry,
  resolvePath,
  type PathSegment,
  type ResolveGeometryOptions,
  type ResolvePathOptions,
  type ResolvedConnectionSite,
  type ResolvedGeometry,
  type ResolvedPath,
  type ResolvedRect,
} from './path.js';
export {
  BUCKETS,
  BUCKET_MEMBERS,
  COLLAPSED_WHITESPACE,
  OVER_LONG_FORMULAS,
  PRESET_SOURCE,
  PRESET_TOTALS,
  getPreset,
  presetNames,
} from './presets/index.js';
export { custGeom } from './types.js';
export type {
  Geometry,
  Point,
  PresetAdjustHandle,
  PresetAdjustHandlePolar,
  PresetAdjustHandleXY,
  PresetBucketName,
  PresetCommand,
  PresetConnectionSite,
  PresetGuide,
  PresetPath,
  PresetPathFill,
  PresetPoint,
  PresetShape,
  PresetTextRect,
} from './types.js';

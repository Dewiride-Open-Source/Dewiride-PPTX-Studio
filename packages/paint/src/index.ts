/**
 * `@pptx-studio/paint` - DrawingML colour and fills.
 *
 * Sub-phase 2.6: the six colour bases, the twenty-eight transforms in document
 * order, and resolution against a theme and a colour map. Sub-phase 2.7: the
 * gradient - which stops win, what curve joins them, and where the ramp runs -
 * and the 54 preset pattern tiles. Sub-phase 2.8: the stroke - the eleven dash
 * arrays and what a cap does to them, the five compound forms, the joins and the
 * miter limit Office actually uses, the six arrowheads - and the effects, whose
 * four blur radii turn out to be one operator with one constant. `a:blipFill`
 * has a place in the `Fill` union and nothing behind it.
 *
 * Everything this package decides was measured against Microsoft PowerPoint
 * rather than argued from the standard. The fixtures are
 * `corpus/ground-truth/color-transforms.json` (0.7-C, the transforms),
 * `corpus/ground-truth/color-bases.json` (2.6, the bases, alpha, the colour map
 * and the percentage grammar), `corpus/ground-truth/fills.json` (2.7, gradients
 * and patterns) and `corpus/ground-truth/lines.json` (2.8, strokes and
 * effects); the reasoning is in `docs/adr/phase-0-foundation/0007-ground-truth.md`,
 * `docs/adr/phase-2-geometry-and-paint/0021-colour.md`, `docs/adr/phase-2-geometry-and-paint/0022-fills.md` and
 * `docs/adr/phase-2-geometry-and-paint/0023-lines.md`.
 */

export { PaintError, type PaintErrorCode } from './errors.js';

export {
  clamp01,
  fromLinear,
  hslToRgb,
  luma709,
  parseHex,
  rgbToHsl,
  toByte,
  toHex,
  toLinear,
  wrapHue,
  type Channel,
  type Hsl,
} from './colors/transfer.js';

export {
  COLOR_TRANSFORM_OPS,
  MAPPED_COLOR_NAMES,
  SCHEME_SLOTS,
  transformShape,
  type AngleTransformOp,
  type ClrMap,
  type ClrScheme,
  type Color,
  type ColorTransform,
  type ColorTransformOp,
  type FlagTransformOp,
  type MappedColorName,
  type PercentTransformOp,
  type Rgba,
  type SchemeColorName,
  type SchemeSlot,
} from './types.js';

export { applyTransforms } from './apply.js';

export {
  mapSchemeName,
  resolveColor,
  toCss,
  toHexColor,
  type ColorContext,
} from './colors/resolve.js';

export { parseAngle, parsePercentage, parseSrgbValue } from './parse.js';

export { PRESET_COLORS, isPresetColorName } from './colors/preset-colors.js';

export { SYSTEM_COLORS, isSystemColorName } from './colors/sys-colors.js';

export {
  WHOLE_RECT,
  type BlipFill,
  type BlipEffect,
  type BlipStretch,
  type BlipTile,
  type TileAlign,
  type Fill,
  type GradientFill,
  type GradientShade,
  type GradientStop,
  type GroupFill,
  type LinearShade,
  type NoFill,
  type PathShade,
  type PatternFill,
  type RelativeRect,
  type SolidFill,
  type TileFlipMode,
} from './fills/fill.js';

export {
  RAMP_BLEND_GAMMA,
  TWO_STOP_RAMP,
  fromRampSpace,
  toRampSpace,
  twoStopWeight,
} from './fills/gradient-ramp.js';

export {
  gradientColorAt,
  insetRect,
  twoColorRamp,
  linearGradientVector,
  pathGeometry,
  pathPositionAt,
  resolveGradientStops,
  sortStops,
  svgStops,
  type GradientVector,
  type PathGeometry,
  type ResolvedStop,
  type TwoColorRamp,
  type SvgStop,
} from './fills/gradient.js';
export {
  blipPlacement,
  brightOffset,
  contrastScale,
  BLIP_LUMA,
  type BlipPlacement,
  type ImageSize,
  type Rect,
  type StretchPlacement,
  type TilePlacement,
} from './fills/blip.js';

export {
  ANTIALIASED_PATTERNS,
  PATTERN_PIXEL_EMU,
  PATTERN_TILES,
  PATTERN_TILE_PIXELS,
  PRESET_PATTERN_NAMES,
  isPresetPatternName,
  type PatternTile,
} from './fills/pattern-tiles.js';

export { patternInk, patternPixel, resolvePattern, type ResolvedPattern } from './fills/pattern.js';

export {
  COMPOUND_NAMES,
  COMPOUND_RUNS,
  COMPOUND_RUNS_60,
  isCompoundName,
} from './lines/compound-table.js';

export {
  PRESET_DASHES,
  PRESET_DASH_NAMES,
  isPresetDashName,
  type DashSegment,
} from './lines/dash-table.js';

export {
  MARKERS,
  MARKER_SIZE,
  MARKER_SIZE_NAMES,
  MARKER_TYPE_NAMES,
  isMarkerSizeName,
  isMarkerTypeName,
  type Marker,
  type MarkerAnchor,
} from './lines/marker-table.js';

export {
  DEFAULT_LINE_CAP,
  DEFAULT_LINE_JOIN,
  DEFAULT_LINE_WIDTH,
  DEFAULT_MITER_LIMIT,
  EMU_PER_POINT,
  compoundRails,
  dashArray,
  dashSegments,
  markerGeometry,
  markerOvershoot,
  resolveLine,
  svgStroke,
  type Line,
  type LineCap,
  type LineDash,
  type LineEnd,
  type LineEndSize,
  type LineJoin,
  type LineJoinKind,
  type MarkerGeometry,
  type PenAlignment,
  type ResolvedLine,
  type StrokeRail,
  type SvgStroke,
} from './lines/line.js';

export {
  BLUR_RADIUS_TO_SIGMA,
  DEFAULT_SHADOW_ALIGN,
  EDGE_GROWTH_PER_RADIUS,
  GLOW_RADIUS_SCALE,
  anchorPoint,
  blurSigma,
  effectFilter,
  effectMargin,
  glowGeometry,
  shadowBox,
  shadowMatrix,
  shadowOffset,
  softEdgeGeometry,
  type Blur,
  type Box,
  type Effect,
  type FilterGraph,
  type FilterPrimitive,
  type Glow,
  type InnerShadow,
  type Matrix,
  type OuterShadow,
  type PresetShadow,
  type RectAlignment,
  type Reflection,
  type ShadowGeometry,
  type SoftEdge,
} from './effect.js';

/**
 * `@pptx-studio/paint` - DrawingML colour and fills.
 *
 * Sub-phase 2.6: the six colour bases, the twenty-eight transforms in document
 * order, and resolution against a theme and a colour map. Sub-phase 2.7: the
 * gradient - which stops win, what curve joins them, and where the ramp runs -
 * and the 54 preset pattern tiles. Dashes, arrowheads and effects are 2.8 and
 * are not here yet; `a:blipFill` has a place in the `Fill` union and nothing
 * behind it.
 *
 * Everything this package decides was measured against Microsoft PowerPoint
 * rather than argued from the standard. The fixtures are
 * `corpus/ground-truth/color-transforms.json` (0.7-C, the transforms),
 * `corpus/ground-truth/color-bases.json` (2.6, the bases, alpha, the colour map
 * and the percentage grammar) and `corpus/ground-truth/fills.json` (2.7,
 * gradients and patterns); the reasoning is in `docs/adr/0007-ground-truth.md`,
 * `docs/adr/0021-colour.md` and `docs/adr/0022-fills.md`.
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
} from './transfer.js';

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

export { mapSchemeName, resolveColor, toCss, toHexColor, type ColorContext } from './resolve.js';

export { parseAngle, parsePercentage, parseSrgbValue } from './parse.js';

export { PRESET_COLORS, isPresetColorName } from './preset-colors.js';

export { SYSTEM_COLORS, isSystemColorName } from './sys-colors.js';

export {
  WHOLE_RECT,
  type BlipFill,
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
} from './fill.js';

export {
  RAMP_BLEND_GAMMA,
  TWO_STOP_RAMP,
  fromRampSpace,
  toRampSpace,
  twoStopWeight,
} from './gradient-ramp.js';

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
} from './gradient.js';

export {
  ANTIALIASED_PATTERNS,
  PATTERN_PIXEL_EMU,
  PATTERN_TILES,
  PATTERN_TILE_PIXELS,
  PRESET_PATTERN_NAMES,
  isPresetPatternName,
  type PatternTile,
} from './pattern-tiles.js';

export { patternInk, patternPixel, resolvePattern, type ResolvedPattern } from './pattern.js';

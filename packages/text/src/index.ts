/**
 * `@pptx-studio/text` - text measurement and the line model.
 *
 * Sub-phase 3.2: what a line box is, how far apart two baselines sit, and how
 * wide a run is. Every rule was measured against Microsoft PowerPoint in
 * experiment T2 - 2091 probes across five packages, read back through the
 * object model - rather than argued from the standard, because on the questions
 * that decide a line count the standard either says nothing or says something
 * PowerPoint does not do:
 *
 * - The line box is a **font-independent 1.2 x the font size**. Thirteen faces,
 *   seven sizes, one number. ECMA-376 does not mention it.
 * - `a:lnSpc` **percent multiplies the line box**, not the font size, so 150% at
 *   18pt is 32.4pt and not 27pt; `a:spcPts` is used raw with no 1.2 anywhere.
 * - The percentage is **rounded half up to a whole percent** before use, so
 *   `val="106667"` is 107%.
 * - `a:rPr/@kern` is a **minimum font size, inclusive**, and `kern="0"` means
 *   never - not the boolean every implementation reaches for first.
 * - `a:rPr/@spc` is an **absolute length in points**, applied after every
 *   character including the last, which is what CSS `letter-spacing` does.
 *
 * The fixture is `corpus/ground-truth/text-metrics.json`; the reasoning and the
 * refuted alternatives are in `docs/adr/phase-3-text/0028-measurement-and-the-line-model.md`.
 *
 * Sub-phase 3.3 added line breaking - which turned out not to be a tailoring of
 * UAX#14 but a much smaller rule that UAX#14 is a superset of - and 3.4 added
 * autofit, whose fit test forced the one question 3.2 left open: what the
 * *last* line of a block measures. It is `0.75 x advance + b x size` with `b` a
 * property of the typeface, floored at the advance while the advance is inside
 * `C x size`. See `autofit.ts`, `docs/adr/phase-3-text/0029-line-breaking.md` and
 * `docs/adr/phase-3-text/0030-autofit.md`.
 */

export { TextError, type TextErrorCode } from './errors.js';

export {
  NATURAL_LINE_FACTOR,
  blockHeight,
  lineAdvance,
  lineTop,
  naturalLineHeight,
  quantisePercent,
  type LineSpacing,
} from './lines/line-model.js';

export {
  BUILT_IN_KINSOKU,
  HANGING_PUNCTUATION,
  kinsokuInForce,
  type KinsokuSets,
} from './lines/kinsoku.js';

export {
  breakOpportunities,
  isEastAsian,
  toCodePoints,
  wrapText,
  type BreakOptions,
  type BreakTailoring,
  type LineBox,
  type RangeMeasurer,
  type WrapInput,
} from './lines/break.js';

export {
  APPROXIMATE_FACE_METRICS,
  AUTOFIT_LADDER,
  MEASURED_FACE_METRICS,
  SHAPE_AUTOFIT_SLACK,
  autofitAdvance,
  effectiveFontSize,
  faceLineMetrics,
  fitAutofit,
  hasFaceLineMetrics,
  lastLineHeight,
  requiredShapeHeight,
  storedScale,
  textBodyHeight,
  type AutofitRequest,
  type AutofitResult,
  type AutofitScale,
  type FaceLineMetrics,
  type ParagraphBox,
} from './lines/autofit.js';

export {
  createCanvasMeasurer,
  cssFamily,
  cssFont,
  cssLetterSpacing,
  kerningEnabled,
  type Advance,
  type RunFont,
  type TextMeasurer,
} from './runs/measure.js';

export {
  ALPHABETS,
  AUTONUMBER_SCHEMES,
  START_AT_MAX,
  START_AT_MIN,
  UNMEASURED_SCHEMES,
  autonumberRule,
  autonumberTypeface,
  formatAutonumber,
  isAutonumberScheme,
  type Alphabet,
  type AutonumberRule,
  type AutonumberScheme,
} from './bullets/autonumber.js';

export {
  BLIP_BULLET_HEIGHT,
  SYMBOL_TYPEFACES,
  blipBulletWidth,
  bulletLayout,
  drawableBullet,
  isStartAtInRange,
  isSymbolTypeface,
  numberParagraphs,
  symbolBulletChar,
  type BulletColor,
  type BulletContext,
  type BulletFont,
  type BulletKind,
  type BulletLayout,
  type BulletSize,
  type DrawableBullet,
  type NumberedParagraph,
  type ResolvedBullet,
} from './bullets/bullets.js';

export {
  RESERVED_FIELD_TYPES,
  datePatternFor,
  formatDateTimeField,
  isReservedFieldType,
  renderDatePattern,
  renderField,
  type FieldRequest,
  type RenderedField,
  type ReservedFieldType,
} from './fields/fields.js';

export { FIELD_DATE_PATTERNS } from './fields/field-formats.gen.js';

export {
  scriptSlotOf,
  splitScriptRuns,
  typefaceFor,
  type RunTypefaces,
  type ScriptRun,
  type ScriptSlot,
} from './runs/script-runs.js';

export {
  DEFAULT_INSETS,
  ELLIPSIS,
  MAX_COLUMNS,
  VERTICAL_AXES,
  anchorFraction,
  blockOrigin,
  columnBox,
  contentBox,
  drawnLines,
  emptyParagraphHeight,
  frameAxes,
  offsetAlong,
  type Anchor,
  type Box,
  type DrawnLines,
  type Edge,
  type FrameAxes,
  type Insets,
  type OverflowLine,
  type VertOverflow,
  type VerticalText,
} from './frames/frame.js';

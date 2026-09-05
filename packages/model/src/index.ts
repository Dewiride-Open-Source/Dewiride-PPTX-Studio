/**
 * `@pptx-studio/model` - the sheets, the chain between them, and the resolver.
 *
 * A `.pptx` renders correctly or it does not, and almost all of the difference
 * is here rather than in any renderer. A slide placeholder that PowerPoint wrote
 * has no geometry, no fill, no line and no style of its own; every one of those
 * comes down a chain of two hops whose rules are asymmetric, undocumented, and
 * different from what every implementation this project has read assumes.
 *
 * Sub-phase 2.9 builds four things:
 *
 * - **The chain**, from each part's own `.rels` and from nothing else. Not from
 *   `p:sldLayoutIdLst`, not from a part's number in its folder, and above all
 *   not from the `theme` relationship on `ppt/presentation.xml`, which always
 *   names `theme1.xml` and is right on exactly the decks that have one master.
 * - **The matcher.** Slide to layout on `@idx` alone; layout to master on the
 *   folded type, taking the first. Both measured, both surprising, and the
 *   evidence is in `corpus/ground-truth/sheets.json`.
 * - **The resolver**, which returns a value *and where it came from*, so the
 *   inspector, per-property Reset, layout scoring and theme invalidation all
 *   read the same function the renderer reads.
 * - **The style matrix and `phClr`**, without which every themed shape is black,
 *   and the background, which shares the matrix and the 1000-offset with it.
 *
 * Nothing here is resolved eagerly and nothing is flattened. `undefined` means
 * the file did not say, and that is the fact Change Layout, provenance and
 * theme swapping are all built on.
 *
 * The measurements are `corpus/ground-truth/sheets.json`; the reasoning is
 * `docs/adr/phase-2-geometry-and-paint/0024-model-parse-and-resolve.md`.
 */

export { ModelError, MODEL_ERROR_CODES, isModelError, type ModelErrorCode } from './errors.js';

export {
  MASTER_PLACEHOLDER_TYPES,
  PLACEHOLDER_TYPES,
  type Background,
  type ColorMapOverride,
  type FontCollection,
  type FontScheme,
  type FormatScheme,
  type NormalPlaceholder,
  type Origin,
  type Placeholder,
  type PlaceholderSize,
  type PlaceholderType,
  type Resolved,
  type Shape,
  type ShapeKind,
  type ShapeStyle,
  type Sheet,
  type SheetKind,
  type StyleRef,
  type FontRef,
  type Theme,
  type Xfrm,
} from './types.js';

export {
  COLOR_TRANSFORM_OPS,
  parseColorChild,
  parseColorElement,
  parseEffects,
  parseFill,
  parseFillElement,
  parseLine,
  parseLineElement,
} from './parse/paint.js';

export { parseGeometry, type ShapeGeometry } from './parse/geometry.js';

export { parseClrMap, parseSheet, parseTheme } from './parse/sheet.js';

export {
  DEFAULT_PLACEHOLDER_IDX,
  DEFAULT_PLACEHOLDER_TYPE,
  ORPHAN_RECT,
  inheritanceChain,
  masterPlaceholderType,
  matchInLayout,
  matchInMaster,
  normalizePlaceholder,
  placeholders,
  type ChainLink,
} from './resolve/placeholder.js';

export {
  colorContextOf,
  colorMapOf,
  masterOf,
  resolve,
  resolveXfrm,
  schemeOf,
  sheetChain,
  themeOf,
} from './resolve/resolve.js';

export {
  STYLE_MATRIX_OFFSET,
  phClrOf,
  resolveAppearance,
  resolveOnSheet,
  resolveSolidFill,
  styleMatrixEffects,
  styleMatrixFill,
  styleMatrixLine,
  styleMatrixTarget,
  type ResolvedAppearance,
  type StyleMatrixTarget,
} from './style.js';

export {
  backgroundSheet,
  resolveBackground,
  resolveBackgroundColor,
  type ResolvedBackground,
} from './resolve/background.js';

export { loadDocument, type Document, type DocumentProblem } from './document.js';

export {
  LEVELS,
  themeFontRef,
  type Caps,
  type FontAlign,
  type ListStyle,
  type Paragraph,
  type ParaProps,
  type RunProps,
  type Spacing,
  type Strike,
  type TextAlign,
  type TextBody,
  type TextBreak,
  type TextContent,
  type BulletAutoNum,
  type BulletColor,
  type BulletFont,
  type BulletKind,
  type BulletSize,
  type TextField,
  type TextRun,
  type TextStyleBucket,
  type TextStyles,
  type ThemeFontRef,
  type Typeface,
  type Underline,
} from './text.js';

export {
  parseDefaultTextStyle,
  parseListStyle,
  parseParaProps,
  parseRunProps,
  parseTextBody,
  parseTextBodyChild,
  parseTextStyles,
} from './parse/text.js';

export { BUILTIN_TEXT_STYLES, TEXT_FLOOR, type BuiltinLevel } from './builtin-text-styles.js';

export {
  bucketOf,
  floorOf,
  resolveBulletAutoNum,
  resolveBulletBlip,
  resolveBulletChar,
  resolveBulletColor,
  resolveBulletFont,
  resolveBulletKind,
  resolveBulletSize,
  resolveIndent,
  resolveMarginLeft,
  resolveParagraph,
  resolveRun,
  resolveSize,
  textLevels,
  type TextContext,
} from './resolve/text.js';

export {
  resolveAnchor,
  resolveAnchorCtr,
  resolveBody,
  resolveColumns,
  resolveInsets,
  resolveVertical,
  resolveWrap,
  type ResolvedInsets,
} from './resolve/body.js';

export { parseBodyProps, parseBodyPropsChild } from './parse/body.js';

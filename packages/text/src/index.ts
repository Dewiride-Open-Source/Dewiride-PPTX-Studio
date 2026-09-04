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
 * The fixture is `corpus/ground-truth/text-metrics.json`; the reasoning, the
 * refuted alternatives, and the one question left open - what the *last* line of
 * a block measures, which needs a per-face term no browser reports - are in
 * `docs/adr/0028-measurement-and-the-line-model.md`.
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
} from './line-model.js';

export {
  createCanvasMeasurer,
  cssFamily,
  cssFont,
  cssLetterSpacing,
  kerningEnabled,
  type Advance,
  type RunFont,
  type TextMeasurer,
} from './measure.js';

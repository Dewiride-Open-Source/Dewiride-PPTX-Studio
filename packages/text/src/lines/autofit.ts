/**
 * Autofit: what `a:normAutofit` and `a:spAutoFit` mean, measured in experiment
 * T4 by making PowerPoint compute an autofit and reading the numbers it wrote
 * into its own file.
 *
 * The plan for this sub-phase asked for two modes, and both survive contact
 * with the measurement:
 *
 * - **View.** A stored `@fontScale` or `@lnSpcReduction` is applied exactly as
 *   written. PowerPoint does not recompute one on open - a deck stating 25% on
 *   two words in a huge box renders at 25%, and one stating 100% on twelve
 *   lines in a two-line box renders overflowing - and it does not snap a value
 *   to the ladder either, so a 45% written by another producer stays 45%.
 * - **Edit.** A recomputation walks a **15-rung ladder** and takes the first
 *   rung that fits. It never invents a scale between rungs.
 *
 * Everything below the ladder is arithmetic, and every constant in it was
 * measured rather than read out of ECMA-376, which says nothing about any of
 * this beyond naming the two attributes.
 *
 * ## The parts of a text height
 *
 * ```
 * effectiveSize = fontScale is 100% ? size : round(size x fontScale)
 * advance       = a:lnSpc in points ? those points
 *                                   : max(1%, round(lnSpc% - reduction)) x 1.2 x effectiveSize
 * lastLine      = advance <= C x effectiveSize ? max(advance, fitted) : fitted
 *                 where fitted = 0.75 x advance + b x effectiveSize
 * height        = (lines - 1) x advance + lastLine + the paragraph spacing terms
 * ```
 *
 * Four of those are worth saying out loud because the obvious reading is wrong:
 *
 * 1. **`round(size x fontScale)` is to a whole point**, so 18pt at 92.5% lays
 *    out at 17pt and not 16.65pt - but only when a scale applies. A `@fontScale`
 *    of exactly 100% leaves 10.5pt alone.
 * 2. **The reduction is subtracted from the line-spacing percentage**, not
 *    multiplied into the advance. Identical at single spacing; 4% apart at 150%.
 * 3. **Exact-point line spacing ignores the reduction entirely.** 30pt stays
 *    30pt at every rung, so autofit cannot shrink such a paragraph at all.
 * 4. **The last line is not the advance.** `b` and `C` are properties of the
 *    typeface - see `FaceLineMetrics` - and this is the term sub-phase 3.2
 *    left open in `blockHeight`.
 *
 * The fixture is `corpus/ground-truth/autofit.json`; the reasoning, the refuted
 * alternatives and the questions still open are in `docs/adr/phase-3-text/0030-autofit.md`.
 */

import { TextError } from '../errors.js';
import { NATURAL_LINE_FACTOR, quantisePercent, type LineSpacing } from './line-model.js';

/**
 * One rung: the pair of attributes PowerPoint writes into `a:normAutofit`.
 *
 * Both are in thousandths of a percent, the units the attributes use, so a
 * value read from a file goes straight in and one written out goes straight
 * back - there is no scaling step in which a factor of a thousand can be lost.
 */
export interface AutofitScale {
  /** `a:normAutofit/@fontScale`. */
  readonly fontScale: number;
  /** `a:normAutofit/@lnSpcReduction`. */
  readonly lnSpcReduction: number;
}

/**
 * The ladder, in the order PowerPoint tries it.
 *
 * Font scale descending, and within one font scale the smaller reduction
 * first. That ordering is measured, not assumed: at 32pt and at 11pt the
 * rounding of `size x scale` makes `(85000, 10000)` a *taller* rung than
 * `(92500, 20000)`, and PowerPoint takes the second anyway - so "the first rung
 * in this order that fits" is right and "the tallest rung that fits" is wrong,
 * scoring 3066 of 3262.
 *
 * The font scales step by 7.5 points of a percent from 100 down to 25, and the
 * reduction is 20% for all of them except the top three, which carry a smaller
 * one as well. There is no rung below 25%: text that does not fit there is left
 * overflowing.
 */
export const AUTOFIT_LADDER: readonly AutofitScale[] = [
  { fontScale: 100000, lnSpcReduction: 0 },
  { fontScale: 100000, lnSpcReduction: 10000 },
  { fontScale: 92500, lnSpcReduction: 0 },
  { fontScale: 92500, lnSpcReduction: 10000 },
  { fontScale: 92500, lnSpcReduction: 20000 },
  { fontScale: 85000, lnSpcReduction: 10000 },
  { fontScale: 85000, lnSpcReduction: 20000 },
  { fontScale: 77500, lnSpcReduction: 20000 },
  { fontScale: 70000, lnSpcReduction: 20000 },
  { fontScale: 62500, lnSpcReduction: 20000 },
  { fontScale: 55000, lnSpcReduction: 20000 },
  { fontScale: 47500, lnSpcReduction: 20000 },
  { fontScale: 40000, lnSpcReduction: 20000 },
  { fontScale: 32500, lnSpcReduction: 20000 },
  { fontScale: 25000, lnSpcReduction: 20000 },
];

/**
 * The two numbers the last-line rule needs that belong to the typeface rather
 * than to the format.
 *
 * They are not derivable from anything a browser reports. Chromium's
 * `TextMetrics.fontBoundingBoxDescent` was the obvious candidate and it is
 * refuted: Arial and Verdana report the same 0.2119 and 0.2100 while their
 * measured `lastLineCoefficient` differs by a tenth of that again, and
 * `fontBoundingBoxAscent + fontBoundingBoxDescent` puts Verdana above Courier
 * New where `floorRatio` puts it below. So they are measured per face and
 * carried as data, and a caller holding a face this experiment did not measure
 * has to say what it wants used.
 */
export interface FaceLineMetrics {
  /**
   * `b`: the per-size term in `lastLine = 0.75 x advance + b x size`.
   *
   * Spans 0.2053 on Tahoma to 0.3000 on Courier New. Fitted across seven sizes
   * and two line spacings a face, to a worst residual of 0.0043pt.
   */
  readonly lastLineCoefficient: number;
  /**
   * `C`: the multiple of the effective size up to which the last line takes the
   * whole advance instead of the fitted value.
   *
   * Indistinguishable from the 1.2 line box on five of the six faces measured,
   * and 1.2698 on Courier New - which is the only reason it is a field rather
   * than the constant 1.2.
   */
  readonly floorRatio: number;
}

/**
 * The six faces T4 measured, by the name a run states in `a:latin/@typeface`.
 *
 * Generated data would be better, but six rows are not worth a codegen step and
 * `packages/text/src/lines/autofit.test.ts` asserts every one of them against the
 * fixture - so a value edited here without re-measuring fails the suite.
 */
export const MEASURED_FACE_METRICS: Readonly<Record<string, FaceLineMetrics>> = {
  Arial: { lastLineCoefficient: 0.227618991, floorRatio: 1.2 },
  'Times New Roman': { lastLineCoefficient: 0.234424476, floorRatio: 1.2 },
  Verdana: { lastLineCoefficient: 0.207355431, floorRatio: 1.2 },
  'Courier New': { lastLineCoefficient: 0.300000058, floorRatio: 1.269842 },
  Georgia: { lastLineCoefficient: 0.231489897, floorRatio: 1.2 },
  Tahoma: { lastLineCoefficient: 0.205318787, floorRatio: 1.2 },
};

/**
 * What to use for a typeface nobody measured.
 *
 * Arial's coefficient with the 1.2 line box, which five of the six faces share.
 * It is exported under a name that says what it is so that reaching for it is a
 * decision at the call site with a diagnostic beside it, rather than a default
 * this module applies behind the caller's back. The error a caller gets from
 * `faceLineMetrics` is the other half of the same rule.
 *
 * The cost of being wrong is bounded and small: the coefficient spans 0.0947
 * across six faces, and it multiplies the size of exactly one line in a block.
 * At 18pt that is at most 1.7pt on a block of any length, and it is zero
 * whenever the line spacing is within the range where the floor applies -
 * which includes every paragraph at single spacing.
 */
export const APPROXIMATE_FACE_METRICS: FaceLineMetrics = {
  lastLineCoefficient: 0.227618991,
  floorRatio: NATURAL_LINE_FACTOR,
};

/**
 * The measured metrics for a typeface, or a typed failure naming it.
 *
 * Throwing rather than returning a default is the point: a renderer that hits
 * this has a face it cannot lay out exactly, and that is a diagnostic for the
 * fonts panel in 3.7, not something to paper over here.
 */
export function faceLineMetrics(typeface: string): FaceLineMetrics {
  const metrics = MEASURED_FACE_METRICS[typeface];
  if (metrics === undefined) {
    throw new TextError(
      'TEXT_FACE_METRICS',
      `no measured line metrics for ${typeface}; pass APPROXIMATE_FACE_METRICS deliberately if that is what you want`,
      typeface,
    );
  }
  return metrics;
}

/** Whether this typeface was one of the six measured. */
export function hasFaceLineMetrics(typeface: string): boolean {
  return Object.hasOwn(MEASURED_FACE_METRICS, typeface);
}

/* -------------------------------------------------------------------------- */
/* the arithmetic                                                             */
/* -------------------------------------------------------------------------- */

/** Thousandths of a percent, which is what both attributes hold. */
const FULL_SCALE = 100000;

/**
 * The smallest line spacing PowerPoint will lay out, as a fraction of the line
 * box.
 *
 * A 15% spacing reduced by 20% is -5%, and PowerPoint draws it at one per cent
 * of the line box rather than at zero or at something negative. Measured at
 * five sizes: 8pt gives 0.095pt a line where 1% of the box is 0.096, and 40pt
 * gives 0.480 where 1% is 0.480.
 */
const MIN_SPACING = 0.01;

function checkScale(value: number, what: string): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new TextError(
      'TEXT_AUTOFIT_SCALE',
      `${what} ${String(value)} is not a non-negative percentage`,
      String(value),
    );
  }
}

/**
 * The size a run is actually drawn at, in points.
 *
 * `round(nominal x fontScale)` to a **whole point**, halves up - measured on
 * seventy-seven direct advance readings, against no rounding (35/77), the
 * nearest half point (41/77), the nearest quarter (38/77) and round-half-to-even
 * (74/77, refuted by five rungs of a 20pt sweep where the product lands exactly
 * on a half).
 *
 * The exception is the one a plausible implementation gets wrong: a `@fontScale`
 * of exactly 100% - or none at all - leaves the size alone, fraction included.
 * A 10.5pt run at the rung that only reduces the line spacing stays 10.5pt, and
 * rounding it to 11 there is a visible half point on every line.
 */
export function effectiveFontSize(sz: number, fontScale = FULL_SCALE): number {
  if (!Number.isFinite(sz) || sz <= 0) {
    throw new TextError('TEXT_SIZE', `font size ${String(sz)} is not a positive size`, String(sz));
  }
  checkScale(fontScale, 'a:normAutofit/@fontScale');
  const points = sz / 100;
  if (fontScale === FULL_SCALE) return points;
  return Math.round((points * fontScale) / FULL_SCALE);
}

/**
 * Baseline-to-baseline distance under an autofit scale, in points.
 *
 * The two halves of this are separately measured and separately surprising:
 *
 * - A percentage `a:lnSpc` has the reduction **subtracted from it**, and the
 *   difference is quantised to a whole percent afterwards. `quantise(p) - r` and
 *   `quantise(p - r)` are the same number until one of them is a fraction of a
 *   percent, and a 149.6% spacing reduced by 12.5% is 137%, not 137.5%. Scored
 *   77/77 against 70/77 for quantising first.
 * - An exact-point `a:lnSpc` **ignores the reduction**. 30pt stays 30pt at every
 *   rung, which is why a paragraph with exact line spacing cannot be autofitted
 *   shorter and why `ln-pts` produced a staircase four points wide.
 */
export function autofitAdvance(
  effectiveSizePt: number,
  lnSpcReduction = 0,
  lnSpc?: LineSpacing,
): number {
  checkScale(lnSpcReduction, 'a:normAutofit/@lnSpcReduction');
  if (lnSpc !== undefined && lnSpc.kind === 'points') {
    if (!Number.isFinite(lnSpc.value) || lnSpc.value < 0) {
      throw new TextError(
        'TEXT_LINE_SPACING',
        `a:spcPts/@val ${String(lnSpc.value)} is not a non-negative length`,
        String(lnSpc.value),
      );
    }
    return lnSpc.value / 100;
  }
  const stated = lnSpc?.value ?? FULL_SCALE;
  checkScale(stated, 'a:spcPct/@val');
  // The difference is clamped before it is quantised rather than after, which
  // is the same answer - a negative difference quantises to a negative and the
  // one-per-cent floor takes over either way - and keeps a reduction larger
  // than the spacing from handing `quantisePercent` a value its contract
  // forbids.
  const multiple = Math.max(MIN_SPACING, quantisePercent(Math.max(0, stated - lnSpcReduction)));
  return multiple * NATURAL_LINE_FACTOR * effectiveSizePt;
}

/**
 * The height of the last line of a block, in points.
 *
 * This is the term `blockHeight` in the line model takes as a parameter and
 * sub-phase 3.2 could not measure, because at single spacing it is exactly the
 * advance and the two cannot be told apart. Under a stated `a:lnSpc` they come
 * apart, and the `last` deck read the difference off six faces at seven sizes:
 * `0.75 x advance + b x size`, with the 0.75 the same for every face to six
 * decimal places and `b` a property of the typeface.
 *
 * The floor is the part that is easy to get wrong. Below `C x size` the last
 * line takes the whole advance instead - which is why a block at single spacing
 * measures `n x advance` and why 3.2's entire table fitted that. Above it the
 * fitted value is all there is: `min(advance, 1.2 x size)` looks like the same
 * rule and is refuted by three probes.
 */
export function lastLineHeight(
  advancePt: number,
  effectiveSizePt: number,
  metrics: FaceLineMetrics,
): number {
  const fitted = 0.75 * advancePt + metrics.lastLineCoefficient * effectiveSizePt;
  // The comparison is against a product of a measured advance and a stated
  // ratio, so a paragraph at exactly single spacing sits on the boundary and
  // binary rounding decides which side. A tolerance far below a printer's dot
  // keeps that from flipping the branch.
  if (advancePt <= metrics.floorRatio * effectiveSizePt + 1e-9) {
    return Math.max(advancePt, fitted);
  }
  return fitted;
}

/** One paragraph, already laid out at some candidate scale. */
export interface ParagraphBox {
  /** How many lines it wrapped into at this scale. At least one. */
  readonly lines: number;
  /** `a:rPr/@sz` in hundredths of a point, before any scale. */
  readonly sz: number;
  /** The metrics of the face that governs the paragraph's last line. */
  readonly metrics: FaceLineMetrics;
  readonly lnSpc?: LineSpacing | undefined;
  readonly spcBef?: LineSpacing | undefined;
  readonly spcAft?: LineSpacing | undefined;
}

/**
 * `a:spcBef` or `a:spcAft` in points.
 *
 * A percentage is a fraction of the **scaled** 1.2 line box, so it shrinks with
 * the font; an exact length is used **raw** and does not. Both halves are
 * measured - 36/36 on the emitted rungs against 30/36 for scaling the exact
 * lengths and 24/36 for reading the percentage as a fraction of the font size
 * rather than of the box.
 */
function paragraphSpace(value: LineSpacing | undefined, effectiveSizePt: number): number {
  if (value === undefined) return 0;
  if (value.kind === 'points') {
    if (!Number.isFinite(value.value) || value.value < 0) {
      throw new TextError(
        'TEXT_LINE_SPACING',
        `a:spcPts/@val ${String(value.value)} is not a non-negative length`,
        String(value.value),
      );
    }
    return value.value / 100;
  }
  checkScale(value.value, 'a:spcPct/@val');
  return (value.value / FULL_SCALE) * NATURAL_LINE_FACTOR * effectiveSizePt;
}

/**
 * The height a laid-out text body occupies, in points.
 *
 * `spcFirstLastPara` decides whether the outer two spacing terms count, and it
 * **defaults false** - the first paragraph's `spcBef` and the last's `spcAft`
 * are dropped. Measured both ways round: on the emitted rungs (36/36 against
 * 28/36) and on the block height PowerPoint reports for a body it is not
 * autofitting at all (18/18).
 *
 * Note what the flag does *not* do. With it on, the block is a `spcBef` taller
 * but the first line still sits exactly on the frame top - the extra space is
 * not above the first line. Where it goes is 3.6's question, about anchoring;
 * here only the height matters.
 */
export function textBodyHeight(
  paragraphs: readonly ParagraphBox[],
  scale: AutofitScale,
  spcFirstLastPara = false,
): number {
  if (paragraphs.length === 0) return 0;
  const count = paragraphs.length;
  let total = 0;
  let finalAdvance = 0;
  let finalLastLine = 0;
  for (const [index, paragraph] of paragraphs.entries()) {
    if (!Number.isInteger(paragraph.lines) || paragraph.lines < 1) {
      throw new TextError(
        'TEXT_LINE_COUNT',
        `paragraph ${String(index)} has ${String(paragraph.lines)} lines`,
        String(paragraph.lines),
      );
    }
    const effective = effectiveFontSize(paragraph.sz, scale.fontScale);
    const advance = autofitAdvance(effective, scale.lnSpcReduction, paragraph.lnSpc);
    total += paragraph.lines * advance;
    if (index === count - 1) {
      finalAdvance = advance;
      finalLastLine = lastLineHeight(advance, effective, paragraph.metrics);
    }
    const first = index === 0;
    const last = index === count - 1;
    if (spcFirstLastPara || !first) total += paragraphSpace(paragraph.spcBef, effective);
    if (spcFirstLastPara || !last) total += paragraphSpace(paragraph.spcAft, effective);
  }
  // One last line for the whole body, not one per paragraph. Measured: a body
  // of eight single-line paragraphs is seven advances and one last line, not
  // eight last lines - the two differ by 8% at 150% spacing and the fixture
  // separates them on hundreds of probes.
  //
  // Every probe held the paragraphs identical, so a body whose *last*
  // paragraph has a different size or line spacing from the others is
  // generalisation rather than measurement. It is the only reading consistent
  // with what was measured, and the ADR records it as open.
  //
  // With `spcFirstLastPara` the last line takes a full advance instead of the
  // trimmed height, measured in T6 over five line spacings.
  return spcFirstLastPara ? total : total - finalAdvance + finalLastLine;
}

/* -------------------------------------------------------------------------- */
/* the two modes                                                              */
/* -------------------------------------------------------------------------- */

/**
 * View mode: the scale a file states, applied as written.
 *
 * There is no snapping and no recomputation. A `@fontScale` of 45000 - which no
 * rung carries - is applied as 45%, and a file stating 100% on text that
 * overflows is left overflowing, because that is what PowerPoint does with a
 * file it did not write. Absent attributes mean 100% and 0%, which is the
 * schema default and also what PowerPoint emits when the text already fits.
 */
export function storedScale(fontScale?: number, lnSpcReduction?: number): AutofitScale {
  const scale: AutofitScale = {
    fontScale: fontScale ?? FULL_SCALE,
    lnSpcReduction: lnSpcReduction ?? 0,
  };
  checkScale(scale.fontScale, 'a:normAutofit/@fontScale');
  checkScale(scale.lnSpcReduction, 'a:normAutofit/@lnSpcReduction');
  return scale;
}

export interface AutofitRequest {
  /**
   * The height the text has to fit in, in points: the shape height less `tIns`
   * and `bIns`.
   *
   * Measured 20/20 against 12/20 for testing against the shape height itself,
   * so the insets are the caller's to subtract and are named here rather than
   * passed as a rectangle, to make forgetting them harder.
   */
  readonly bodyHeightPt: number;
  /** `a:bodyPr/@spcFirstLastPara`, resolved. Absent behaves as false. */
  readonly spcFirstLastPara?: boolean | undefined;
  /**
   * Lay the text out at a candidate scale and report each paragraph.
   *
   * A callback rather than a fixed list because the line count is a consequence
   * of the scale: shrinking the font rewraps the text, and a fit test that used
   * one line count for every rung would be wrong on every paragraph that wraps.
   * Use `effectiveFontSize` to get the size to wrap at.
   */
  readonly layout: (scale: AutofitScale) => readonly ParagraphBox[];
}

export interface AutofitResult {
  readonly scale: AutofitScale;
  /** The paragraphs as `layout` reported them at the chosen scale. */
  readonly paragraphs: readonly ParagraphBox[];
  /** The height those paragraphs occupy, in points. */
  readonly heightPt: number;
  /** True when even the last rung overflows, so the text is drawn overflowing. */
  readonly overflows: boolean;
}

/**
 * Edit mode: recompute the scale the way PowerPoint does.
 *
 * Walks `AUTOFIT_LADDER` and takes the first rung whose laid-out text fits,
 * falling back to the last rung - which is what PowerPoint does, leaving the
 * text overflowing rather than shrinking past 25%.
 *
 * Never emit a scale this did not return. PowerPoint re-snaps anything else to
 * a rung the next time the text is touched, and the text visibly jumps.
 */
export function fitAutofit(request: AutofitRequest): AutofitResult {
  if (!Number.isFinite(request.bodyHeightPt)) {
    throw new TextError(
      'TEXT_AUTOFIT_HEIGHT',
      `body height ${String(request.bodyHeightPt)} is not a usable length`,
      String(request.bodyHeightPt),
    );
  }
  const spcFirstLastPara = request.spcFirstLastPara ?? false;
  let last: AutofitResult | undefined;
  for (const scale of AUTOFIT_LADDER) {
    const paragraphs = request.layout(scale);
    const heightPt = textBodyHeight(paragraphs, scale, spcFirstLastPara);
    // The height is a sum of products of measured lengths and stated ratios, so
    // a rung that fits exactly - and several do, at every whole-point box - can
    // land a few ulps above the box. The tolerance is four orders of magnitude
    // below the quarter point a box height can be stated in.
    if (heightPt <= request.bodyHeightPt + 1e-7) {
      return { scale, paragraphs, heightPt, overflows: false };
    }
    last = { scale, paragraphs, heightPt, overflows: true };
  }
  if (last === undefined) throw new TextError('TEXT_AUTOFIT_SCALE', 'the ladder is empty', null);
  return last;
}

/* -------------------------------------------------------------------------- */
/* the other autofit                                                          */
/* -------------------------------------------------------------------------- */

/**
 * The height `a:spAutoFit` gives a shape, in points.
 *
 * The text height plus both insets, times a constant just under one per cent
 * over. 181 probes across four faces bracket it in
 * `[1.009764816, 1.009766650]`, and it scales the insets as well as the text,
 * so it is not a font metric - it is slack PowerPoint adds so the text it just
 * measured definitely fits. What it *is* remains open; see the ADR.
 *
 * The value here is `517/512`, which sits inside that bracket. A dyadic
 * rational inside a measured range is a better guess at what a program computed
 * than the range's midpoint is, and every probe reproduces from it - but it is
 * a guess about the *form* of a number whose *value* is measured, and the ADR
 * says so.
 *
 * PowerPoint applies the result only when it differs from the current height by
 * more than a couple of points: the boxes from 87 to 90pt were all left alone
 * against a target of 87.244. That hysteresis is deliberately **not**
 * implemented here, because this function answers "how tall should the shape
 * be" and the decision to leave a shape alone belongs to the command that
 * resizes it.
 */
export const SHAPE_AUTOFIT_SLACK = 517 / 512;

export function requiredShapeHeight(
  paragraphs: readonly ParagraphBox[],
  insets: { readonly topPt: number; readonly bottomPt: number },
  scale: AutofitScale = { fontScale: FULL_SCALE, lnSpcReduction: 0 },
  spcFirstLastPara = false,
): number {
  const text = textBodyHeight(paragraphs, scale, spcFirstLastPara);
  if (!Number.isFinite(insets.topPt) || !Number.isFinite(insets.bottomPt)) {
    throw new TextError(
      'TEXT_AUTOFIT_HEIGHT',
      `insets ${String(insets.topPt)}/${String(insets.bottomPt)} are not usable lengths`,
      null,
    );
  }
  return SHAPE_AUTOFIT_SLACK * (text + insets.topPt + insets.bottomPt);
}

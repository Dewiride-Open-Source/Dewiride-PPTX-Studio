/**
 * What a line box is, measured against PowerPoint in experiment T2.
 *
 * Everything here is arithmetic over numbers the file states. No font is
 * consulted, and that is the finding rather than a simplification: the line
 * advance is **font-independent**, and thirteen faces whose vertical metrics
 * disagree violently - Impact, Verdana, Courier New among them - reported the
 * same advance at every one of seven sizes, to a worst spread of 0.00003pt.
 *
 * This is why `line-height: normal` must never reach a rendered slide. A
 * browser resolves it from the font's own `hhea.lineGap` on macOS and
 * `OS/2.usWinAscent` on Windows, so the same markup gives different line counts
 * on different machines. PowerPoint asks the font nothing.
 *
 * The fixture is `corpus/ground-truth/text-metrics.json`; the reasoning, the
 * refuted alternatives and the one thing still open are in
 * `docs/adr/0028-measurement-and-the-line-model.md`.
 */

import { TextError } from './errors.js';

/**
 * `a:lnSpc`, as the file states it.
 *
 * Structurally the same as `@pptx-studio/model`'s `Spacing`, so a resolved
 * paragraph property passes straight in. It is restated here rather than
 * imported because this package is a layer below the document model and must
 * not depend on it.
 */
export interface LineSpacing {
  readonly kind: 'percent' | 'points';
  /**
   * For `percent`, `a:spcPct/@val` in thousandths of a percent - the value
   * `ST_Percentage` parses to, so both the `150000` and the `150%` spellings
   * arrive here as 150000. For `points`, `a:spcPts/@val` in hundredths of a
   * point.
   */
  readonly value: number;
}

/**
 * The line box at single spacing, as a multiple of the font size.
 *
 * Measured, not assumed: 13 faces x 7 sizes, including a fractional 10.5pt,
 * every one of them exactly 1.2. A font-derived line height would have given
 * thirteen different numbers.
 */
export const NATURAL_LINE_FACTOR = 1.2;

/** Hundredths of a point to points, rejecting what `@sz` cannot hold. */
function sizeToPoints(sz: number): number {
  if (!Number.isFinite(sz) || sz <= 0) {
    throw new TextError('TEXT_SIZE', `font size ${String(sz)} is not a positive size`, String(sz));
  }
  return sz / 100;
}

/**
 * `a:spcPct/@val` to a multiple, rounded half up to a whole percent.
 *
 * This is a real quantisation and not a tidy-up. Measured on the advance
 * between two baselines: 99.5% lays out as 100%, 100.4% as 100%, and 100.5% as
 * 101%. Ceiling would have sent 100.2% up and floor would have sent 100.5%
 * down; both are refuted, and so is using the value as given (353/371 against
 * 371/371).
 *
 * It matters outside this table. `@val` is in thousandths of a percent, and
 * producers other than PowerPoint write values like `106667` for "1.07 lines" -
 * which PowerPoint lays out as 107%, not as 106.667%.
 */
export function quantisePercent(value: number): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new TextError(
      'TEXT_LINE_SPACING',
      `a:spcPct/@val ${String(value)} is not a non-negative percentage`,
      String(value),
    );
  }
  return Math.round(value / 1000) / 100;
}

/** The line box at single spacing, in points. */
export function naturalLineHeight(sz: number): number {
  return NATURAL_LINE_FACTOR * sizeToPoints(sz);
}

/**
 * Baseline-to-baseline distance, in points.
 *
 * Two rules, and the asymmetry between them is measured rather than tidy:
 *
 * - **Percent** multiplies the *line box*, not the font size. At 150% on an
 *   18pt run that is 32.4pt, not 27pt - a 20% difference on every line of every
 *   deck that sets line spacing. Scored 463/463 against 66/463 for the reading
 *   that applies the percentage to the size.
 * - **Exact points** is used as given, with no 1.2 anywhere. Scored 463/463
 *   against 397/463 for the reading that scales it too.
 *
 * An absent `a:lnSpc` is single spacing: measured, a file that states none lays
 * out identically to one that states `100000`.
 */
export function lineAdvance(sz: number, lnSpc?: LineSpacing): number {
  if (lnSpc === undefined || lnSpc.kind === 'percent') {
    const multiple = quantisePercent(lnSpc?.value ?? 100000);
    return multiple * naturalLineHeight(sz);
  }
  if (!Number.isFinite(lnSpc.value) || lnSpc.value < 0) {
    throw new TextError(
      'TEXT_LINE_SPACING',
      `a:spcPts/@val ${String(lnSpc.value)} is not a non-negative length`,
      String(lnSpc.value),
    );
  }
  return lnSpc.value / 100;
}

function checkCount(value: number, what: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new TextError(
      'TEXT_LINE_COUNT',
      `${what} ${String(value)} is not a non-negative integer`,
      String(value),
    );
  }
}

/**
 * The top of line `index`, in points from the top of the text.
 *
 * Zero for the first line, and that is measured: the text block starts exactly
 * on the frame top, with no half-leading above it, at every spacing from 50% to
 * 300% and at every size. A layout engine that centres the leading - as CSS
 * does - puts every line of every slide in the wrong place.
 */
export function lineTop(index: number, advance: number): number {
  checkCount(index, 'line index');
  return index * advance;
}

/**
 * The height of a block of `lineCount` lines, in points.
 *
 * `(n - 1) * advance + lastLineHeight`, scored 463/463. Every line box except
 * the last measures exactly the advance; the last one does not, and what it
 * does measure needs a per-face term no browser reports - see
 * `corpus/ground-truth/text-metrics.json` under `lastLine`, and sub-phase 3.6.
 * The caller supplies it rather than this function guessing, because a plausible
 * guess here is exactly the silent wrong answer the repository forbids.
 */
export function blockHeight(lineCount: number, advance: number, lastLineHeight: number): number {
  checkCount(lineCount, 'line count');
  if (lineCount === 0) return 0;
  return (lineCount - 1) * advance + lastLineHeight;
}

/**
 * Typed failures, as everywhere else in this repository.
 *
 * Measurement fails for exactly two kinds of reason: it was handed a number the
 * format cannot produce, or it was asked to measure in an environment that
 * cannot measure. Each gets its own code, because the fix is different - the
 * first is a parser that let something through, the second is a caller running
 * the measurer somewhere there is no canvas.
 *
 * Note what is *not* here. There is no error for text that comes out wider than
 * its shape, for a line spacing of 300%, or for an eight-point run. Those are
 * ordinary results that PowerPoint also produces, and T2 measured what it does
 * with them.
 */
export type TextErrorCode =
  /** A font size that is not a positive, finite number of hundredths of a point. */
  | 'TEXT_SIZE'
  /** An `a:lnSpc` value outside what `ST_TextSpacing` can hold. */
  | 'TEXT_LINE_SPACING'
  /** A line index or count that is not a non-negative integer. */
  | 'TEXT_LINE_COUNT'
  /** A typeface name that cannot be written into a CSS font shorthand. */
  | 'TEXT_FONT_FAMILY'
  /** A wrap width or hyphen advance that is not a usable length. */
  | 'TEXT_BREAK_WIDTH'
  /**
   * No `OffscreenCanvas`, or no 2d context on one.
   *
   * The measurer is built for a Web Worker, where `OffscreenCanvas` is the only
   * canvas there is. A caller that reaches this on the main thread has imported
   * the Worker-side measurer into the document side by mistake, which is worth
   * saying rather than silently falling back to a DOM canvas the plan
   * deliberately does not use.
   */
  | 'TEXT_NO_CANVAS';

export class TextError extends Error {
  readonly code: TextErrorCode;
  /** The offending value, when the failure is about one. */
  readonly token: string | null;

  constructor(code: TextErrorCode, message: string, token: string | null = null) {
    super(message);
    this.name = 'TextError';
    this.code = code;
    this.token = token;
  }
}

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
  /** An `a:normAutofit` scale outside what `ST_TextFontScalePercentOrPercentString` can hold. */
  | 'TEXT_AUTOFIT_SCALE'
  /** A body height or inset that is not a usable length. */
  | 'TEXT_AUTOFIT_HEIGHT'
  /**
   * A typeface whose per-face line metrics were never measured.
   *
   * Two of the numbers the last line of a block needs belong to the typeface
   * and are not derivable from anything a browser reports - see `autofit.ts`.
   * A caller that reaches this either passes `APPROXIMATE_FACE_METRICS` on
   * purpose, with a diagnostic, or supplies its own; what it must not do is
   * get a plausible number from here without having asked for one.
   */
  | 'TEXT_FACE_METRICS'
  /** An `@u` or `@strike` value that is not in its ST_ simple type. */
  | 'TEXT_DECORATION'
  /** An `a:buAutoNum/@type` that is not one of the 41 schemes PowerPoint accepts. */
  | 'TEXT_AUTONUMBER_SCHEME'
  /**
   * One of the two schemes T5 measured but could not read.
   *
   * `arabic1Minus` and `arabic2Minus` are shaped into glyph ids before they are
   * drawn, so the measurement holds indices into a font rather than characters.
   * Distinct from `TEXT_AUTONUMBER_SCHEME` because the scheme is real and the
   * gap is ours - "not a scheme" would send somebody to go and check the file.
   */
  | 'TEXT_AUTONUMBER_UNMEASURED'
  /**
   * A bullet number outside `ST_TextBulletStartAtNum`, which is 1 to 32767.
   *
   * Not pedantry: `startAt="0"`, `"-1"`, `"32768"` and `"65536"` are each a
   * package PowerPoint **repairs** rather than reads, measured one probe to a
   * package. A reader that quietly renders one has accepted a file the format's
   * own author refuses.
   */
  | 'TEXT_AUTONUMBER_VALUE'
  /** A bullet whose kind and content disagree - an `autonum` with no scheme, say. */
  | 'TEXT_BULLET'
  /**
   * A date field whose `@lang` and `@type` were never measured.
   *
   * PowerPoint formats dates with Windows' own per-locale patterns, and
   * `Intl.DateTimeFormat` reproduces them on a quarter of the readings. So the
   * patterns are a measured table, and a cell outside it is a gap the caller
   * has to handle - by showing the cached text, which is what `renderField`
   * does - rather than a licence to invent a plausible date.
   */
  | 'TEXT_FIELD_LOCALE'
  /** A field kind or slide number that cannot be rendered as asked. */
  | 'TEXT_FIELD'
  /** A code point outside Unicode, handed to the script-run splitter. */
  | 'TEXT_SCRIPT_RUN'
  /**
   * An `a:bodyPr` value the text frame cannot lay out.
   *
   * A `@numCol` outside `ST_TextColumnCount`, a `@vert` or `@anchor` that is not
   * one of the enumerated values, or a length that is not finite. Each is a
   * package PowerPoint repairs rather than reads.
   */
  | 'TEXT_FRAME'
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

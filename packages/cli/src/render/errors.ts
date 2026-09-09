/**
 * Typed failures for rendering outside a browser.
 *
 * The browser renderer cannot produce any of these: it is handed a measuring
 * canvas and a font stack, and the platform answers. In Node both of those are
 * ours to find, so both are ours to fail at. ADR 0042.
 */
export type RenderErrorCode =
  /** Bytes that are not an SFNT font, or a table the reader needs is absent. */
  | 'CLI_FONT_UNREADABLE'
  /** A font directory that was named on the command line and does not exist. */
  | 'CLI_FONT_DIR'
  /** No face on this machine can stand in for a typeface the deck names. */
  | 'CLI_NO_FACE'
  /** A run whose resolved size is not a positive number of hundredths of a point. */
  | 'CLI_FONT_SIZE'
  /** A slide index outside the deck. */
  | 'CLI_NO_SLIDE'
  /** A width that is not a positive whole number of pixels. */
  | 'CLI_WIDTH'
  /** An output path that names no directory, or a directory that is a file. */
  | 'CLI_OUTPUT_PATH';

export const RENDER_ERROR_CODES: readonly RenderErrorCode[] = [
  'CLI_FONT_UNREADABLE',
  'CLI_FONT_DIR',
  'CLI_NO_FACE',
  'CLI_FONT_SIZE',
  'CLI_NO_SLIDE',
  'CLI_WIDTH',
  'CLI_OUTPUT_PATH',
];

export class RenderError extends Error {
  override readonly name = 'RenderError';
  readonly code: RenderErrorCode;
  /** The font file, deck path or typeface the failure is about. */
  readonly subject: string;

  constructor(code: RenderErrorCode, message: string, subject: string) {
    super(message);
    this.code = code;
    this.subject = subject;
  }
}

export function isRenderError(error: unknown): error is RenderError {
  return error instanceof RenderError;
}

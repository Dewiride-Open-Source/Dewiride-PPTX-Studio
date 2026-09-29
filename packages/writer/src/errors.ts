/**
 * The writer's own failures: a hook that threw, a collection not shown to be safe, and its one
 * assertion on its own output. Container and rule failures stay `OpcError` and `ValidateError`.
 */

export const WRITER_ERROR_CODES = [
  /** A `prepare` hook threw. The hook's name is in `detail.hook`. */
  'ERR_PREPARE_FAILED',
  /**
   * The relationship graph could not be walked completely, so no part can be
   * shown to be unreferenced. See `collect.ts` - the sweep refuses rather than
   * guesses, because the guess deletes user data.
   */
  'ERR_COLLECTION_UNSAFE',
  /**
   * A part nobody edited came out of the writer with different bytes.
   *
   * The one thing this package promises. It is checked against the archive we
   * actually emitted rather than against our intention, which is why it is an
   * error here and not a finding in the report.
   */
  'ERR_PRESERVATION_BROKEN',
] as const;

export type WriterErrorCode = (typeof WRITER_ERROR_CODES)[number];

export interface WriterErrorDetail {
  /** The prepare hook that threw, for `ERR_PREPARE_FAILED`. */
  readonly hook?: string;
  /** The part the failure is about, when it is about one. */
  readonly part?: string;
  /** Relationship parts that would not parse, for `ERR_COLLECTION_UNSAFE`. */
  readonly blockedBy?: readonly string[];
}

export class WriterError extends Error {
  override readonly name = 'WriterError';
  readonly code: WriterErrorCode;
  readonly detail: WriterErrorDetail;

  constructor(
    code: WriterErrorCode,
    message: string,
    detail: WriterErrorDetail = {},
    options: { cause?: unknown } = {},
  ) {
    super(message, options);
    this.code = code;
    this.detail = detail;
  }
}

export function isWriterError(value: unknown): value is WriterError {
  return value instanceof WriterError;
}

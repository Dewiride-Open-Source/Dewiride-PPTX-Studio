/**
 * The single error type this package throws.
 *
 * Same invariant as `@pptx-studio/opc` and `@pptx-studio/xml`: **every failure
 * path throws a `ValidateError` and nothing else.**
 *
 * There is a second thing worth saying here, because this package is the one
 * where the distinction matters most. A *finding* is not an error. A finding is
 * the normal output of this package - it says a rule was broken, names the part
 * and the element, and is data a caller can render, sort or ignore. An error is
 * thrown only when validation itself could not be carried out, or when a caller
 * asked for bytes and the report says they must not have them.
 *
 * Getting that split wrong is how a validator becomes something people disable.
 * If every broken rule were an exception, the only way to see the second
 * problem would be to fix the first, and a caller with a deck that breaks nine
 * rules would learn about them one export at a time.
 */

/**
 * Machine-readable failure reasons.
 *
 * Deliberately few. Almost everything this package has to say is a finding.
 */
export const VALIDATE_ERROR_CODES = [
  /**
   * The report carries at least one fatal finding this session introduced, and
   * a caller asked for the bytes anyway.
   *
   * The one that stops a repair prompt reaching a user. `detail.report` carries
   * the whole report, so the caller does not have to re-run validation to find
   * out what was wrong.
   */
  'ERR_VALIDATION_FAILED',
  /**
   * Validation could not be carried out at all.
   *
   * Not "the package is invalid" - "the package could not be read far enough to
   * ask the questions". A `[Content_Types].xml` that will not parse is this,
   * because without it no part has a content type and there is nothing to
   * validate against.
   */
  'ERR_UNVALIDATABLE',
  /** A caller named a rule id that does not exist. */
  'ERR_UNKNOWN_RULE',
] as const;

export type ValidateErrorCode = (typeof VALIDATE_ERROR_CODES)[number];

export interface ValidateErrorDetail {
  /** The part the failure is about, if it is about one. */
  readonly part?: string;
  /** The rule id a caller got wrong, for `ERR_UNKNOWN_RULE`. */
  readonly rule?: string;
  /**
   * The full report, for `ERR_VALIDATION_FAILED`.
   *
   * Typed `unknown` rather than `Report` on purpose: `report.ts` imports this
   * module for the error, so typing this field would make the two circular.
   * Callers narrow it with `isReport`.
   */
  readonly report?: unknown;
}

export class ValidateError extends Error {
  readonly code: ValidateErrorCode;
  readonly detail: ValidateErrorDetail;

  constructor(code: ValidateErrorCode, message: string, detail: ValidateErrorDetail = {}) {
    super(message);
    this.name = 'ValidateError';
    this.code = code;
    this.detail = detail;
  }
}

export function isValidateError(value: unknown): value is ValidateError {
  return value instanceof ValidateError;
}

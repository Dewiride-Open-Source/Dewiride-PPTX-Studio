/**
 * The single error type this package throws.
 *
 * Same invariant as `@pptx-studio/opc`: **every failure path throws an
 * `XmlError` and nothing else.** Not a `RangeError` from a slice, not a
 * `TypeError` from a decoder. Untrusted bytes reach this package directly from
 * the ZIP reader, so the error surface is a security boundary and not developer
 * ergonomics.
 */

/**
 * Machine-readable failure reasons.
 *
 * Several of these were chosen by breaking a real deck and watching the
 * installed PowerPoint refuse it, rather than by reading the specification and
 * guessing which rules are enforced. The ones marked below are cases where
 * PowerPoint returns `0x80070570` ("corrupted and unreadable") - which means
 * refusing them is not us being stricter than the format, it is us agreeing
 * with the only implementation that matters.
 */
export const XML_ERROR_CODES = [
  /**
   * The document declares a DOCTYPE.
   *
   * Its own code rather than `ERR_MALFORMED_XML`, because a DTD is not a
   * malformed file - it is a well-formed attack surface. Rejecting the
   * construct outright removes XXE, parameter entities and the billion-laughs
   * expansion entirely, because there is no internal subset left to expand.
   *
   * Verified against PowerPoint: a slide part carrying `<!DOCTYPE p:sld>` is
   * refused with `0x80070570`. Rejecting it costs us no real file.
   */
  'ERR_DOCTYPE_FORBIDDEN',
  /** The bytes are not well-formed XML. The detail carries the offset. */
  'ERR_MALFORMED_XML',
  /**
   * An end tag names an element other than the one that is open.
   *
   * Verified against PowerPoint: refused.
   */
  'ERR_MISMATCHED_TAG',
  /** Input ended with elements still open. */
  'ERR_UNCLOSED_TAG',
  /** More than one element at the top level, or none at all. */
  'ERR_ROOT_ELEMENT',
  /**
   * One element carries the same attribute name twice.
   *
   * A well-formedness error in XML 1.0 §3.1, and verified against PowerPoint:
   * `<a:off x="0" y="0" x="0"/>` is refused with `0x80070570`.
   */
  'ERR_DUPLICATE_ATTRIBUTE',
  /**
   * A reference names an entity that does not exist.
   *
   * With no DTD, the only entities are the five predefined ones. Verified
   * against PowerPoint: `&nbsp;` in an `<a:t>` is refused.
   */
  'ERR_UNKNOWN_ENTITY',
  /**
   * A character, or a character reference, outside the XML 1.0 `Char`
   * production (§2.2).
   *
   * Verified against PowerPoint: a literal U+000B in an `<a:t>` is refused.
   * This is the rule that stops `&#xD800;` smuggling a lone surrogate into a
   * string that must survive being re-encoded as UTF-8.
   */
  'ERR_INVALID_CHARACTER',
  /**
   * The bytes are not UTF-8, or declare an encoding this package does not
   * implement.
   *
   * A deliberate and documented gap rather than a bug: PowerPoint honours the
   * `encoding` pseudo-attribute and opens a part written in windows-1252 or
   * UTF-16, and we do not. All 2834 XML parts in our corpus are UTF-8, and
   * failing loudly with the encoding named beats mis-decoding silently.
   */
  'ERR_UNSUPPORTED_ENCODING',
  /**
   * A structural limit was exceeded: source length, nesting depth, attribute
   * count, node count, or the length of a reference body.
   */
  'ERR_LIMIT_EXCEEDED',
] as const;

export type XmlErrorCode = (typeof XML_ERROR_CODES)[number];

/** Structured context for a failure. Every field is optional and diagnostic. */
export interface XmlErrorDetail {
  /** Offset into the decoded source, in UTF-16 code units, when positional. */
  readonly offset?: number;
  /**
   * The element, attribute or entity name the failure is attributable to.
   *
   * Explicitly nullable rather than merely optional: under
   * `exactOptionalPropertyTypes` a caller that has a `string | undefined` in
   * hand cannot pass it to an optional-only field, and every caller here does.
   */
  readonly name?: string | undefined;
  /** The limit that was exceeded. */
  readonly limit?: number;
  /** The value that exceeded it. */
  readonly actual?: number;
  /** Whatever a dependency threw, preserved rather than swallowed. */
  readonly cause?: unknown;
}

export class XmlError extends Error {
  readonly code: XmlErrorCode;
  readonly detail: XmlErrorDetail;

  constructor(code: XmlErrorCode, message: string, detail: XmlErrorDetail = {}) {
    super(message, detail.cause === undefined ? undefined : { cause: detail.cause });
    this.name = 'XmlError';
    this.code = code;
    this.detail = detail;
  }
}

export function isXmlError(value: unknown): value is XmlError {
  return value instanceof XmlError;
}

/**
 * A 1-based line and column for an offset, computed on demand.
 *
 * Deliberately not tracked during the scan. Every real part is one line, so
 * maintaining a line counter through the hot loop would cost every document to
 * benefit only the diagnostics of a broken one.
 */
export function positionOf(source: string, offset: number): { line: number; column: number } {
  const clamped = Math.max(0, Math.min(offset, source.length));
  let line = 1;
  let lineStart = 0;
  for (let i = 0; i < clamped; i++) {
    if (source.charCodeAt(i) === 0x0a) {
      line++;
      lineStart = i + 1;
    }
  }
  return { line, column: clamped - lineStart + 1 };
}

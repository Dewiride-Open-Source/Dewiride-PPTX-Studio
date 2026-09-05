/**
 * The single error type this package throws.
 *
 * The invariant is narrow and load-bearing: **every failure path out of
 * `@pptx-studio/opc` throws an `OpcError` and nothing else.** Not a
 * `RangeError`, not a `TypeError`, not whatever a dependency happened to
 * throw. Phase 12's fuzzing harness asserts exactly this, and it is not a
 * hypothetical concern - `fflate`'s inflate throws a bare
 * `RangeError: offset is out of bounds` when a DEFLATE stored block overruns a
 * caller-supplied output buffer, which is precisely the situation a malicious
 * archive engineers. See `docs/adr/phase-0-foundation/0002-zip-reader.md`.
 *
 * Untrusted input arrives here first, so the error surface is a security
 * boundary rather than developer ergonomics.
 */

/**
 * Machine-readable failure reasons.
 *
 * The UI switches on these, so they are part of the public API and are added to
 * rather than renamed.
 */
export const OPC_ERROR_CODES = [
  /** The bytes are not a ZIP archive at all - no end-of-central-directory record. */
  'ERR_NOT_A_ZIP',
  /** OLE2/CFB magic: a password-protected `.pptx`, or a pre-2007 binary `.ppt`. */
  'ERR_ENCRYPTED_PACKAGE',
  /** A structure runs off the end of the buffer. */
  'ERR_TRUNCATED',
  /** The archive itself is larger than the configured budget. */
  'ERR_ARCHIVE_TOO_LARGE',
  /** More central-directory entries than the configured budget allows. */
  'ERR_TOO_MANY_ENTRIES',
  /** The running inflated-bytes counter would pass its limit. */
  'ERR_BUDGET_EXCEEDED',
  /** One entry declares more inflated bytes than a single part may have. */
  'ERR_ENTRY_TOO_LARGE',
  /** One entry's declared inflation ratio is above the configured ceiling. */
  'ERR_RATIO_EXCEEDED',
  /** Compression method other than stored (0) or deflate (8). */
  'ERR_UNSUPPORTED_COMPRESSION',
  /** The entry sets a general-purpose encryption bit. */
  'ERR_ENTRY_ENCRYPTED',
  /** The local file header disagrees with the central directory. */
  'ERR_HEADER_MISMATCH',
  /** Two entries claim overlapping byte ranges - the quoted-overlap bomb shape. */
  'ERR_OVERLAPPING_ENTRY',
  /** Inflated bytes do not match the stored CRC-32. */
  'ERR_CRC_MISMATCH',
  /** The DEFLATE stream is malformed. */
  'ERR_INFLATE_FAILED',
  /** A ZIP64 field holds a value this reader cannot represent. */
  'ERR_ZIP64_UNSUPPORTED',
  /** Two entries claim the same name - which one a consumer sees becomes ambiguous. */
  'ERR_DUPLICATE_ENTRY',
  /** A ZIP entry name is not a usable OPC part name. */
  'ERR_INVALID_PART_NAME',
  /** A relationship target resolves outside the package root. */
  'ERR_TARGET_ESCAPES_PACKAGE',

  // --- 0.3: the package on top of the archive ----------------------------

  /** A package-level XML part is not well-formed enough to read. */
  'ERR_MALFORMED_XML',
  /**
   * The XML declares a DOCTYPE.
   *
   * Its own code rather than `ERR_MALFORMED_XML` because this is not a
   * malformed file, it is a well-formed attack: a DTD is where XXE and the
   * billion-laughs expansion live. We reject the construct rather than
   * implementing it safely.
   */
  'ERR_DOCTYPE_FORBIDDEN',
  /** `[Content_Types].xml` or `_rels/.rels` is absent. Neither is optional. */
  'ERR_MISSING_PACKAGE_PART',
  /** A part has neither an `Override` nor a `Default` for its extension. */
  'ERR_MISSING_CONTENT_TYPE',
  /** Two `Default`s for one extension, or two `Override`s for one part, disagree. */
  'ERR_CONTENT_TYPE_CONFLICT',
  /** A content type is not a syntactically valid media type. */
  'ERR_INVALID_CONTENT_TYPE',
  /** A relationship `Id` is not an `xsd:ID`, which is to say not an XML `NCName`. */
  'ERR_INVALID_RELATIONSHIP_ID',
  /** Two relationships in one `.rels` part share an `Id`. */
  'ERR_DUPLICATE_RELATIONSHIP_ID',
  /** An internal relationship target names a part the package does not contain. */
  'ERR_DANGLING_RELATIONSHIP',
  /** A relationship targets something no relationship may target - another `.rels` part. */
  'ERR_INVALID_RELATIONSHIP_TARGET',
  /** A part was asked for by name and is not in the store. */
  'ERR_PART_NOT_FOUND',
  /** A part was added under a name the store already holds. */
  'ERR_PART_EXISTS',
  /** The package would need ZIP64 to be written, and this writer emits ZIP32 only. */
  'ERR_ZIP32_OVERFLOW',
] as const;

export type OpcErrorCode = (typeof OPC_ERROR_CODES)[number];

/** Structured context for a failure. Every field is optional and diagnostic. */
export interface OpcErrorDetail {
  /** ZIP entry name or OPC part name, when the failure is attributable to one. */
  readonly entry?: string;
  /** Byte offset into the archive, when the failure is positional. */
  readonly offset?: number;
  /** The budget that was exceeded. */
  readonly limit?: number;
  /** The value that exceeded it. */
  readonly actual?: number;
  /** Whatever a dependency threw, preserved rather than swallowed. */
  readonly cause?: unknown;
}

export class OpcError extends Error {
  readonly code: OpcErrorCode;
  readonly detail: OpcErrorDetail;

  constructor(code: OpcErrorCode, message: string, detail: OpcErrorDetail = {}) {
    super(message, detail.cause === undefined ? undefined : { cause: detail.cause });
    this.name = 'OpcError';
    this.code = code;
    this.detail = detail;
  }
}

export function isOpcError(value: unknown): value is OpcError {
  return value instanceof OpcError;
}

/**
 * Run `fn`, converting anything it throws that is not already an `OpcError`
 * into one. This is the only sanctioned way to call into a dependency that
 * handles untrusted bytes.
 */
export function guard<T>(
  code: OpcErrorCode,
  message: string,
  detail: OpcErrorDetail,
  fn: () => T,
): T {
  try {
    return fn();
  } catch (cause) {
    if (isOpcError(cause)) throw cause;
    const reason = cause instanceof Error ? cause.message : String(cause);
    throw new OpcError(code, message + ' (' + reason + ')', { ...detail, cause });
  }
}

/**
 * Typed failures, as everywhere else in this repository.
 *
 * Colour resolution can fail for exactly one reason: it was asked for something
 * it was not given. A `schemeClr` with no theme, a `phClr` outside a style
 * invocation, a `sysClr` this machine has no value for. Each of those is a
 * caller mistake with a different fix, so each gets its own code rather than a
 * shared "could not resolve".
 *
 * Note what is *not* here. There is no error for a colour that comes out too
 * dark, out of gamut, or with a saturation above 1. Those are ordinary results
 * of arithmetic PowerPoint also performs, and 0.7-C measured what it does with
 * them; see `apply.ts`.
 */
export type PaintErrorCode =
  /** An `a:srgbClr/@val` that is not six hex digits. */
  | 'COLOR_HEX'
  /** A `ST_Percentage` or `ST_Angle` attribute that does not parse. */
  | 'COLOR_NUMBER'
  /** A scheme colour name outside `ST_SchemeColorVal`. */
  | 'COLOR_SCHEME_NAME'
  /** A `schemeClr` was resolved with no `clrScheme` in the context. */
  | 'COLOR_NO_SCHEME'
  /** `schemeClr val="phClr"` outside a style invocation that supplies one. */
  | 'COLOR_NO_PHCLR'
  /** A theme whose slots refer to each other in a cycle. */
  | 'COLOR_SCHEME_CYCLE'
  /** A `sysClr` with no `lastClr` and a name this build has no value for. */
  | 'COLOR_SYS_UNKNOWN'
  /** A `prstClr` name outside `ST_PresetColorVal`. */
  | 'COLOR_PRST_UNKNOWN'
  /** A transform name outside the twenty-eight DrawingML defines. */
  | 'COLOR_TRANSFORM'
  /**
   * An `a:gradFill` with no stops to interpolate between.
   *
   * PowerPoint refuses the package for this - both an empty `a:gsLst` and a
   * single stop come back repaired - so a deck in the wild will not contain one,
   * and a caller that has synthesised one has made a mistake worth naming.
   */
  | 'FILL_NO_STOPS'
  /** A `pattFill/@prst` outside the 54 names PowerPoint accepts. */
  | 'FILL_PATTERN_UNKNOWN';

export class PaintError extends Error {
  readonly code: PaintErrorCode;
  /** The offending token, when the failure is about one. */
  readonly token: string | null;

  constructor(code: PaintErrorCode, message: string, token: string | null = null) {
    super(message);
    this.name = 'PaintError';
    this.code = code;
    this.token = token;
  }
}

/**
 * Every failure in the fidelity harness is a `FidelityError` with a code.
 *
 * The harness exists to be believed, so it has no degraded mode: a run that
 * cannot make the comparison it claims to make stops and names what was
 * missing rather than scoring something else. ADR 0035.
 */

export type FidelityErrorCode =
  /** A grid file's magic, version or dimensions are not what this reads. */
  | 'FID_GRID_MALFORMED'
  /** Two grids being compared do not describe the same raster. */
  | 'FID_GRID_MISMATCH'
  /** A raster whose dimensions are not a whole number of cells. */
  | 'FID_RASTER_GEOMETRY'
  /** The SVG's intrinsic size is not the size it was asked to draw at. */
  | 'FID_SVG_INTRINSIC_SIZE'
  /** A typeface the baseline was recorded with measures differently here. */
  | 'FID_FONT_METRICS_CHANGED'
  /** A typeface the baseline was recorded with rasterises differently here. */
  | 'FID_FONT_COVERAGE_CHANGED'
  /** No environment lock matches this machine, so nothing may be compared. */
  | 'FID_ENV_UNKNOWN'
  /** A slide rasterised twice while recording did not agree with itself. */
  | 'FID_RASTER_NONDETERMINISTIC'
  /** The browser is not the one the baseline was recorded through. */
  | 'FID_BROWSER_CHANGED'
  /** A slide was scored against a baseline that has no digest for it. */
  | 'FID_BASELINE_INCOMPLETE'
  /** An oracle PNG carries a colour profile, so decoding it is not portable. */
  | 'FID_ORACLE_COLOR_MANAGED'
  /** A slide named in the manifest has no committed oracle grid. */
  | 'FID_ORACLE_MISSING'
  /** The renderer's output changed against the recorded digest. */
  | 'FID_RENDER_CHANGED'
  /** A re-record was asked for without saying why, or without the count. */
  | 'FID_RECORD_UNJUSTIFIED'
  /** A re-record was attempted in CI, where no human is reading the diff. */
  | 'FID_RECORD_IN_CI';

export const FIDELITY_ERROR_CODES: readonly FidelityErrorCode[] = [
  'FID_GRID_MALFORMED',
  'FID_GRID_MISMATCH',
  'FID_RASTER_GEOMETRY',
  'FID_SVG_INTRINSIC_SIZE',
  'FID_FONT_METRICS_CHANGED',
  'FID_FONT_COVERAGE_CHANGED',
  'FID_ENV_UNKNOWN',
  'FID_RASTER_NONDETERMINISTIC',
  'FID_BROWSER_CHANGED',
  'FID_BASELINE_INCOMPLETE',
  'FID_ORACLE_COLOR_MANAGED',
  'FID_ORACLE_MISSING',
  'FID_RENDER_CHANGED',
  'FID_RECORD_UNJUSTIFIED',
  'FID_RECORD_IN_CI',
];

export class FidelityError extends Error {
  readonly code: FidelityErrorCode;
  /** Whatever names the thing that went wrong: a slide, a typeface, a file. */
  readonly subject: string | undefined;

  constructor(code: FidelityErrorCode, message: string, subject?: string) {
    super(message);
    this.name = 'FidelityError';
    this.code = code;
    this.subject = subject;
  }
}

export function isFidelityError(value: unknown): value is FidelityError {
  return value instanceof FidelityError;
}

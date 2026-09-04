/**
 * Every failure in this package is a `RenderError` with a machine-readable code.
 *
 * The repository rule, and it matters more here than elsewhere: a renderer runs
 * in a Web Worker over files nobody vetted, and a bare `TypeError` from three
 * layers down tells the caller nothing about which shape on which slide it
 * could not draw.
 *
 * What is *not* an error is worth stating too. A shape with no geometry, a
 * placeholder that matched nothing, a `fillRef` past the end of the style
 * matrix, an `a:blipFill` whose image has not been decoded: PowerPoint draws
 * all four without complaint, and so does this. An error here means the caller
 * asked for something impossible, or the package was handed markup that names a
 * shape nobody has heard of.
 */

export type RenderErrorCode =
  /** `a:prstGeom/@prst` is not one of the 187 `ST_ShapeType` values. */
  | 'RENDER_UNKNOWN_PRESET'
  /** A group tree nests deeper than `MAX_GROUP_DEPTH`. Hostile input. */
  | 'RENDER_GROUP_DEPTH'
  /** The host handed `mount` something that is not an element. */
  | 'RENDER_NO_HOST';

export const RENDER_ERROR_CODES: readonly RenderErrorCode[] = [
  'RENDER_UNKNOWN_PRESET',
  'RENDER_GROUP_DEPTH',
  'RENDER_NO_HOST',
];

export class RenderError extends Error {
  readonly code: RenderErrorCode;
  /** Whatever names the thing that went wrong: a shape name, a preset, a host. */
  readonly subject: string | undefined;

  constructor(code: RenderErrorCode, message: string, subject?: string) {
    super(message);
    this.name = 'RenderError';
    this.code = code;
    this.subject = subject;
  }
}

export function isRenderError(value: unknown): value is RenderError {
  return value instanceof RenderError;
}

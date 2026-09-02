/**
 * Typed failures, as everywhere else in this repository.
 *
 * A bare `RangeError` out of a geometry routine tells a caller nothing it can
 * branch on, and this package is reached from a render loop where the useful
 * response to "this one preset is broken" is to draw the other 40 shapes and
 * report the one. So every throw carries a `code`.
 */
export type GeometryErrorCode =
  /** The generated preset data did not decode. A build bug, never a user's file. */
  | 'PRESET_DECODE'
  /** A preset name that is not one of the 187. */
  | 'PRESET_UNKNOWN'
  /** A `fmla` whose first token is not one of the seventeen operators. */
  | 'FMLA_OPERATOR'
  /** An operator given the wrong number of operands. */
  | 'FMLA_ARITY'
  /** An operand that is neither an integer literal nor a guide in scope. */
  | 'FMLA_OPERAND';

export class GeometryError extends Error {
  readonly code: GeometryErrorCode;
  /** The preset being decoded when this happened, when there is one. */
  readonly preset: string | null;

  constructor(code: GeometryErrorCode, message: string, preset: string | null = null) {
    super(message);
    this.name = 'GeometryError';
    this.code = code;
    this.preset = preset;
  }
}

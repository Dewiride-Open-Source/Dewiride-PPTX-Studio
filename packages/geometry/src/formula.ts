import { GeometryError } from './errors.js';

/**
 * The seventeen `a:gd/@fmla` operators, as pure arithmetic.
 *
 * A formula is an operator token followed by its operands, already resolved to
 * numbers. Nothing here knows what a guide is, what a shape is, or where an
 * operand came from - that is `evaluate.ts`. This module is the part that can
 * be checked against arithmetic alone.
 *
 * ## How this table was established
 *
 * Not by recalling it. Every operator below is cross-checked against the 187
 * transcoded presets, because the preset data is a corpus of 3922 worked
 * examples and several operators give themselves away in it. The evidence is
 * recorded next to each operator rather than in a commit message, because the
 * cost of a wrong row is a shape that draws quietly wrong rather than throwing.
 *
 * The single richest witness is `blockArc`, which pins six of the seventeen at
 * once - see the notes on `pin`, the ternary, and the two arc-tangent
 * operators.
 */

// The table itself lives in line comments: the multiply-divide operator's own
// two characters end a block comment, which is a real syntax error and not a
// formatting quibble.
//
//   */     x y z   ->  x * y / z
//   +-     x y z   ->  x + y - z
//   +/     x y z   ->  (x + y) / z
//   ?:     x y z   ->  x > 0 ? y : z
//   abs    x       ->  |x|
//   at2    x y     ->  arctan(y / x), in angle units, NOT normalised
//   cat2   x y z   ->  x * cos(arctan(z / y))
//   cos    x y     ->  x * cos(y)
//   max    x y     ->  the larger
//   min    x y     ->  the smaller
//   mod    x y z   ->  sqrt(x^2 + y^2 + z^2)     <- a length, not modulo
//   pin    x y z   ->  y clamped to [x, z]
//   sat2   x y z   ->  x * sin(arctan(z / y))
//   sin    x y     ->  x * sin(y)
//   sqrt   x       ->  sqrt(|x|)                 <- Office, not ECMA. See below.
//   tan    x y     ->  x * tan(y)
//   val    x       ->  x

/** One full turn, in 60000ths of a degree. */
export const FULL_CIRCLE = 21600000;

/** DrawingML states angles in 60000ths of a degree, everywhere, without exception. */
export const ANGLE_UNITS_PER_DEGREE = 60000;

/** 60000ths of a degree to radians. `pi / (180 * 60000)`. */
const TO_RADIANS = Math.PI / 10800000;

/** Radians back to 60000ths of a degree. */
const FROM_RADIANS = 10800000 / Math.PI;

/** The operator tokens, as a type. */
export type FmlaOperator =
  | '*/'
  | '+-'
  | '+/'
  | '?:'
  | 'abs'
  | 'at2'
  | 'cat2'
  | 'cos'
  | 'max'
  | 'min'
  | 'mod'
  | 'pin'
  | 'sat2'
  | 'sin'
  | 'sqrt'
  | 'tan'
  | 'val';

/**
 * How many operands each operator takes.
 *
 * Confirmed against the corpus rather than assumed: every one of the 3922
 * guides in the 187 presets uses exactly these counts, with one documented
 * exception handled in `checkArity`.
 */
export const FMLA_ARITY: Readonly<Record<FmlaOperator, number>> = {
  '*/': 3,
  '+-': 3,
  '+/': 3,
  '?:': 3,
  abs: 1,
  at2: 2,
  cat2: 3,
  cos: 2,
  max: 2,
  min: 2,
  mod: 3,
  pin: 3,
  sat2: 3,
  sin: 2,
  sqrt: 1,
  tan: 2,
  val: 1,
};

const OPERATORS = new Set<string>(Object.keys(FMLA_ARITY));

export function isFmlaOperator(token: string): token is FmlaOperator {
  return OPERATORS.has(token);
}

/** Convert an angle in 60000ths of a degree to radians. */
export function angleToRadians(angle: number): number {
  return angle * TO_RADIANS;
}

/** Convert radians to 60000ths of a degree. */
export function radiansToAngle(radians: number): number {
  return radians * FROM_RADIANS;
}

/**
 * Where a formula came from, for the error message and for `GeometryError.preset`.
 *
 * A renderer's useful response to one broken shape is to draw the other forty
 * and report the one, which it can only do if the failure says which shape it
 * was. Hence a structured site rather than a pre-formatted string.
 */
export interface FormulaSite {
  readonly preset: string | null;
  readonly guide: string | null;
}

/**
 * The `shape.guide: ` prefix a failure message carries, or nothing.
 *
 * Nothing is the normal case for an `a:custGeom`, which has no name, and for a
 * failure that is not inside a named guide - so the empty label has to collapse
 * rather than leave a bare colon behind.
 */
export function siteLabel(site: FormulaSite | null): string {
  if (site === null) return '';
  const parts = [site.preset, site.guide].filter((p): p is string => p !== null);
  return parts.length === 0 ? '' : `${parts.join('.')}: `;
}

function fail(
  code: 'FMLA_OPERATOR' | 'FMLA_ARITY',
  message: string,
  site: FormulaSite | null,
): never {
  throw new GeometryError(code, `${siteLabel(site)}${message}`, site?.preset ?? null);
}

/**
 * Refuse an operand count the operator does not take.
 *
 * ## The one exception, and why it is narrow
 *
 * Eight formulas in POI's file give the add-subtract operator four operands
 * where it takes three. All eight are in the circular-arrow family and all
 * eight have the same shape:
 *
 * ```text
 * circularArrow.xB = +- xH 0 dxB 0
 * ```
 *
 * The fourth operand is the literal `0` in every case. That makes the anomaly
 * *provably* harmless rather than probably harmless: `x + y - z` ignoring a
 * trailing zero, `x + y - z + w` and `x + y - z - w` all give the same number
 * when `w` is zero, so every reading of the extra operand agrees and there is
 * nothing to get wrong.
 *
 * So a fourth operand is accepted when it is exactly zero and refused
 * otherwise. A future POI that writes a non-zero fourth operand is a case where
 * the readings genuinely diverge, and the honest response to that is to stop,
 * not to pick one silently. Accepting the benign case while refusing the
 * ambiguous one is the whole point of not simply widening the arity to four.
 */
function checkArity(op: FmlaOperator, args: readonly number[], site: FormulaSite | null): void {
  const want = FMLA_ARITY[op];
  if (args.length === want) return;

  if (op === '+-' && args.length === 4 && args[3] === 0) return;

  fail(
    'FMLA_ARITY',
    `operator "${op}" takes ${String(want)} operand(s), given ${String(args.length)}`,
    site,
  );
}

/**
 * Apply one operator to operands that are already numbers.
 *
 * Everything is an IEEE double and nothing is rounded or truncated. That is a
 * decision, not an omission: `circularArrow` chains over 200 guides deep, and
 * rounding each intermediate to an integer would accumulate visible error by
 * the end of it. Integers appear in the source data, not in the arithmetic.
 *
 * Division by zero yields the IEEE result and does not throw - see
 * `evaluateGuides` in `evaluate.ts`, where the reachability of that case is
 * measured rather than hoped about.
 *
 * @param op the operator token
 * @param args its operands, already resolved
 * @param site optional origin, used for the error message and `error.preset`
 */
export function applyOperator(
  op: string,
  args: readonly number[],
  site: FormulaSite | null = null,
): number {
  if (!isFmlaOperator(op)) fail('FMLA_OPERATOR', `unknown formula operator "${op}"`, site);
  checkArity(op, args, site);

  // Read past the arity check, so these are present by construction. The
  // non-null assertions are the price of `noUncheckedIndexedAccess`, which is
  // worth keeping on for the rest of the package.
  const x = args[0] as number;
  const y = args[1] as number;
  const z = args[2] as number;

  switch (op) {
    case '*/':
      return (x * y) / z;

    case '+-':
      // `blockArc.iswAng = +- 0 0 swAng` is the negation idiom, which only
      // works if the third operand is subtracted.
      return x + y - z;

    case '+/':
      return (x + y) / z;

    case '?:':
      // The threshold is zero and the comparison is strict. `blockArc` uses it
      // to normalise a sweep angle - `?: sw11 sw11 sw12`, where `sw12` is
      // `sw11` plus a full circle - and `circularArrow` uses the same idiom
      // three times over. At exactly zero the two branches agree there (0 and
      // 360 degrees are the same angle), so the preset corpus cannot separate
      // `> 0` from `>= 0`; ECMA specifies greater-than, and that is what this
      // is. See `formula.test.ts`, which pins the boundary so a future change
      // to it is deliberate.
      return x > 0 ? y : z;

    case 'abs':
      return Math.abs(x);

    case 'at2':
      // `arctan(y / x)`, so the operands are in the opposite order to
      // `Math.atan2`. `circularArrow.u13 = at2 1 u12` is the witness: it reads
      // as arctan of `u12`, which requires the *first* operand to be the
      // denominator.
      //
      // The result is NOT normalised into [0, 360). Three separate guides in
      // `circularArrow` follow an `at2` with `?: a a (a + 21600000)`, which is
      // the caller normalising by hand and would be dead code if this operator
      // had done it. Returning a normalised angle here would silently disable
      // that idiom and reverse the sweep direction of every circular arrow.
      //
      // The origin is guarded explicitly rather than left to IEEE, because
      // `Math.atan2(0, -0)` is pi, not zero. A guide chain that arrives at a
      // negative zero - trivially, `*/ 0 1 -1` - would otherwise turn a
      // zero-length vector into a 180 degree angle, which is the kind of error
      // that draws a plausible shape pointing the wrong way.
      if (x === 0 && y === 0) return 0;
      return radiansToAngle(Math.atan2(y, x));

    case 'cat2':
      // x * cos(arctan(z / y)). The operand order matters and is not guessable
      // from the name, so it is pinned by geometry instead:
      //
      //   blockArc:  wt1 = sin wd2 stAng      = (w/2) * sin(t)
      //              ht1 = cos hd2 stAng      = (h/2) * cos(t)
      //              dx1 = cat2 wd2 ht1 wt1
      //              dy1 = sat2 hd2 ht1 wt1
      //
      // With the order below, `dx1` and `dy1` are (w/2)cos(u), (h/2)sin(u) for
      // u = atan2(wt1, ht1) - which is exactly the ellipse-angle unskew, so
      // (hc+dx1, vc+dy1) lands on the ellipse. Swap `y` and `z` and the start
      // of a `blockArc` at angle zero moves from 3 o'clock to 6 o'clock.
      return x * Math.cos(Math.atan2(z, y));

    case 'cos':
      return x * Math.cos(angleToRadians(y));

    case 'max':
      return Math.max(x, y);

    case 'min':
      return Math.min(x, y);

    case 'mod':
      // The vector modulus, sqrt(x^2 + y^2 + z^2) - a length. NOT modulo.
      //
      // Arity alone settles it: a binary remainder cannot take three operands,
      // and all 48 uses in the corpus take three. The names confirm it -
      // `cornerTabs.md = mod w h 0` and `mathMultiply.dl = mod w h 0` are the
      // shape's diagonal, `gear6.lAD1 = mod xAD1 yAD1 0` is a segment length.
      //
      // The third operand is the literal zero in all 48 uses, so every preset
      // use degenerates to a plain hypotenuse. It is still computed as three
      // dimensions, because custom geometry in a user's deck is under no such
      // restriction.
      return Math.sqrt(x * x + y * y + z * z);

    case 'pin':
      // Clamp the MIDDLE operand: `pin min value max`.
      //
      // `blockArc.stAng = pin 0 adj1 21599999` is decisive - the adjust value
      // is what is being constrained, to a legal angle. Reading the first
      // operand as the value would clamp the constant 0 by the adjust, which
      // is meaningless.
      //
      // Written as an explicit if-chain rather than as `max(x, min(y, z))`,
      // and the two are NOT the same function. They agree whenever the bounds
      // are the right way round and diverge when they are inverted: for
      // `pin 10 20 3` the chain gives 3 and the min/max composition gives 10.
      // Apache POI rewrote this operator into the min/max form in 5.2.5, so
      // the divergence is live in a shipped implementation rather than
      // theoretical. The chain is what ECMA describes, so the chain is what
      // this is.
      //
      // Nothing in the preset data inverts its bounds - every use is of the
      // form `pin 0 adjN <constant>`. Hand-authored custom geometry can.
      if (y < x) return x;
      if (y > z) return z;
      return y;

    case 'sat2':
      // x * sin(arctan(z / y)). The sine half of the pair above.
      return x * Math.sin(Math.atan2(z, y));

    case 'sin':
      // The multiply by the first operand is real, not a misreading of a unary
      // sine: `blockArc.wt1 = sin wd2 stAng` is half the width times the sine
      // of the angle, which is what makes the ellipse construction work.
      return x * Math.sin(angleToRadians(y));

    case 'sqrt':
      // sqrt(|x|), NOT sqrt(x) - a deliberate divergence from ECMA, following
      // Office.
      //
      // Microsoft's own implementer notes for the standard (MS-OI29500) record
      // this as one of the four places Office departs from ECMA in this
      // section: the standard defines `sqrt x` as the square root of x, and
      // Office computes the square root of its absolute value. Apache POI
      // implements the standard and returns NaN; LibreOffice ends at zero.
      //
      // Following the standard here is the worse failure. NaN is contagious:
      // every guide downstream of it becomes NaN, and a shape whose points are
      // NaN does not draw a wrong shape, it silently draws nothing. All 19
      // uses in the presets take a computed guide as their operand - the
      // radicands are driven by adjust values, in the circular and curved
      // arrows, `moon` and `teardrop` - so the sign is not a constant anyone
      // can eyeball.
      //
      // Whether the negative case is actually reachable from a legal adjust
      // value in the shipped presets is measured in `evaluate.test.ts` rather
      // than assumed, in either direction.
      return Math.sqrt(Math.abs(x));

    case 'tan':
      return x * Math.tan(angleToRadians(y));

    case 'val':
      return x;
  }
}

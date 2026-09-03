import { PaintError } from './errors.js';

/**
 * The two attribute grammars a colour is written in, and the one that has two
 * spellings.
 *
 * This package takes numbers, not markup - 2.9 owns the tree walk. What it does
 * own is the step from an attribute's *text* to a number, because `ST_Percentage`
 * is a union type and getting it wrong is silent.
 */

const INTEGER = /^[-+]?\d+$/;
// `40%`, `-12.5%`, `.5%`. A sign, then digits with an optional fraction or a
// bare fraction, then the per-cent sign.
const PERCENT_LITERAL = /^[-+]?(?:\d+(?:\.\d+)?|\.\d+)%$/;

/**
 * A `ST_Percentage` attribute, in hundred-thousandths.
 *
 * **Both spellings are legal and both appear.** The Transitional grammar writes
 * `val="40000"`; the Strict grammar writes `val="40%"`. That would be a
 * footnote if the two were kept apart by the document's own flavour, and they
 * are not: 2.6 wrote `<a:lumMod val="60%"/><a:lumOff val="40%"/>` into a
 * Transitional slide and PowerPoint opened it without a repair prompt and
 * painted `8FAADC` - the same colour, to the byte, as the `60000`/`40000`
 * control alongside it. `val="60.5%"` works too, which the integer form cannot
 * express at that precision.
 *
 * The failure this prevents is not subtle in its effect and is very subtle in
 * its cause: `parseInt("150%")` is `150`, so a percentage that meant 150%
 * silently becomes 0.15%, and a line spacing or a tint collapses to nothing. The
 * plan flags this for `ST_Percentage` in text (3.1); it is the same attribute
 * type, and colour reaches it first.
 */
export function parsePercentage(token: string): number {
  if (INTEGER.test(token)) return Number(token);
  if (PERCENT_LITERAL.test(token)) return Number(token.slice(0, -1)) * 1000;
  throw new PaintError('COLOR_NUMBER', `"${token}" is not a ST_Percentage`, token);
}

/**
 * An angle attribute, in sixtieths of a degree.
 *
 * Integer only. Unlike `ST_Percentage` there is no second spelling: the Strict
 * schema writes angles the same way, so `hue="7200000"` is 120 degrees in both.
 *
 * Out-of-range values are returned rather than rejected. `a:hslClr/@hue` is a
 * `ST_PositiveFixedAngle` and PowerPoint enforces it - a deck written with
 * `hue="24000000"` was refused outright and only opened after a repair, which
 * reset the hue to zero. We are a reader, and refusing to display a file
 * PowerPoint would repair is a worse outcome than displaying it wrapped.
 */
export function parseAngle(token: string): number {
  if (INTEGER.test(token)) return Number(token);
  throw new PaintError('COLOR_NUMBER', `"${token}" is not an angle`, token);
}

/**
 * `a:srgbClr/@val`, normalised to six upper-case hex digits.
 *
 * A leading `#` is not legal in the format and is accepted here anyway, because
 * a caller that has been handed a CSS colour by a UI will otherwise fail at the
 * least useful moment.
 */
export function parseSrgbValue(token: string): string {
  const raw = token.startsWith('#') ? token.slice(1) : token;
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) {
    throw new PaintError('COLOR_HEX', `"${token}" is not six hex digits`, token);
  }
  return raw.toUpperCase();
}

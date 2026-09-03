/**
 * The three spaces DrawingML computes colour in, and the way out to a byte.
 *
 * Everything in this file is a *measured* answer, not a chosen one. ECMA-376
 * describes `a:tint` as "10% of the input colour combined with 90% white" and
 * never says which space the combining happens in; the credible implementations
 * disagree, and they disagree by up to 73 units out of 255 on colours that
 * appear in every corporate template. Sub-phase 0.7-C put 214 swatches in front
 * of Microsoft PowerPoint and read back what it painted. See
 * `docs/adr/0007-ground-truth.md` and `corpus/ground-truth/color-transforms.json`.
 *
 * The three spaces, and which transforms live in each:
 *
 * | space          | transfer                    | transforms                                        |
 * | -------------- | --------------------------- | ------------------------------------------------- |
 * | linear sRGB    | the sRGB piecewise curve    | `tint` `shade` `inv` `red*` `green*` `blue*`      |
 * | sRGB           | none                        | `gamma` `invGamma` `gray`                         |
 * | HSL over sRGB  | none                        | `hue*` `sat*` `lum*` `comp`                       |
 *
 * The middle row is not a typo, and it is the strongest evidence in the set:
 * `a:gamma` **is** the sRGB de-linearisation and `a:invGamma` is its exact
 * inverse. The file format exposes the curve directly, which independently
 * confirms the curve the top row uses.
 *
 * Scoreboard for the top row, worst error against PowerPoint over 102 tint and
 * shade swatches:
 *
 * | model                                     | worst  | exact   |
 * | ----------------------------------------- | ------ | ------- |
 * | sRGB piecewise linearisation              | 0/255  | 102/102 |
 * | power law gamma 2.3 (LibreOffice)         | 8/255  | 41/102  |
 * | Apache POI's asymmetric pair              | 73/255 | 51/102  |
 * | arithmetic straight on the 0-255 values   | 73/255 | 0/102   |
 *
 * Apache POI is the interesting loser: it linearises `shade` and computes
 * `tint` straight on the sRGB values, so it is exactly right on half the
 * question and 73/255 wrong on the other half.
 */

/** A channel, nominally 0..1 - but see `clamp01` for why "nominally". */
export type Channel = number;

/* -------------------------------------------------------------------------- */
/* sRGB <-> linear                                                            */
/* -------------------------------------------------------------------------- */

/**
 * sRGB to linear light. IEC 61966-2-1's piecewise curve, exactly as written.
 *
 * The knee at 0.04045 matters less than it looks - it moves a mid-grey by under
 * a unit - but the 2.4 exponent against LibreOffice's 2.3 is worth 8/255 in the
 * midtones, which is visible side by side on a themed fill.
 */
export function toLinear(c: Channel): Channel {
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}

/**
 * Linear light back to sRGB.
 *
 * **Clamp before calling this.** `Math.pow` of a negative base and a fractional
 * exponent is `NaN`, not a dark colour, and `a:redOff val="-50000"` on a dark
 * channel reaches it. `clamp01` is not a tidy-up here, it is what stops a
 * legal file from producing `NaN` and painting nothing.
 */
export function fromLinear(c: Channel): Channel {
  return c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;
}

export function clamp01(c: number): number {
  return c < 0 ? 0 : c > 1 ? 1 : c;
}

/* -------------------------------------------------------------------------- */
/* HSL                                                                        */
/* -------------------------------------------------------------------------- */

/**
 * Hue in degrees, saturation and lightness as fractions.
 *
 * `s` and `l` are deliberately not documented as 0..1. PowerPoint does not
 * clamp them between transforms: `satMod val="200000"` and `satMod val="300000"`
 * on `#4472C4` produce *different* colours (`0460FF` and `004EFF`), which a
 * model that pinned `s` at 1 cannot express - it gives the same answer for both
 * and is 19/255 wrong on the second. Clamping happens once, per channel, on the
 * way back to RGB. See `apply.ts`.
 */
export interface Hsl {
  h: number;
  s: number;
  l: number;
}

/** The bi-hexcone conversion. Hue of a grey is 0, which is what PowerPoint uses. */
export function rgbToHsl(r: Channel, g: Channel, b: Channel): Hsl {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) * 60;
  else if (max === g) h = ((b - r) / d + 2) * 60;
  else h = ((r - g) / d + 4) * 60;
  return { h, s, l };
}

/**
 * Back to RGB. Out-of-range `s` and `l` are carried through the arithmetic
 * rather than rejected, and the result may fall outside 0..1 - that is the
 * point of not clamping in HSL, and the caller clamps.
 */
export function hslToRgb(hsl: Hsl): [Channel, Channel, Channel] {
  const { h, s, l } = hsl;
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const component = (t: number): number => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  const k = h / 360;
  return [component(k + 1 / 3), component(k), component(k - 1 / 3)];
}

/** Fold a hue in degrees into [0, 360). */
export function wrapHue(h: number): number {
  const x = h % 360;
  return x < 0 ? x + 360 : x;
}

/* -------------------------------------------------------------------------- */
/* luma                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * Rec.709 luma of the **sRGB** values, with no linearisation.
 *
 * That is not the photometrically defensible thing to do - luminance is defined
 * on linear light - but it is what `a:gray` measurably does: `#FF0000` becomes
 * `363636`, which is `0.2126 x 255` rounded, not the 127 a linear-light
 * conversion gives.
 */
export function luma709(r: Channel, g: Channel, b: Channel): Channel {
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/* -------------------------------------------------------------------------- */
/* bytes                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * A channel to a byte, breaking an exact tie towards an odd result.
 *
 * ## The ties are real, and no rounding mode explains them
 *
 * Across the 269 swatches in the two ground-truth fixtures whose value a correct
 * model predicts, seven land on **exactly** `x.5` - a genuine tie in exact
 * arithmetic, not a floating-point near-miss - and PowerPoint takes the lower
 * byte four times and the higher one three times:
 *
 * | value | PowerPoint | where                             |
 * | ----- | ---------- | --------------------------------- |
 * | 127.5 | 127        | `#FF0000` `lumMod 50%`            |
 * | 127.5 | 127        | `#FF0000` `lumOff 25%`            |
 * | 153.5 | 153        | `#808080` `lumOff 10%`            |
 * | 93.5  | 93         | `hslClr` hue 218, sat 100%, lum 50% |
 * | 110.5 | 111        | `hslClr` hue 218, sat 50%, lum 50%  |
 * | 129.5 | 130        | `#4472C4` `satOff -50%`           |
 * | 134.5 | 135        | `#4472C4` `satOff -50%`           |
 *
 * Half-down fits four of them, half-up three, half-even one. Holding the
 * intermediate in single precision does not explain it either: three of the
 * seven move to the *wrong* side of the tie under float32. Whatever PowerPoint
 * is doing here is not a rounding mode applied to this number.
 *
 * ## So this is a fit, and it is stated as one
 *
 * Round-half-to-odd is the best of the four. It fits six of the seven, and it
 * agrees with half-down on every case sub-phase 0.7-C had, so nothing that
 * experiment settled is disturbed. The one it misses is `129.5 -> 130`, which is
 * asserted in `color.test.ts` rather than smoothed away.
 *
 * That leaves this package exact on 268 of those 269 swatches and off by one
 * unit in one channel on the remaining one. 0.7-C concluded "PowerPoint takes
 * the lower value" from three samples; four more were enough to show that was a
 * coincidence, which is an argument for measuring more of them rather than for
 * trusting the rule.
 */
export function toByte(c: Channel): number {
  // `clamp01` passes `NaN` through - both of its comparisons are false - and a
  // `NaN` here would come out of the arithmetic as a `NaN` byte and print as
  // `"NaNNaNNaN"` in a fill attribute. Nothing in this package can produce one,
  // which is exactly why the guard is cheap and the failure would be baffling.
  if (!Number.isFinite(c)) return 0;
  const scaled = clamp01(c) * 255;
  const floor = Math.floor(scaled);
  // A tolerance, because the tie is a property of the exact arithmetic rather
  // than of the order this expression happened to evaluate in: (2/3 - 218/360)
  // x 6 x 255 is exactly 93.5 and evaluates to 93.50000000000001 in doubles.
  const rounded =
    Math.abs(scaled - floor - 0.5) > 1e-9
      ? Math.round(scaled)
      : floor % 2 === 1
        ? floor
        : floor + 1;
  return rounded < 0 ? 0 : rounded > 255 ? 255 : rounded;
}

const HEX6 = /^[0-9a-fA-F]{6}$/;

/** Split `4472C4` into three channels. Case-insensitive; six digits, no `#`. */
export function parseHex(value: string): [Channel, Channel, Channel] | null {
  if (!HEX6.test(value)) return null;
  return [
    parseInt(value.slice(0, 2), 16) / 255,
    parseInt(value.slice(2, 4), 16) / 255,
    parseInt(value.slice(4, 6), 16) / 255,
  ];
}

/** Six upper-case hex digits, the way the file format writes them. */
export function toHex(r: Channel, g: Channel, b: Channel): string {
  return (
    toByte(r).toString(16).padStart(2, '0') +
    toByte(g).toString(16).padStart(2, '0') +
    toByte(b).toString(16).padStart(2, '0')
  ).toUpperCase();
}

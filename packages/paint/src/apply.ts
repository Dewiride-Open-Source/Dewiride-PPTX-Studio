import { PaintError } from './errors.js';
import {
  clamp01,
  fromLinear,
  hslToRgb,
  luma709,
  rgbToHsl,
  toLinear,
  wrapHue,
  type Channel,
} from './colors/transfer.js';
import type { ColorTransform, Rgba } from './types.js';

/**
 * The transform chain: twenty-eight operators over three colour spaces, applied
 * in document order.
 *
 * ## Every transform starts and ends at a displayable colour
 *
 * This is the shape of the thing, and it is not what the obvious optimisation
 * suggests. Each operator converts the current sRGB triple into the space it
 * works in, does its arithmetic there, and converts straight back - clamping
 * each channel on the way. Consecutive operators in the same space do *not*
 * share an excursion.
 *
 * That costs a conversion per operator and it is not negotiable, because the
 * clamp in the middle is observable. On `#4472C4`:
 *
 * | chain                        | PowerPoint | carrying the state across |
 * | ---------------------------- | ---------- | ------------------------- |
 * | `lumOff 60%` then `lumOff -60%` | `666666` | `4472C4`                  |
 * | `satMod 300%` then `satMod 33.333%` | `556FAA` | `4472C4`             |
 *
 * The first chain drives lightness to 1.16, which paints white; taking 0.6 back
 * off *white* is mid-grey, not the colour you started with. A model that carries
 * the un-clamped lightness across the boundary returns the original and is 94
 * units out of 255 wrong. This package shipped that model for an afternoon;
 * `docs/adr/phase-2-geometry-and-paint/0021-colour.md` records why the earlier measurement did not catch
 * it.
 *
 * Two things that could have made the boundary invisible were checked and do
 * not: an intervening `a:gamma`/`a:invGamma` round trip changes nothing, and
 * neither does an intervening `a:alpha`. Alpha is a fourth channel that no
 * colour operator reads, and it does not interrupt anything.
 *
 * ## What is *not* clamped is the arithmetic inside one operator
 *
 * `satMod val="200000"` and `satMod val="300000"` produce different colours -
 * `0460FF` and `004EFF`. Saturation of 1.56 is meaningful while the conversion
 * back to RGB is happening; it stops being meaningful the moment that
 * conversion finishes. Pinning saturation at 1 before the conversion gives one
 * answer for both and is 19/255 wrong on the second.
 *
 * ## Three spaces
 *
 * `transfer.ts` carries the table and the scoreboard. In short: `tint`, `shade`,
 * `inv` and the nine per-channel operators work in **linear light**; `gamma`,
 * `invGamma` and `gray` work on the **sRGB** values as they stand; the hue,
 * saturation and lightness families and `comp` work in **HSL over sRGB**. A
 * renderer that does everything in one space is wrong on more than half of them.
 *
 * ## Measured, and inferred
 *
 * Measured against PowerPoint on 464 swatches across 0.7-C and 2.6: the space of
 * every operator, document order, the clamp at every boundary, the absence of a
 * clamp within an operator, that alpha is independent, and the three absolute
 * channel setters.
 *
 * Inferred, and marked at its site: nothing here, any more. The last inference -
 * that `a:red`, `a:green` and `a:blue` sit in linear light by symmetry with
 * `redMod` and `redOff` - was measured in 2.6 and holds.
 */

const PERCENT = 100000;
const ANGLE = 60000;

/** One transform: sRGB and alpha in, sRGB and alpha out, both clamped. */
interface Working {
  r: Channel;
  g: Channel;
  b: Channel;
  a: number;
}

function linear(w: Working, f: (rgb: [number, number, number]) => void): void {
  const rgb: [number, number, number] = [toLinear(w.r), toLinear(w.g), toLinear(w.b)];
  f(rgb);
  // Clamp before de-linearising, and not as a tidy-up: `Math.pow` of a negative
  // base and a fractional exponent is `NaN`, and `a:redOff val="-50000"` on a
  // dark channel reaches it. An unclamped round trip turns a legal file into a
  // shape that paints nothing at all.
  w.r = fromLinear(clamp01(rgb[0]));
  w.g = fromLinear(clamp01(rgb[1]));
  w.b = fromLinear(clamp01(rgb[2]));
}

function hsl(w: Working, f: (h: { h: number; s: number; l: number }) => void): void {
  const value = rgbToHsl(w.r, w.g, w.b);
  f(value);
  const [r, g, b] = hslToRgb(value);
  w.r = clamp01(r);
  w.g = clamp01(g);
  w.b = clamp01(b);
}

/**
 * Apply one transform chain to a colour that has already been reduced to sRGB.
 *
 * The input's channels are sRGB in 0..1; the output's are clamped. `resolveColor`
 * is what callers normally want - this is exported because a theme slot carries
 * transforms of its own, so two chains run in sequence, and because it is what
 * the ground-truth fixtures exercise directly.
 */
export function applyTransforms(base: Rgba, transforms: readonly ColorTransform[]): Rgba {
  const w: Working = { r: clamp01(base.r), g: clamp01(base.g), b: clamp01(base.b), a: base.a };

  for (const t of transforms) {
    switch (t.op) {
      /* ---- alpha: a fourth channel, in no colour space at all ------------ */
      case 'alpha':
        w.a = t.val / PERCENT;
        break;
      case 'alphaOff':
        w.a += t.val / PERCENT;
        break;
      case 'alphaMod':
        w.a *= t.val / PERCENT;
        break;

      /* ---- linear light --------------------------------------------------- */
      case 'tint': {
        // "p of the input colour combined with (1-p) white", in linear light.
        const p = t.val / PERCENT;
        linear(w, (c) => {
          c[0] = c[0] * p + (1 - p);
          c[1] = c[1] * p + (1 - p);
          c[2] = c[2] * p + (1 - p);
        });
        break;
      }
      case 'shade': {
        const p = t.val / PERCENT;
        linear(w, (c) => {
          c[0] *= p;
          c[1] *= p;
          c[2] *= p;
        });
        break;
      }
      case 'inv':
        // A complement in *linear* light. `#4472C4` inverts to `F8EBB3`, not to
        // the channel complement `BB8D3B`: `255 - c` is wrong by up to 100
        // units, and is nothing like `comp`, which is a hue rotation.
        linear(w, (c) => {
          c[0] = 1 - c[0];
          c[1] = 1 - c[1];
          c[2] = 1 - c[2];
        });
        break;
      case 'red':
      case 'green':
      case 'blue':
      case 'redMod':
      case 'greenMod':
      case 'blueMod':
      case 'redOff':
      case 'greenOff':
      case 'blueOff': {
        const p = t.val / PERCENT;
        const i = t.op.startsWith('red') ? 0 : t.op.startsWith('green') ? 1 : 2;
        linear(w, (c) => {
          if (t.op.endsWith('Mod')) c[i] *= p;
          else if (t.op.endsWith('Off')) c[i] += p;
          else c[i] = p;
        });
        break;
      }

      /* ---- sRGB, untransfered -------------------------------------------- */
      case 'gamma':
        // `a:gamma` *is* the sRGB de-linearisation, applied to the sRGB values
        // as they stand. The format exposing the curve directly is the strongest
        // independent confirmation that tint and shade use the same one.
        w.r = fromLinear(w.r);
        w.g = fromLinear(w.g);
        w.b = fromLinear(w.b);
        break;
      case 'invGamma':
        w.r = toLinear(w.r);
        w.g = toLinear(w.g);
        w.b = toLinear(w.b);
        break;
      case 'gray': {
        // Rec.709 luma of the sRGB values, with no linearisation. `#FF0000`
        // becomes `363636`; in linear light it would be `7F7F7F`.
        const y = luma709(w.r, w.g, w.b);
        w.r = y;
        w.g = y;
        w.b = y;
        break;
      }

      /* ---- HSL over the sRGB values --------------------------------------- */
      case 'hue':
        hsl(w, (c) => (c.h = wrapHue(t.val / ANGLE)));
        break;
      case 'hueOff':
        hsl(w, (c) => (c.h = wrapHue(c.h + t.val / ANGLE)));
        break;
      case 'hueMod':
        hsl(w, (c) => (c.h = wrapHue(c.h * (t.val / PERCENT))));
        break;
      case 'comp':
        // The complement as DrawingML means it: half a turn of hue, saturation
        // and lightness untouched. Identical to `max + min - c` per channel, and
        // nothing like `inv` above.
        hsl(w, (c) => (c.h = wrapHue(c.h + 180)));
        break;
      case 'sat':
        hsl(w, (c) => (c.s = t.val / PERCENT));
        break;
      case 'satOff':
        hsl(w, (c) => (c.s += t.val / PERCENT));
        break;
      case 'satMod':
        hsl(w, (c) => (c.s *= t.val / PERCENT));
        break;
      case 'lum':
        hsl(w, (c) => (c.l = t.val / PERCENT));
        break;
      case 'lumOff':
        hsl(w, (c) => (c.l += t.val / PERCENT));
        break;
      case 'lumMod':
        hsl(w, (c) => (c.l *= t.val / PERCENT));
        break;

      default: {
        // Unreachable through the exported types; reachable from a parser that
        // trusted a tag name. Typed, like every other failure in this repository.
        const op: string = (t as { op: string }).op;
        throw new PaintError('COLOR_TRANSFORM', `"${op}" is not a DrawingML colour transform`, op);
      }
    }
  }

  return { r: w.r, g: w.g, b: w.b, a: clamp01(w.a) };
}

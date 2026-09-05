/**
 * The candidate models for `a:tint` and `a:shade`, and for the HSL transforms.
 *
 * This is not `@pptx-studio/paint`. Phase 2.6 writes that, and it writes *one*
 * model - the one this file's measurement picks. What lives here is the set of
 * mutually exclusive hypotheses, kept side by side so the fixture can say which
 * one PowerPoint actually implements and by how much the others are wrong.
 *
 * ECMA-376 §20.1.2.3.34 (`tint`) and §20.1.2.3.31 (`shade`) describe the
 * *blend* - "a 10% tint is 10% of the input colour combined with 90% white" -
 * and are silent on the colour space it happens in. That silence is the whole
 * question.
 */

export type Channel = number; // 0..1

/* -------------------------------------------------------------------------- */
/* transfer functions                                                         */
/* -------------------------------------------------------------------------- */

/** LibreOffice's power law. `oox/source/drawingml/color.cxx`, DEC_GAMMA = 2.3. */
const GAMMA = 2.3;
const toLinearGamma = (c: Channel): Channel => Math.pow(c, GAMMA);
const fromLinearGamma = (c: Channel): Channel => Math.pow(c, 1 / GAMMA);

/** The sRGB piecewise transfer function, as Apache POI's DrawPaint uses. */
const toLinearSrgb = (c: Channel): Channel =>
  c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
const fromLinearSrgb = (c: Channel): Channel =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;

const identity = (c: Channel): Channel => c;

/* -------------------------------------------------------------------------- */
/* tint and shade                                                             */
/* -------------------------------------------------------------------------- */

export interface BlendModel {
  readonly id: string;
  readonly note: string;
  // `this: void` because these are pulled off the object and called bare.
  tint(this: void, c: Channel, p: number): Channel;
  shade(this: void, c: Channel, p: number): Channel;
}

function blendModel(
  id: string,
  note: string,
  to: (c: Channel) => Channel,
  from: (c: Channel) => Channel,
): BlendModel {
  return {
    id,
    note,
    // "a p tint is p of the input combined with (1-p) white"
    tint: (c, p) => from(to(c) * p + (1 - p)),
    // "a p shade is p of the input combined with (1-p) black"
    shade: (c, p) => from(to(c) * p),
  };
}

export const BLEND_MODELS: readonly BlendModel[] = [
  blendModel('srgb', 'arithmetic straight on the 0-255 values', identity, identity),
  blendModel('gamma2.3', 'power law, gamma 2.3 - LibreOffice', toLinearGamma, fromLinearGamma),
  blendModel('linear-srgb', 'sRGB piecewise linearisation', toLinearSrgb, fromLinearSrgb),
  {
    id: 'poi-asymmetric',
    note: 'tint in sRGB space, shade linearised - Apache POI DrawPaint',
    tint: (c, p) => c * p + (1 - p),
    shade: (c, p) => fromLinearSrgb(toLinearSrgb(c) * p),
  },
];

/* -------------------------------------------------------------------------- */
/* HSL                                                                        */
/* -------------------------------------------------------------------------- */

export interface Hsl {
  h: number; // 0..360
  s: number; // 0..1
  l: number; // 0..1
}

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

export function hslToRgb(hsl: Hsl): [Channel, Channel, Channel] {
  const { h, s, l } = hsl;
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number): number => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  const k = h / 360;
  return [hue(k + 1 / 3), hue(k), hue(k - 1 / 3)];
}

export interface HslModel {
  readonly id: string;
  readonly note: string;
  /** Apply an ordered chain of HSL-space transforms to an sRGB triple. */
  apply(
    this: void,
    rgb: readonly [Channel, Channel, Channel],
    ops: readonly HslOp[],
  ): [Channel, Channel, Channel];
}

export type HslOp =
  | { readonly kind: 'lumMod' | 'lumOff' | 'lum'; readonly p: number }
  | { readonly kind: 'satMod' | 'satOff' | 'sat'; readonly p: number }
  | { readonly kind: 'hueMod' | 'hueOff' | 'hue'; readonly p: number };

const clamp = (n: number): number => Math.min(1, Math.max(0, n));

function hslModel(
  id: string,
  note: string,
  to: (c: Channel) => Channel,
  from: (c: Channel) => Channel,
): HslModel {
  return {
    id,
    note,
    apply(rgb, ops) {
      const hsl = rgbToHsl(to(rgb[0]), to(rgb[1]), to(rgb[2]));
      for (const op of ops) {
        switch (op.kind) {
          case 'lumMod':
            hsl.l = hsl.l * op.p;
            break;
          case 'lumOff':
            hsl.l = hsl.l + op.p;
            break;
          case 'lum':
            hsl.l = op.p;
            break;
          case 'satMod':
            hsl.s = hsl.s * op.p;
            break;
          case 'satOff':
            hsl.s = hsl.s + op.p;
            break;
          case 'sat':
            hsl.s = op.p;
            break;
          case 'hueMod':
            hsl.h = (hsl.h * op.p) % 360;
            break;
          case 'hueOff':
            hsl.h = (((hsl.h + op.p) % 360) + 360) % 360;
            break;
          case 'hue':
            hsl.h = ((op.p % 360) + 360) % 360;
            break;
        }
      }
      // Saturation and lightness are deliberately NOT clamped above. Measured
      // against PowerPoint: `satMod val="200000"` and `satMod val="300000"` on
      // one base produce *different* colours, which a model that clamped S to 1
      // cannot express - it would give the same answer for both. The clamp
      // happens once, on the way out, per channel.
      // Clamp in the space the arithmetic happened in, then convert back - a
      // negative channel handed to a power law is NaN, not a dark colour.
      const out = hslToRgb(hsl);
      return [from(clamp(out[0])), from(clamp(out[1])), from(clamp(out[2]))];
    },
  };
}

export const HSL_MODELS: readonly HslModel[] = [
  hslModel('hsl-srgb', 'HSL directly on the sRGB values, no gamma', identity, identity),
  hslModel(
    'hsl-gamma2.3',
    'HSL after a gamma 2.3 linearisation - LibreOffice',
    toLinearGamma,
    fromLinearGamma,
  ),
  hslModel(
    'hsl-linear-srgb',
    'HSL after sRGB piecewise linearisation',
    toLinearSrgb,
    fromLinearSrgb,
  ),
];

/* -------------------------------------------------------------------------- */
/* helpers                                                                    */
/* -------------------------------------------------------------------------- */

export const to8 = (c: Channel): number => Math.max(0, Math.min(255, Math.round(c * 255)));
export const from8 = (n: number): Channel => n / 255;

export function parseHex(value: string): [Channel, Channel, Channel] {
  return [
    from8(parseInt(value.slice(0, 2), 16)),
    from8(parseInt(value.slice(2, 4), 16)),
    from8(parseInt(value.slice(4, 6), 16)),
  ];
}

export function toHex(rgb: readonly [Channel, Channel, Channel]): string {
  return rgb.map((c) => to8(c).toString(16).padStart(2, '0').toUpperCase()).join('');
}

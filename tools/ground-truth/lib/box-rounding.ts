/**
 * Every way a browser could turn a face box in font units into the pixels it
 * reports, so a fractional box can say which one it took.
 *
 * Blink's own two are here, half up and the Linux borrow from the ascent, so
 * each is scored rather than assumed. T13 on a 2048, a 2000 and a 2560 em;
 * ADR 0052, ADR 0053.
 */

/** One face box, in font units or in pixels, the descent positive below the baseline. */
export interface Box {
  readonly ascent: number;
  readonly descent: number;
}

/** The pixels a browser reports for a box in font units at `px`. */
export type BoxRounding = (box: Box, unitsPerEm: number, px: number) => Box;

const halfUp = (x: number): number => Math.floor(x + 0.5);

function halfEven(x: number): number {
  const whole = Math.floor(x);
  const fraction = x - whole;
  if (fraction < 0.5) return whole;
  if (fraction > 0.5) return whole + 1;
  return whole % 2 === 0 ? whole : whole + 1;
}

const exact = (box: Box, unitsPerEm: number, px: number): Box => ({
  ascent: (box.ascent * px) / unitsPerEm,
  descent: (box.descent * px) / unitsPerEm,
});

/** The em ratio held as a 32-bit float, which is where a half can lose its half. */
const single = (box: Box, unitsPerEm: number, px: number): Box => ({
  ascent: Math.fround(box.ascent / unitsPerEm) * px,
  descent: Math.fround(box.descent / unitsPerEm) * px,
});

const each =
  (round: (x: number) => number, scale: typeof exact = exact): BoxRounding =>
  (box, unitsPerEm, px) => {
    const at = scale(box, unitsPerEm, px);
    return { ascent: round(at.ascent), descent: round(at.descent) };
  };

export const BOX_ROUNDINGS: Readonly<Record<string, BoxRounding>> = {
  exact,
  'round half up': each(halfUp),
  'round half up, the em ratio in single precision': each(halfUp, single),
  'round half to even': each(halfEven),
  floor: each(Math.floor),
  ceil: each(Math.ceil),
  'round half up, borrowing one from the ascent where the descent rounded down': (
    box,
    unitsPerEm,
    px,
  ) => {
    const at = exact(box, unitsPerEm, px);
    let ascent = halfUp(at.ascent);
    let descent = halfUp(at.descent);
    if (descent < at.descent && ascent >= 1) {
      descent += 1;
      ascent -= 1;
    }
    return { ascent, descent };
  },
};

/** The rounding a fixture named, or a loud refusal to score under one it did not. */
export function boxRoundingNamed(name: string): BoxRounding {
  const rounding = BOX_ROUNDINGS[name];
  if (rounding === undefined)
    throw new Error(`no face box rounding is called ${JSON.stringify(name)}`);
  return rounding;
}

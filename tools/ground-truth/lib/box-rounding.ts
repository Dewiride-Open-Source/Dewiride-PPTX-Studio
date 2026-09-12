/**
 * Every way a browser could turn a face box in font units into the pixels it
 * reports, so a fractional box can say which one it took.
 *
 * Blink's own two are here, half up and the Linux borrow from the ascent, so
 * each is scored rather than assumed. T13 on a 2048 and a 2000 em, ADR 0052.
 */

/** One face box in pixels at the size it was read. */
export interface BoxPx {
  readonly ascent: number;
  readonly descent: number;
}

export type BoxRounding = (exact: BoxPx) => BoxPx;

const halfUp = (x: number): number => Math.floor(x + 0.5);

function halfEven(x: number): number {
  const whole = Math.floor(x);
  const fraction = x - whole;
  if (fraction < 0.5) return whole;
  if (fraction > 0.5) return whole + 1;
  return whole % 2 === 0 ? whole : whole + 1;
}

const each =
  (round: (x: number) => number): BoxRounding =>
  (box) => ({ ascent: round(box.ascent), descent: round(box.descent) });

export const BOX_ROUNDINGS: Readonly<Record<string, BoxRounding>> = {
  exact: (box) => box,
  'round half up': each(halfUp),
  'round half to even': each(halfEven),
  floor: each(Math.floor),
  ceil: each(Math.ceil),
  'round half up, borrowing one from the ascent where the descent rounded down': (box) => {
    let ascent = halfUp(box.ascent);
    let descent = halfUp(box.descent);
    if (descent < box.descent && ascent >= 1) {
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

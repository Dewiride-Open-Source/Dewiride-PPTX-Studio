/**
 * Turn the C3 fixture into the two tables `@pptx-studio/paint` ships.
 *
 * ```
 * node tools/ground-truth/write-paint-tables.ts
 * ```
 *
 * Writes `packages/paint/src/pattern-tiles.ts` and
 * `packages/paint/src/gradient-ramp.ts` from `corpus/ground-truth/fills.json`.
 * Neither table is hand-written and neither may be hand-edited: `fill.test.ts`
 * re-derives both from the same fixture and fails if they have drifted, which is
 * the only thing standing between a measured constant and somebody's memory of
 * one.
 */

import { readFileSync, writeFileSync } from 'node:fs';

interface Tile {
  period: [number, number] | null;
  rows: string[];
  soft: number;
}

interface Probe {
  id: string;
  group: string;
  strip?: string;
  tile?: Tile;
}

interface Fixture {
  stripSamples: number;
  blendGamma: number;
  probes: Probe[];
}

const fixture = JSON.parse(readFileSync('corpus/ground-truth/fills.json', 'utf8')) as Fixture;

const byId = new Map(fixture.probes.map((p) => [p.id, p]));

/* -------------------------------------------------------------------------- */
/* the two-stop ramp                                                          */
/* -------------------------------------------------------------------------- */

function channel(id: string): number[] {
  const packed = byId.get(id)?.strip;
  if (packed === undefined) throw new Error(`no strip for ${id}`);
  const out: number[] = [];
  for (let i = 0; i + 6 <= packed.length; i += 6) out.push(parseInt(packed.slice(i, i + 2), 16));
  return out;
}

const G = fixture.blendGamma;
const N = fixture.stripSamples;

/**
 * The blend weight at each sample, taken from whichever of the two ramps
 * resolves it. See the long note in `analyse-fills.ts`: inverting a 2.2 power on
 * an eight-bit export costs all the precision at whichever end is near 1.
 */
export function weightCurve(bw: number[], wb: number[], gamma: number): number[] {
  return bw.map((v, k) => {
    const a = Math.pow(v / 255, gamma);
    return a <= 0.5 ? a : 1 - Math.pow(wb[k]! / 255, gamma);
  });
}

/**
 * Choose the fewest knots whose chords stay within `tol` bytes of the curve, in
 * **both** directions.
 *
 * Two things about that sentence are load-bearing.
 *
 * The error is measured on the *byte* the ramp would paint, not on the weight,
 * because that is what a viewer sees and because linear interpolation between
 * two weights is a different curve from linear interpolation between the two
 * colours they imply - nineteen bytes different, on the steep end.
 *
 * And it is measured on a black-to-white ramp *and* a white-to-black one,
 * because the curve is steep at one end and a knot set fitted to one direction
 * leaves the other's steep end bare. Fitting only the ascending ramp gives ten
 * knots that are within 1.2 bytes forwards and thirteen backwards. Since any
 * other pair of stop colours has a smaller dynamic range than black and white,
 * bounding those two bounds every gradient there is.
 */
export function chooseKnots(weights: number[], gamma: number, tol: number): number[] {
  const clamp = (w: number): number => Math.min(1, Math.max(0, w));
  const up = (w: number): number => 255 * Math.pow(clamp(w), 1 / gamma);
  const down = (w: number): number => 255 * Math.pow(clamp(1 - w), 1 / gamma);
  const knots = [0, weights.length - 1];
  for (;;) {
    let worst = 0;
    let at = -1;
    let seg = -1;
    for (let s = 0; s < knots.length - 1; s++) {
      const a = knots[s]!;
      const b = knots[s + 1]!;
      for (let i = a + 1; i < b; i++) {
        const u = (i - a) / (b - a);
        const e = Math.max(
          ...[up, down].map((f) => {
            const chord = f(weights[a]!) + (f(weights[b]!) - f(weights[a]!)) * u;
            return Math.abs(chord - f(weights[i]!));
          }),
        );
        if (e > worst) {
          worst = e;
          at = i;
          seg = s;
        }
      }
    }
    if (worst <= tol || at < 0) return knots;
    knots.splice(seg + 1, 0, at);
  }
}

const weights = weightCurve(channel('ramp-bw'), channel('ramp-wb'), G);
// Two bytes. Fifteen knots get there; twenty-eight are needed for 1.5, and buy
// four tenths of a byte for nearly twice the markup in every gradient we emit.
const knots = chooseKnots(weights, G, 2).map((i) => [i / (N - 1), weights[i]!] as [number, number]);

const rampFile = `/**
 * The curve PowerPoint draws between two gradient stops - and the fact that it
 * only draws it sometimes.
 *
 * ## What is measured here
 *
 * A \`a:gradFill\` with **two distinct colours and a stop at each end** is not
 * interpolated linearly. A black-to-white ramp paints \`BABABA\` at its
 * midpoint, not \`808080\`. Anything else - three distinct colours, or two
 * colours whose stops do not reach 0 and 100% - is exactly linear in sRGB:
 * measured at worst 0.8 of a byte over 1920 samples, against 59 for the curve.
 *
 * Note "two colours", not "two stops". PowerPoint's own gallery writes variant 4
 * of every gradient style as three stops in two colours, and that *is* curved;
 * three stops in three colours at the same three positions is linear. See
 * \`twoColorRamp\` in \`gradient.ts\` for why, and for the mechanism that
 * predicts it.
 *
 * The blend space for the curved case is a plain 2.2 power, not the sRGB
 * piecewise transfer and not linear light. One weight curve, applied in that
 * space, reproduces five independently measured colour pairs - white to black,
 * two greys, red to green, blue to yellow, blue to red - to within one byte per
 * channel. That is what makes it a model rather than a table of one ramp.
 *
 * ## What is not known
 *
 * No closed form fits. A power law, a smoothstep, a logistic power and the sRGB
 * and linear-light transfers were all fitted over 1920 samples; the best is
 * eight bytes out at its worst. The plan predicted a "1.875 gamma ramp", and
 * \`1 - (1 - t)^1.875\` is indeed the closest single-parameter fit to the middle
 * of the range, but it misses the dark end by thirteen. So the curve ships as
 * what it is: measurement, not derivation.
 *
 * ## Why fifteen knots, and why they are not evenly spaced
 *
 * SVG interpolates between stops linearly in sRGB, and \`color-interpolation\`
 * does not apply to gradients in any shipping browser, so the only way to draw
 * this curve is to pre-sample it. Fifteen stops, adaptively placed, hold both
 * measured ramps - black to white and white to black, the two widest there are -
 * to within 1.9 of 255. Uniform spacing needs sixty-five to do worse, because
 * almost all the curvature is in the first twentieth of the ramp.
 *
 * The knots are fitted to **both** directions at once. A set fitted only to the
 * ascending ramp puts its knots where that ramp is steep and leaves the
 * descending ramp's steep end bare: ten such knots are within 1.2 bytes forwards
 * and thirteen backwards, which is a bug that only shows on gradients that run
 * light to dark.
 *
 * GENERATED by tools/ground-truth/write-paint-tables.ts from
 * corpus/ground-truth/fills.json. Do not edit: fill.test.ts re-derives it.
 */

/** The space the two-stop curve blends in: a plain 2.2 power, measured. */
export const RAMP_BLEND_GAMMA = ${String(G)};

/**
 * \`[position, weight]\` pairs, position and weight both 0 to 1.
 *
 * \`weight\` is the fraction of the *second* stop's colour, in the 2.2 power
 * space - so the colour at \`position\` is
 * \`fromRampSpace(a + (b - a) * weight)\` per channel, where \`a\` and \`b\` are
 * the two stop colours in that space.
 */
export const TWO_STOP_RAMP: readonly (readonly [number, number])[] = [
${knots.map(([p, w]) => `  [${p.toFixed(6)}, ${w.toFixed(6)}],`).join('\n')}
];

/** Into the blend space. */
export function toRampSpace(channel: number): number {
  return Math.pow(channel < 0 ? 0 : channel > 1 ? 1 : channel, RAMP_BLEND_GAMMA);
}

/** Out of the blend space. */
export function fromRampSpace(value: number): number {
  return Math.pow(value < 0 ? 0 : value > 1 ? 1 : value, 1 / RAMP_BLEND_GAMMA);
}

/**
 * The blend weight at \`t\`.
 *
 * Between two knots this interpolates in the **display** domain, not in the
 * weight domain, because that is what a viewer does with the stops we emit: SVG
 * and CSS both walk linearly from one stop's colour to the next. Interpolating
 * the weight instead is a different curve, and on the steep dark end of the ramp
 * it is nineteen bytes away from what PowerPoint paints; this way the worst
 * disagreement over the whole ramp is 1.6.
 *
 * So the ten knots are the model, and everything between them is defined to be
 * whatever a linear interpolation of the knot colours gives - which is the only
 * honest definition, since it is the only thing we can actually cause a browser
 * to draw.
 */
export function twoStopWeight(t: number): number {
  if (t <= 0) return TWO_STOP_RAMP[0]![1];
  const last = TWO_STOP_RAMP[TWO_STOP_RAMP.length - 1]!;
  if (t >= 1) return last[1];
  for (let i = 1; i < TWO_STOP_RAMP.length; i++) {
    const [pb, wb] = TWO_STOP_RAMP[i]!;
    if (t <= pb) {
      const [pa, wa] = TWO_STOP_RAMP[i - 1]!;
      if (pb === pa) return wb;
      const u = (t - pa) / (pb - pa);
      const da = fromRampSpace(wa);
      const db = fromRampSpace(wb);
      return toRampSpace(da + (db - da) * u);
    }
  }
  return last[1];
}
`;

writeFileSync('packages/paint/src/gradient-ramp.ts', rampFile);

/* -------------------------------------------------------------------------- */
/* the 54 tiles                                                               */
/* -------------------------------------------------------------------------- */

const patterns = fixture.probes.filter((p) => p.group === 'pattern');
const entries = patterns.map((p) => {
  const tile = p.tile;
  if (!tile?.period) throw new Error(`no tile for ${p.id}`);
  return {
    name: p.id.slice('pat-'.length),
    w: tile.period[0],
    h: tile.period[1],
    cover: tile.rows.join(''),
    soft: tile.soft > 0,
  };
});

const tileFile = `/**
 * The 54 \`ST_PresetPatternVal\` tiles, measured out of PowerPoint's own export.
 *
 * ## Where these come from
 *
 * Not from Mono libgdiplus, which the plan proposed and which would have needed
 * a download. PowerPoint was asked to author one shape per \`MsoPatternType\`
 * through its object model; it accepted exactly 54 and spelled each name into
 * \`a:pattFill/@prst\` itself, so the enumeration below is PowerPoint's rather
 * than a transcription of anyone's. The pixels were then read back from a slide
 * export at 1280 x 720 - which over a 13.333 inch slide is exactly 96 DPI, the
 * one resolution at which the tile is 1:1 and its bits are not resampled.
 *
 * ## The tile is a physical size, not a shape fraction and not device pixels
 *
 * Eight pixels at 96 DPI: one twelfth of an inch, six points, 76200 EMU. The
 * same shape exported at 640, 1280, 1920 and 2560 pixels wide gives periods of
 * 4, 8, 12 and 16 pixels, so the tile scales with the *rendering resolution* and
 * not with the shape. In SVG that is \`patternUnits="userSpaceOnUse"\` with a
 * fixed size in slide units - never \`objectBoundingBox\`, which would stretch
 * the hatching of a wide shape into slanted mush.
 *
 * ## Coverage, not bits
 *
 * Fifty-one of the 54 are one-bit. \`dnDiag\`, \`upDiag\` and \`diagCross\` are
 * drawn as antialiased diagonal lines and their tiles contain intermediate
 * values, so every tile is stored as coverage. \`lgGrid\` and \`cross\` measure
 * byte-identical, which is the only such pair among the 54 and matches the GDI+
 * heritage where \`LargeGrid\` is an alias of \`Cross\`.
 *
 * The \`pctNN\` names are labels rather than measurements: \`pct75\` covers 88%
 * of its tile and \`pct20\` covers 13%.
 *
 * GENERATED by tools/ground-truth/write-paint-tables.ts from
 * corpus/ground-truth/fills.json. Do not edit: fill.test.ts re-derives it.
 */

/** One tile, at its minimal repeating period. */
export interface PatternTile {
  /** Width of the repeat, in pixels at 96 DPI. */
  readonly w: number;
  /** Height of the repeat, in pixels at 96 DPI. */
  readonly h: number;
  /**
   * Row-major coverage, two hex digits per pixel: \`00\` is the background
   * colour, \`FF\` the foreground, and anything between is antialiasing.
   */
  readonly cover: string;
}

/** One tile pixel, in EMU. 914400 / 96. */
export const PATTERN_PIXEL_EMU = 9525;

/** The full tile is always eight of those square, whatever the minimal period is. */
export const PATTERN_TILE_PIXELS = 8;

export const PATTERN_TILES: Readonly<Record<string, PatternTile>> = {
${entries.map((e) => `  ${e.name}: { w: ${String(e.w)}, h: ${String(e.h)}, cover: '${e.cover}' },`).join('\n')}
};

/** The 54 names, in \`MsoPatternType\` order - the order PowerPoint enumerated them. */
export const PRESET_PATTERN_NAMES: readonly string[] = [
${entries.map((e) => `  '${e.name}',`).join('\n')}
];

/** The three tiles PowerPoint antialiases rather than drawing one-bit. */
export const ANTIALIASED_PATTERNS: readonly string[] = [
${entries
  .filter((e) => e.soft)
  .map((e) => `  '${e.name}',`)
  .join('\n')}
];

/**
 * Is this one of the 54?
 *
 * Deliberately a plain boolean and not a type predicate: \`PATTERN_TILES\` is
 * keyed by \`string\`, so a predicate would narrow the *failing* branch to
 * \`never\` and make the error path unwritable.
 */
export function isPresetPatternName(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(PATTERN_TILES, name);
}
`;

writeFileSync('packages/paint/src/pattern-tiles.ts', tileFile);

console.log(
  `gradient-ramp.ts: ${String(knots.length)} knots; pattern-tiles.ts: ${String(entries.length)} tiles`,
);

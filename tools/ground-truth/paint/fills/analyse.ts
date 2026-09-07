/**
 * Experiment C3, step 3 - read the bitmaps and say what PowerPoint did.
 *
 * ```
 * node tools/ground-truth/paint/fills/analyse.ts <work-dir> [--fixture corpus/ground-truth/fills.json]
 * ```
 *
 * Each section states its hypothesis before it prints its evidence, so a reader
 * can disagree with the question rather than only with the answer. The fixture
 * it writes is the committed measurement; everything printed here is derived
 * from it and nothing in the package depends on this file having been run.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { hex, readBmp, type Bitmap } from '../../lib/bmp.ts';

/* -------------------------------------------------------------------------- */
/* inputs                                                                     */
/* -------------------------------------------------------------------------- */

const arg = process.argv[2];
if (arg === undefined)
  throw new Error('usage: tools/ground-truth/paint/fills/analyse.ts <work-dir> [--fixture <path>]');
// Narrowed once, here: `bitmap` below is a hoisted function declaration and does
// not see the narrowing of a module-level `const` that a throw established.
const dir: string = arg;
const fixtureFlag = process.argv.indexOf('--fixture');
const fixturePath = fixtureFlag === -1 ? null : process.argv[fixtureFlag + 1];

interface ProbeInput {
  id: string;
  deck: string;
  slide: number;
  index: number;
  group: string;
  question: string;
  sample: 'strip' | 'grid' | 'tile' | 'corner';
  fill: string;
  prst: string;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  hasBackdrop: boolean;
  x: number;
  y: number;
  cx: number;
  cy: number;
}

interface Inputs {
  clrScheme: Record<string, string>;
  slideSize: { cx: number; cy: number };
  decks: { deck: string; file: string; slides: number; probes: number }[];
  probes: ProbeInput[];
}

interface ComStop {
  index: number;
  position: number;
  rgb: number;
  transparency: number;
}

interface ComShape {
  id: string;
  slide: number;
  fillType: number | null;
  foreRgb: number | null;
  backRgb: number | null;
  transparency: number | null;
  pattern: number | null;
  gradStyle: number | null;
  gradAngle: number | null;
  stops: ComStop[];
}

interface DeckReadback {
  deck: string;
  opened: boolean;
  repaired: boolean | null;
  error: string | null;
  shapes: ComShape[];
  bitmaps: { file: string; slide: number; width: number; height: number }[];
}

const inputs = JSON.parse(readFileSync(join(dir, 'fill-inputs.json'), 'utf8')) as Inputs;
const readback = JSON.parse(readFileSync(join(dir, 'com-readback-fills.json'), 'utf8')) as {
  decks: DeckReadback[];
};

const deckOf = new Map(readback.decks.map((d) => [d.deck, d]));
const bitmapCache = new Map<string, Bitmap>();

function bitmap(file: string): Bitmap {
  const cached = bitmapCache.get(file);
  if (cached) return cached;
  const bmp = readBmp(readFileSync(join(dir, file)));
  bitmapCache.set(file, bmp);
  return bmp;
}

/** The probe's rectangle in the pixels of one export, exactly - everything is point-aligned. */
function pixelRect(
  probe: ProbeInput,
  width: number,
  height: number,
): { x0: number; y0: number; x1: number; y1: number } {
  const sx = width / inputs.slideSize.cx;
  const sy = height / inputs.slideSize.cy;
  return {
    x0: Math.round(probe.x * sx),
    y0: Math.round(probe.y * sy),
    x1: Math.round((probe.x + probe.cx) * sx),
    y1: Math.round((probe.y + probe.cy) * sy),
  };
}

/* -------------------------------------------------------------------------- */
/* colour helpers                                                             */
/* -------------------------------------------------------------------------- */

/**
 * The blend space, measured.
 *
 * Not the sRGB piecewise curve, and not linear light: a plain 2.2 power. The
 * three ramps that start and end on colours other than black and white -
 * `ramp-mid`, `ramp-rg`, `ramp-by` - are reproduced to the byte by a 2.2 power
 * and are one to three bytes out under the piecewise sRGB transfer, which is
 * enough to tell them apart because the ramps are 1920 samples long.
 */
const BLEND_GAMMA = 2.2;
const toBlend = (c: number): number => Math.pow(c, BLEND_GAMMA);
const fromBlend = (c: number): number => Math.pow(Math.min(1, Math.max(0, c)), 1 / BLEND_GAMMA);

/* -------------------------------------------------------------------------- */
/* sampling                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * Samples along the horizontal centre line of a ramp strip.
 *
 * 481 across a 1920 px export is every fourth pixel, and it is set by the dark
 * end of the two-stop curve rather than by the rest of it: the curve turns
 * sharply somewhere inside the first fifteen pixels, and at 129 samples the
 * first three land at 0, 15 and 30 px, which straddles the turn and smooths it
 * away. A table fitted to that smoothed curve is three bytes out against the
 * real export; fitted to this one, under one.
 */
const STRIP_N = 481;
/** A 13 x 13 lattice over the interior. Enough to fit a direction or read a path. */
const GRID_N = 13;
/** Pixels either side of a profile corner, at 1:1. Odd, so one lands on it. */
const CORNER_N = 257;

function sampleStrip(bmp: Bitmap, probe: ProbeInput): string {
  const { x0, y0, x1, y1 } = pixelRect(probe, bmp.width, bmp.height);
  const ym = Math.floor((y0 + y1) / 2);
  const out: string[] = [];
  for (let k = 0; k < STRIP_N; k++) {
    const x = Math.min(x1 - 1, x0 + Math.round((k / (STRIP_N - 1)) * (x1 - x0 - 1)));
    out.push(hex(bmp.pixel(x, ym)));
  }
  return out.join('');
}

/** `CORNER_N` pixels at 1:1, centred on the shape's horizontal midpoint. */
function sampleCorner(bmp: Bitmap, probe: ProbeInput): string {
  const { x0, y0, x1, y1 } = pixelRect(probe, bmp.width, bmp.height);
  const ym = Math.floor((y0 + y1) / 2);
  const xm = Math.floor((x0 + x1) / 2);
  const half = (CORNER_N - 1) / 2;
  const out: string[] = [];
  for (let k = -half; k <= half; k++) {
    out.push(hex(bmp.pixel(Math.min(x1 - 1, Math.max(x0, xm + k)), ym)));
  }
  return out.join('');
}

function sampleGrid(bmp: Bitmap, probe: ProbeInput): string {
  const { x0, y0, x1, y1 } = pixelRect(probe, bmp.width, bmp.height);
  // One pixel in from each edge: a rotated shape antialiases against the slide,
  // and a sample that is half background measures the background.
  const ax = x0 + 1;
  const ay = y0 + 1;
  const bx = x1 - 2;
  const by = y1 - 2;
  const out: string[] = [];
  for (let r = 0; r < GRID_N; r++) {
    for (let c = 0; c < GRID_N; c++) {
      const x = ax + Math.round((c / (GRID_N - 1)) * (bx - ax));
      const y = ay + Math.round((r / (GRID_N - 1)) * (by - ay));
      out.push(hex(bmp.pixel(x, y)));
    }
  }
  return out.join('');
}

interface Tile {
  /** Period in pixels, or null when nothing under 64 repeats. */
  period: [number, number] | null;
  /**
   * One row per `period[1]`, each two hex digits per pixel: `00` is pure
   * background, `FF` pure foreground.
   *
   * Coverage rather than a bit, because three of the fifty-four are not
   * one-bit. `dnDiag`, `upDiag` and `diagCross` are drawn as antialiased
   * diagonal lines and their tiles contain greys; a renderer that thresholds
   * them to a bitmap draws visibly harsher hatching than PowerPoint does.
   */
  rows: string[];
  /** How many pixels in the tile are strictly between the two colours. */
  soft: number;
  fg: string;
  bg: string;
}

/**
 * Recover a pattern tile from the pixels.
 *
 * The period is *found*, not assumed: the whole point of measuring is that
 * "every DrawingML pattern is an 8x8 tile" is a claim, and `wave` and `weave`
 * are the ones most likely to break it.
 */
function sampleTile(bmp: Bitmap, probe: ProbeInput): Tile {
  const { x0, y0, x1, y1 } = pixelRect(probe, bmp.width, bmp.height);
  const w = Math.min(96, x1 - x0 - 2);
  const h = Math.min(96, y1 - y0 - 2);
  const px: [number, number, number][][] = [];
  for (let r = 0; r < h; r++) {
    const row: [number, number, number][] = [];
    for (let c = 0; c < w; c++) row.push(bmp.pixel(x0 + 1 + c, y0 + 1 + r));
    px.push(row);
  }

  // The two ends of the tile are its darkest and lightest pixels, which keeps
  // this working for the themed and translucent pattern probes as well as for
  // the plain black-on-white ones.
  const luma = (p: [number, number, number]): number =>
    0.2126 * p[0] + 0.7152 * p[1] + 0.0722 * p[2];
  let dark = px[0]![0]!;
  let light = px[0]![0]!;
  for (const row of px) {
    for (const p of row) {
      if (luma(p) < luma(dark)) dark = p;
      if (luma(p) > luma(light)) light = p;
    }
  }
  // Which end is the foreground comes from the markup, not from the pixels.
  //
  // The tempting heuristic - "the minority colour is the ink" - is wrong for
  // exactly the patterns it matters for: `pct70`, `pct75`, `pct80` and `pct90`
  // are more ink than ground, so the heuristic inverts them, and an inverted
  // `pct75` is bit-identical to `pct20`. That collision is what caught it.
  const declaredFg = /<a:fgClr><a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(probe.fill)?.[1];
  const near = (want: string, a: [number, number, number], b: [number, number, number]) => {
    const target = [0, 1, 2].map((i) => parseInt(want.slice(i * 2, i * 2 + 2), 16));
    const d = (p: [number, number, number]) =>
      Math.abs(p[0] - target[0]!) + Math.abs(p[1] - target[1]!) + Math.abs(p[2] - target[2]!);
    return d(a) <= d(b) ? a : b;
  };
  let fgPx: [number, number, number];
  if (declaredFg !== undefined) {
    fgPx = near(declaredFg, dark, light);
  } else {
    let atDark = 0;
    for (const row of px)
      for (const p of row) if (luma(p) < (luma(dark) + luma(light)) / 2) atDark++;
    fgPx = atDark * 2 <= w * h ? dark : light;
  }
  const bgPx = fgPx === dark ? light : dark;

  const span = luma(fgPx) - luma(bgPx);
  /** Coverage: 0 at the background colour, 255 at the foreground colour. */
  const cover = (p: [number, number, number]): number => {
    if (span === 0) return 0;
    const a = (luma(p) - luma(bgPx)) / span;
    return Math.round(Math.min(1, Math.max(0, a)) * 255);
  };
  const alpha: number[][] = px.map((row) => row.map(cover));

  let soft = 0;
  for (const row of alpha) for (const a of row) if (a !== 0 && a !== 255) soft++;

  const findPeriod = (axis: 'x' | 'y'): number | null => {
    const limit = axis === 'x' ? w : h;
    for (let n = 1; n <= Math.min(64, Math.floor(limit / 2)); n++) {
      let ok = true;
      for (let r = 0; r < h && ok; r++) {
        for (let c = 0; c < w; c++) {
          const r2 = axis === 'y' ? r + n : r;
          const c2 = axis === 'x' ? c + n : c;
          if (r2 >= h || c2 >= w) continue;
          // One unit of slack: the export quantises, and an exact-equality test
          // reports "no period at all" for a tile that is plainly periodic.
          if (Math.abs(alpha[r]![c]! - alpha[r2]![c2]!) > 1) {
            ok = false;
            break;
          }
        }
      }
      if (ok) return n;
    }
    return null;
  };

  const byte = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
  const pw = findPeriod('x');
  const ph = findPeriod('y');
  if (pw === null || ph === null) {
    return { period: null, rows: [], soft, fg: hex(fgPx), bg: hex(bgPx) };
  }
  const rows: string[] = [];
  for (let r = 0; r < ph; r++) {
    rows.push(
      alpha[r]!.slice(0, pw)
        .map((a) => byte(a))
        .join(''),
    );
  }
  return { period: [pw, ph], rows, soft, fg: hex(fgPx), bg: hex(bgPx) };
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface FixtureProbe extends ProbeInput {
  opened: boolean;
  repaired: boolean | null;
  openError: string | null;
  /** `STRIP_N` samples, six hex digits each, concatenated. */
  strip?: string;
  /** `GRID_N * GRID_N` samples in row-major order, six hex digits each. */
  grid?: string;
  /** Export width to `CORNER_N` samples across the profile corner. */
  corner?: Record<string, string>;
  tile?: Tile;
  tileByWidth?: Record<string, [number, number] | null>;
  com?: {
    fillType: number | null;
    foreRgb: string | null;
    backRgb: string | null;
    transparency: number | null;
    pattern: number | null;
    stops: { position: number; rgb: string; transparency: number }[];
  };
}

const bgr = (n: number | null): string | null =>
  n === null
    ? null
    : hex([n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff] as [number, number, number]);

const fixtureProbes: FixtureProbe[] = [];

for (const probe of inputs.probes) {
  const deck = deckOf.get(probe.deck);
  const entry: FixtureProbe = {
    ...probe,
    opened: deck?.opened ?? false,
    repaired: deck?.repaired ?? null,
    openError: deck?.error ?? null,
  };

  const com = deck?.shapes.find((s) => s.id === probe.id);
  if (com) {
    entry.com = {
      fillType: com.fillType,
      foreRgb: bgr(com.foreRgb),
      backRgb: bgr(com.backRgb),
      transparency: com.transparency,
      pattern: com.pattern,
      stops: com.stops.map((s) => ({
        position: s.position,
        rgb: bgr(s.rgb) ?? '000000',
        transparency: s.transparency,
      })),
    };
  }

  const shots = (deck?.bitmaps ?? []).filter((b) => b.slide === probe.slide);
  const main = shots.find((b) => b.width === 1920) ?? shots[0];
  // A tile is read at 96 DPI and nowhere else. 1280 px across a 13.333 in slide
  // is 1:1 with the tile's own resolution, so the bits come out exact; every
  // other width resamples them and turns each diagonal into a grey edge.
  const forTile = shots.find((b) => b.width === 1280) ?? main;
  if (main) {
    const bmp = bitmap(main.file);
    if (probe.sample === 'strip') entry.strip = sampleStrip(bmp, probe);
    else if (probe.sample === 'corner') {
      entry.corner = {};
      for (const shot of shots)
        entry.corner[String(shot.width)] = sampleCorner(bitmap(shot.file), probe);
    } else if (probe.sample === 'grid') entry.grid = sampleGrid(bmp, probe);
    else if (forTile) {
      entry.tile = sampleTile(bitmap(forTile.file), probe);
      if (shots.length > 1) {
        // Only the period at each resolution, not the pixels. What the other
        // exports are for is the question of whether the tile is a fixed number
        // of device pixels, and the period alone answers it.
        const byWidth: Record<string, [number, number] | null> = {};
        for (const shot of shots) {
          byWidth[String(shot.width)] = sampleTile(bitmap(shot.file), probe).period;
        }
        entry.tileByWidth = byWidth;
      }
    }
  }
  fixtureProbes.push(entry);
}

/* -------------------------------------------------------------------------- */
/* 1 - which packages PowerPoint would open                                   */
/* -------------------------------------------------------------------------- */

console.log('1. Which packages PowerPoint opened without repairing them.');
console.log('   Hypothesis: the schema types are enforced, so an out-of-range value is refused.');
for (const d of readback.decks) {
  const state = !d.opened ? 'REFUSED' : d.repaired === true ? 'REPAIRED' : 'ok';
  if (state === 'ok' && !d.deck.startsWith('h-')) continue;
  const probe = inputs.probes.find((p) => p.deck === d.deck);
  console.log(`   ${d.deck.padEnd(14)} ${state.padEnd(9)} ${probe?.question ?? ''}`);
}

/* -------------------------------------------------------------------------- */
/* 2 - the two-stop ramp                                                      */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('2. The curve joining two stops.');
console.log('   Hypothesis: sRGB-linear gives 128 at the midpoint, linear light 188.');

/** Unpack a run of six-hex-digit samples. */
function samples(packed: string): number[][] {
  const out: number[][] = [];
  for (let i = 0; i + 6 <= packed.length; i += 6) {
    out.push([
      parseInt(packed.slice(i, i + 2), 16),
      parseInt(packed.slice(i + 2, i + 4), 16),
      parseInt(packed.slice(i + 4, i + 6), 16),
    ]);
  }
  return out;
}

function stripOf(id: string): number[][] {
  const p = fixtureProbes.find((q) => q.id === id);
  if (p?.strip === undefined) throw new Error(`no strip for ${id}`);
  return samples(p.strip);
}

function gridOf(id: string): number[][] | null {
  const p = fixtureProbes.find((q) => q.id === id);
  return p?.grid === undefined ? null : samples(p.grid);
}

const bwStrip = stripOf('ramp-bw');
const mid = bwStrip[(STRIP_N - 1) / 2]![0]!;
console.log(
  `   ramp-bw midpoint = ${String(mid)}  (${hex(bwStrip[(STRIP_N - 1) / 2] as [number, number, number])})`,
);

/**
 * The measured weight at each of the 129 sample positions, in the 2.2 blend space.
 *
 * Read from *two* ramps, because one is not enough. Recovering the weight from
 * the black-to-white ramp means inverting a 2.2 power, and near the white end
 * one byte of the export is 0.009 of weight - so a renderer that used this
 * weight to paint a channel running the other way would be fourteen bytes out at
 * the far end, which is a fact about the measurement rather than about
 * PowerPoint. The white-to-black ramp has its fine resolution exactly where the
 * other has none, so each half of the range is taken from the ramp that resolves
 * it.
 */
const wbStrip = stripOf('ramp-wb');
const WEIGHT = bwStrip.map((px, k) => {
  const fromBw = toBlend(px[0]! / 255);
  const fromWb = 1 - toBlend(wbStrip[k]![0]! / 255);
  return fromBw <= 0.5 ? fromBw : fromWb;
});
console.log(
  `   weight at 1/4, 1/2, 3/4 = ${WEIGHT[32]!.toFixed(4)}, ${WEIGHT[64]!.toFixed(4)}, ${WEIGHT[96]!.toFixed(4)}`,
);
console.log('   Linear would be 0.2500, 0.5000, 0.7500. It is symmetric, and it is not linear.');

// Cross-validate: predict every other two-stop 0/100 ramp from this one weight.
function predict(c0: [number, number, number], c1: [number, number, number], w: number): number[] {
  return [0, 1, 2].map((i) => {
    const a = toBlend(c0[i]! / 255);
    const b = toBlend(c1[i]! / 255);
    return Math.round(255 * fromBlend(a + (b - a) * w));
  });
}

console.log('   Cross-validation - the same weight, applied to the other two-stop ramps:');
const CROSS: [string, [number, number, number], [number, number, number]][] = [
  ['ramp-wb', [255, 255, 255], [0, 0, 0]],
  ['ramp-mid', [0x40, 0x40, 0x40], [0xc0, 0xc0, 0xc0]],
  ['ramp-rg', [0xff, 0, 0], [0, 0xff, 0]],
  ['ramp-by', [0, 0, 0xff], [0xff, 0xff, 0]],
  ['ramp-hue', [0, 0, 0xff], [0xff, 0, 0]],
];
for (const [id, c0, c1] of CROSS) {
  const actual = stripOf(id);
  let worst = 0;
  for (let k = 0; k < STRIP_N; k++) {
    const p = predict(c0, c1, WEIGHT[k]!);
    for (let i = 0; i < 3; i++) worst = Math.max(worst, Math.abs(p[i]! - actual[k]![i]!));
  }
  console.log(`     ${id.padEnd(10)} worst channel error ${String(worst).padStart(3)}/255`);
}

/* -------------------------------------------------------------------------- */
/* 3 - and the ramps that are not special-cased                               */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('3. Which gradients get that curve.');
console.log('   Hypothesis: only a two-stop gradient whose stops are at 0 and 100%.');
for (const [id, note] of [
  ['ramp-3stop', 'three stops, black / 808080 / white'],
  ['ramp-inset', 'two stops, at 20% and 80%'],
] as const) {
  const s = stripOf(id);
  const a = s[0]![0]!;
  const b = s[STRIP_N - 1]![0]!;
  let worstLinear = 0;
  let worstCurve = 0;
  for (let k = 0; k < STRIP_N; k++) {
    const t = k / (STRIP_N - 1);
    // Position within the stop list, clamped outside it.
    const u = id === 'ramp-inset' ? Math.min(1, Math.max(0, (t - 0.2) / 0.6)) : t;
    const linear = a + (b - a) * u;
    worstLinear = Math.max(worstLinear, Math.abs(linear - s[k]![0]!));
    const curved =
      255 *
      fromBlend(
        toBlend(a / 255) +
          (toBlend(b / 255) - toBlend(a / 255)) * WEIGHT[Math.round(u * (STRIP_N - 1))]!,
      );
    worstCurve = Math.max(worstCurve, Math.abs(curved - s[k]![0]!));
  }
  console.log(
    `   ${id.padEnd(11)} ${note.padEnd(34)} linear-in-sRGB worst ${worstLinear.toFixed(1)}, curved worst ${worstCurve.toFixed(1)}`,
  );
}

/* -------------------------------------------------------------------------- */
/* 4 - stop order                                                             */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('4. Are stops read in document order, or sorted by @pos?');
console.log('   Hypothesis: sorted. PowerPoint writes them unsorted in its own from-centre fill.');
const rev = stripOf('ramp-revorder');
console.log(
  `   ramp-revorder (written 50%, 0%, 100%): ends ${hex(rev[0] as [number, number, number])} / ${hex(rev[STRIP_N - 1] as [number, number, number])}, middle ${hex(rev[(STRIP_N - 1) / 2] as [number, number, number])}`,
);
console.log('   Sorted predicts white / white with black in the middle.');
const revCom = fixtureProbes.find((p) => p.id === 'ramp-revorder')?.com;
if (revCom) {
  console.log(
    `   COM reports the stops as: ${revCom.stops.map((s) => `${String(Math.round(s.position * 1000) / 1000)}=${s.rgb}`).join(', ')}`,
  );
}

/* -------------------------------------------------------------------------- */
/* 5 - the linear gradient's direction                                        */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('5. @ang and @scaled.');
console.log('   Hypothesis: @ang is clockwise from +x with y down; @scaled=1 skews by the aspect.');

/** Fit the direction of steepest increase over the 13x13 grid, in degrees clockwise from +x. */
function fitDirection(id: string): { deg: number; residual: number } | null {
  const p = fixtureProbes.find((q) => q.id === id);
  const g = gridOf(id);
  if (p === undefined || g === null) return null;
  const v = g.map((px) => px[0]!);
  let sx = 0;
  let sy = 0;
  for (let r = 0; r < GRID_N; r++) {
    for (let c = 0; c < GRID_N - 1; c++) sx += v[r * GRID_N + c + 1]! - v[r * GRID_N + c]!;
  }
  for (let r = 0; r < GRID_N - 1; r++) {
    for (let c = 0; c < GRID_N; c++) sy += v[(r + 1) * GRID_N + c]! - v[r * GRID_N + c]!;
  }
  // In *shape* space, so a wide shape's 13 columns are further apart than its rows.
  const w = p.cx;
  const h = p.cy;
  const gx = sx / w;
  const gy = sy / h;
  const deg = (Math.atan2(gy, gx) * 180) / Math.PI;
  return { deg: (deg + 360) % 360, residual: Math.hypot(gx, gy) };
}

for (const p of fixtureProbes.filter((q) => q.group === 'angle')) {
  const fit = fitDirection(p.id);
  const declared = /ang="(-?\d+)"/.exec(p.fill)?.[1];
  const scaled = /scaled="(\d)"/.exec(p.fill)?.[1];
  console.log(
    `   ${p.id.padEnd(18)} ang=${String(declared).padStart(8)} scaled=${String(scaled)} shape ${String(p.cx / p.cy).slice(0, 5)}:1  measured ${fit ? fit.deg.toFixed(1) : '?'} deg`,
  );
}

/* -------------------------------------------------------------------------- */
/* 6 - path gradients                                                          */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('6. Path gradients.');
console.log(
  '   Hypothesis: stop 0 sits at the fillToRect focus and the last stop at the tile edge.',
);
for (const p of fixtureProbes.filter((q) => q.group === 'path')) {
  const g = gridOf(p.id);
  if (g === null) continue;
  const v = g.map((px) => px[0]!);
  const centre = v[Math.floor((GRID_N * GRID_N) / 2)]!;
  const corners = [v[0]!, v[GRID_N - 1]!, v[GRID_N * (GRID_N - 1)]!, v[GRID_N * GRID_N - 1]!];
  console.log(
    `   ${p.id.padEnd(24)} centre ${String(centre).padStart(3)}   corners ${corners.map((c) => String(c).padStart(3)).join(' ')}`,
  );
}

/* -------------------------------------------------------------------------- */
/* 6b - where an off-centre focus puts the ramp                               */
/* -------------------------------------------------------------------------- */

/** The measured two-stop weight at ramp position `t`, between strip samples. */
function weightAt(t: number): number {
  const u = Math.min(1, Math.max(0, t)) * (STRIP_N - 1);
  const i = Math.min(STRIP_N - 2, Math.floor(u));
  return WEIGHT[i]! + (WEIGHT[i + 1]! - WEIGHT[i]!) * (u - i);
}

const SRGB_STOP = /<a:gs pos="(\d+)"><a:srgbClr val="([0-9A-F]{6})"\/><\/a:gs>/g;

/** The two end colours of a plain two-stop `srgbClr` gradient, by `@pos`. */
function endStops(fill: string): [[number, number, number], [number, number, number]] | null {
  const found: { pos: number; rgb: [number, number, number] }[] = [];
  for (const m of fill.matchAll(SRGB_STOP)) {
    const v = m[2]!;
    found.push({
      pos: Number(m[1]),
      rgb: [parseInt(v.slice(0, 2), 16), parseInt(v.slice(2, 4), 16), parseInt(v.slice(4, 6), 16)],
    });
  }
  if (found.length !== 2) return null;
  found.sort((a, b) => a.pos - b.pos);
  if (found[0]!.pos !== 0 || found[1]!.pos !== PERCENT) return null;
  return [found[0]!.rgb, found[1]!.rgb];
}

/** What a two-stop probe paints at ramp position `t`, per channel, 0 to 255. */
function colorAt(
  stops: [[number, number, number], [number, number, number]],
  t: number,
): [number, number, number] {
  const w = weightAt(t);
  return [0, 1, 2].map((i) => {
    const a = toBlend(stops[0][i]! / 255);
    const b = toBlend(stops[1][i]! / 255);
    return 255 * fromBlend(a + (b - a) * w);
  }) as [number, number, number];
}

/** `ST_Percentage`, hundred-thousandths. */
const PERCENT = 100000;

const FILL_TO_RECT = /<a:fillToRect([^/]*)\/>/;
const INSET_OF = {
  l: /\sl="(-?\d+)"/,
  t: /\st="(-?\d+)"/,
  r: /\sr="(-?\d+)"/,
  b: /\sb="(-?\d+)"/,
} as const;

/**
 * The focus a probe's `a:fillToRect` names, in shape fractions.
 *
 * All four insets zero is the top-left corner, not the whole box: PowerPoint
 * paints `path-circle-whole` with its dark end in the corner, and an element
 * carrying no information is indistinguishable from an absent one.
 */
function focusOf(fill: string): [number, number] | null {
  const m = FILL_TO_RECT.exec(fill);
  if (m === null) return null;
  const attrs = m[1] ?? '';
  const read = (k: keyof typeof INSET_OF): number => {
    const hit = INSET_OF[k].exec(attrs);
    return hit === null ? 0 : Number(hit[1]) / PERCENT;
  };
  const l = read('l');
  const t = read('t');
  const r = read('r');
  const b = read('b');
  if (l === 0 && t === 0 && r === 0 && b === 0) return [0, 0];
  return [(l + (1 - r)) / 2, (t + (1 - b)) / 2];
}

const PATH_KIND = /<a:path(?:\s+path="(\w+)")?\s*[/>]/;

/**
 * Ramp position at a point, given a focus, for one candidate reading.
 *
 * Everything is in export pixels measured from the shape's top-left corner. The
 * export is isotropic - 1920 across 12192000 EMU and 1080 across 6858000 is the
 * same scale on both axes - so a circle in shape units is a circle here too, and
 * a model written in shape *fractions* is a different model on any shape that is
 * not square. That is what the wide and tall probes are for.
 */
type PathModel = (fx: number, fy: number, w: number, h: number) => (x: number, y: number) => number;

/** Solve |P - (F + t(C - F))| = tR: the contour through P, in SVG's focal form. */
function focalAt(
  fx: number,
  fy: number,
  cx: number,
  cy: number,
  radius: number,
): (x: number, y: number) => number {
  const ex = cx - fx;
  const ey = cy - fy;
  const a = ex * ex + ey * ey - radius * radius;
  return (x, y) => {
    const dx = x - fx;
    const dy = y - fy;
    const b = -2 * (dx * ex + dy * ey);
    const c = dx * dx + dy * dy;
    if (Math.abs(a) < 1e-9) return b === 0 ? 0 : -c / b;
    const disc = b * b - 4 * a * c;
    if (disc < 0) return 1;
    const s = Math.sqrt(disc);
    const roots = [(-b + s) / (2 * a), (-b - s) / (2 * a)].filter((t) => t >= -1e-9);
    return roots.length === 0 ? 1 : Math.min(...roots);
  };
}

const farthestCorner = (fx: number, fy: number, w: number, h: number): number =>
  Math.max(
    ...(
      [
        [0, 0],
        [w, 0],
        [0, h],
        [w, h],
      ] as const
    ).map(([x, y]) => Math.hypot(x - fx, y - fy)),
  );

const PATH_MODELS: Readonly<Record<string, PathModel>> = {
  // The reading this experiment was run to test: SVG's own focal radial, whose
  // outer circle is the one the centred probes already fixed.
  focal: (fx, fy, w, h) => focalAt(fx, fy, w / 2, h / 2, Math.hypot(w, h) / 2),
  'focal-farcorner': (fx, fy, w, h) => focalAt(fx, fy, w / 2, h / 2, farthestCorner(fx, fy, w, h)),
  // The reading that shipped from the first pass.
  'concentric-farcorner': (fx, fy, w, h) => {
    const radius = farthestCorner(fx, fy, w, h);
    return (x, y) => Math.hypot(x - fx, y - fy) / radius;
  },
  'concentric-halfdiag': (fx, fy, w, h) => {
    const radius = Math.hypot(w, h) / 2;
    return (x, y) => Math.hypot(x - fx, y - fy) / radius;
  },
  // The reading `path="rect"` needs, and the one a renderer reaches for first.
  box: (fx, fy, w, h) => (x, y) => {
    const dx = x - fx;
    const dy = y - fy;
    const rx = dx >= 0 ? w - fx : fx;
    const ry = dy >= 0 ? h - fy : fy;
    const ax = rx <= 0 ? (dx === 0 ? 0 : 2) : Math.abs(dx) / rx;
    const ay = ry <= 0 ? (dy === 0 ? 0 : 2) : Math.abs(dy) / ry;
    return Math.max(ax, ay);
  },
  // Focal, but in a space where the shape is a unit square - an ellipse fitted
  // to the box rather than a circle. Identical on a square, so only the wide and
  // tall probes can tell it apart from `focal`.
  'focal-ellipse': (fx, fy, w, h) => {
    const inner = focalAt((fx / w) * 2, (fy / h) * 2, 1, 1, Math.SQRT2);
    return (x, y) => inner((x / w) * 2, (y / h) * 2);
  },
};

/** A rectangle of the export, in pixels, whose samples are not to be scored. */
interface Exclusion {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/**
 * Root-mean-square byte error of a ramp-position function over a probe's grid.
 *
 * `at` is given a point in pixels from the probe rectangle's top-left, together
 * with the rectangle's size, and returns where along the ramp that point is.
 */
function scoreGrid(
  p: FixtureProbe,
  fill: string,
  at: (x: number, y: number, w: number, h: number) => number,
  exclude?: Exclusion,
): {
  rms: number;
  worst: number;
  n: number;
  worstAt: { u: number; v: number; t: number } | null;
} | null {
  const g = gridOf(p.id);
  const stops = endStops(fill);
  if (g === null || stops === null) return null;
  const { x0, y0, x1, y1 } = pixelRect(p, 1920, 1080);
  const w = x1 - x0;
  const h = y1 - y0;
  const ax = x0 + 1;
  const ay = y0 + 1;
  const bx = x1 - 2;
  const by = y1 - 2;
  let sum = 0;
  let worst = 0;
  let n = 0;
  let worstAt: { u: number; v: number; t: number } | null = null;
  for (let r = 0; r < GRID_N; r++) {
    for (let c = 0; c < GRID_N; c++) {
      const px = ax + Math.round((c / (GRID_N - 1)) * (bx - ax));
      const py = ay + Math.round((r / (GRID_N - 1)) * (by - ay));
      if (
        exclude !== undefined &&
        px >= exclude.x0 &&
        px < exclude.x1 &&
        py >= exclude.y0 &&
        py < exclude.y1
      ) {
        continue;
      }
      const t = at(px - x0 + 0.5, py - y0 + 0.5, w, h);
      const want = colorAt(stops, t);
      const got = g[r * GRID_N + c]!;
      for (let ch = 0; ch < 3; ch++) {
        const e = Math.abs(want[ch]! - got[ch]!);
        sum += e * e;
        // The sample sitting on the focus is excluded: there the two-colour
        // curve is near vertical and half a pixel of lattice error reads as
        // several bytes, which is a fact about the probe rather than the model.
        if (e > worst && t > 0.05) {
          worst = e;
          worstAt = { u: (px - x0) / w, v: (py - y0) / h, t };
        }
      }
      n++;
    }
  }
  return n === 0 ? null : { rms: Math.sqrt(sum / (n * 3)), worst, n, worstAt };
}

/** Root-mean-square byte error of one path model over one probe's grid. */
function scorePath(
  p: FixtureProbe,
  model: PathModel,
  exclude?: Exclusion,
): {
  rms: number;
  worst: number;
  n: number;
  worstAt: { u: number; v: number; t: number } | null;
} | null {
  // An absent `a:fillToRect` is the centre, measured in section 9 - so these
  // probes are scored here too rather than sitting outside the verdict.
  const focus = focusOf(p.fill) ?? [0.5, 0.5];
  return scoreGrid(
    p,
    p.fill,
    (x, y, w, h) => model(focus[0] * w, focus[1] * h, w, h)(x, y),
    exclude,
  );
}

console.log('');
console.log('6b. Where an off-centre focus puts the ramp.');
console.log('    Hypothesis: concentric circles scaled to the farthest corner, which is what');
console.log('    the first pass shipped. The alternative is SVG’s focal radial.');

/** A fit at or under this is the measurement noise floor: the centred probes sit at 0.9. */
const PATH_TOL = 4;

const MODEL_NAMES = Object.keys(PATH_MODELS);
const pathProbes = fixtureProbes.filter(
  (p) =>
    (p.group === 'focus' || p.group === 'path' || p.group === 'pathdef') && p.grid !== undefined,
);
console.log(
  `    ${'probe'.padEnd(25)}${'kind'.padEnd(8)}${MODEL_NAMES.map((n) => n.slice(0, 20).padStart(22)).join('')}`,
);
interface PathRow {
  id: string;
  kind: string;
  scores: Record<string, number>;
}
const pathRows: PathRow[] = [];
for (const p of pathProbes) {
  const kind = PATH_KIND.exec(p.fill)?.[1] ?? '(none)';
  const scores: Record<string, number> = {};
  for (const name of MODEL_NAMES) {
    const s = scorePath(p, PATH_MODELS[name]!);
    if (s !== null) scores[name] = s.rms;
  }
  if (Object.keys(scores).length === 0) continue;
  pathRows.push({ id: p.id, kind, scores });
  console.log(
    `    ${p.id.padEnd(25)}${kind.padEnd(8)}${MODEL_NAMES.map((n) => (scores[n] ?? NaN).toFixed(2).padStart(22)).join('')}`,
  );
}

/**
 * The verdict, and the refusal.
 *
 * `path="shape"` on an ellipse is excluded: it follows the outline, which none of
 * these rectangle models claims to describe, and the first pass already measured
 * it. Everything else has to fall to exactly one model, and every other model has
 * to be refuted by at least one probe - otherwise the probes did not separate
 * them and the fixture would be recording a coincidence.
 */
function verdict(kinds: readonly string[], expected: string, label: string): void {
  const rows = pathRows.filter((r) => kinds.includes(r.kind) && !r.id.includes('ellipse'));
  if (rows.length === 0) throw new Error(`no probes for ${label}`);
  const byWorst = rows
    .map((r) => ({
      id: r.id,
      worst:
        scorePath(
          fixtureProbes.find((p) => p.id === r.id)!,
          PATH_MODELS[expected]!,
        )?.worst ?? Infinity,
      where:
        scorePath(
          fixtureProbes.find((p) => p.id === r.id)!,
          PATH_MODELS[expected]!,
        )?.worstAt ?? null,
    }))
    .sort((a, b) => b.worst - a.worst);
  console.log(
    `    ${label}: worst single byte under "${expected}" is ${byWorst[0]!.worst.toFixed(1)} on ${byWorst[0]!.id}` +
      ` at u=${byWorst[0]!.where?.u.toFixed(2) ?? '?'} v=${byWorst[0]!.where?.v.toFixed(2) ?? '?'} t=${byWorst[0]!.where?.t.toFixed(3) ?? '?'}` +
      ` (then ${byWorst
        .slice(1, 4)
        .map((r) => `${r.id} ${r.worst.toFixed(1)}`)
        .join(', ')})`,
  );
  const failed = rows.filter((r) => (r.scores[expected] ?? Infinity) > PATH_TOL);
  if (failed.length > 0) {
    throw new Error(
      `${label}: "${expected}" does not fit ${failed.map((r) => `${r.id} (${(r.scores[expected] ?? NaN).toFixed(2)})`).join(', ')}`,
    );
  }
  const survivors = MODEL_NAMES.filter(
    (n) => n !== expected && rows.every((r) => (r.scores[n] ?? Infinity) <= PATH_TOL),
  );
  const worstOf = (n: string): number => Math.max(...rows.map((r) => r.scores[n] ?? Infinity));
  console.log(
    `    ${label}: "${expected}" fits all ${String(rows.length)} at rms <= ${worstOf(expected).toFixed(2)};` +
      ` next best is ${MODEL_NAMES.filter((n) => n !== expected)
        .map((n) => `${n} ${worstOf(n).toFixed(1)}`)
        .join(', ')}`,
  );
  if (survivors.length > 0) {
    throw new Error(
      `${label}: ${survivors.join(', ')} fits as well as "${expected}" - the probes do not separate them`,
    );
  }
}

verdict(['circle'], 'focal', 'path="circle"');
verdict(['rect', 'shape', '(none)'], 'box', 'path="rect", "shape" and the default');

/* -------------------------------------------------------------------------- */
/* 7 - tileRect and flip                                                      */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('7. tileRect and flip.');
console.log('   Hypothesis: a tileRect smaller than the shape makes the fill repeat.');
for (const p of fixtureProbes.filter((q) => q.group === 'tile')) {
  const g = gridOf(p.id);
  if (g === null) continue;
  const row = Array.from({ length: GRID_N }, (_, c) => g[Math.floor(GRID_N / 2) * GRID_N + c]![0]!);
  console.log(
    `   ${p.id.padEnd(22)} centre row ${row.map((n) => String(n).padStart(3)).join(' ')}`,
  );
}

/* -------------------------------------------------------------------------- */
/* 8 - the 54 pattern tiles                                                   */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('8. The 54 pattern tiles.');
console.log('   Hypothesis: every one is an 8x8 tile of device pixels, so it does not scale.');
const periods = new Map<string, number>();
for (const p of fixtureProbes.filter((q) => q.group === 'pattern')) {
  const key = p.tile?.period ? `${String(p.tile.period[0])}x${String(p.tile.period[1])}` : 'none';
  periods.set(key, (periods.get(key) ?? 0) + 1);
}
console.log(
  `   periods at 1920px: ${[...periods].map(([k, n]) => `${k} x${String(n)}`).join(', ')}`,
);
const scaling = fixtureProbes.filter((q) => q.group === 'pattern' && q.tileByWidth);
if (scaling.length > 0) {
  const sample = scaling[0]!;
  console.log(
    `   ${sample.id}: ${Object.entries(sample.tileByWidth!)
      .map(([w, t]) => `${w}px -> ${t ? t.join('x') : 'none'}`)
      .join(', ')}`,
  );
}
const impure = fixtureProbes.filter((q) => q.group === 'pattern' && (q.tile?.soft ?? 0) > 0);
console.log(`   tiles containing an antialiased pixel: ${String(impure.length)}`);

/* -------------------------------------------------------------------------- */
/* 9 - what a:path defaults to                                                */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('9. The defaults inside a:path.');
console.log('   Both @path and a:fillToRect are optional, and a renderer has to pick something.');

/** Score one probe against a named model at a focus the markup did not give. */
function scoreAssuming(p: FixtureProbe, model: string, focus: [number, number]): number | null {
  const inset = (v: number): string => String(Math.round(v * PERCENT));
  const withFocus = p.fill.replace(
    /<a:path(\s+path="\w+")?\s*\/>/,
    `<a:path$1><a:fillToRect l="${inset(focus[0])}" t="${inset(focus[1])}" r="${inset(1 - focus[0])}" b="${inset(1 - focus[1])}"/></a:path>`,
  );
  if (withFocus === p.fill) return null;
  return scorePath({ ...p, fill: withFocus }, PATH_MODELS[model]!)?.rms ?? null;
}

// Not exactly zero: four zero insets are the one case PowerPoint reads as
// carrying no information, so `focusOf` would hand back the corner either way
// and the two candidates would stop being distinguishable.
const DEFAULT_FOCI: readonly (readonly [string, [number, number]])[] = [
  ['centre', [0.5, 0.5]],
  ['top-left', [0.0001, 0.0001]],
];
for (const p of fixtureProbes.filter((q) => q.group === 'pathdef')) {
  const kind = PATH_KIND.exec(p.fill)?.[1] ?? '(none)';
  if (focusOf(p.fill) !== null) {
    const box = scorePath(p, PATH_MODELS['box']!)?.rms ?? NaN;
    const focal = scorePath(p, PATH_MODELS['focal']!)?.rms ?? NaN;
    console.log(
      `   ${p.id.padEnd(22)} @path=${kind.padEnd(8)} box ${box.toFixed(2).padStart(6)}   focal ${focal.toFixed(2).padStart(6)}`,
    );
    continue;
  }
  for (const model of ['box', 'focal'] as const) {
    const cells = DEFAULT_FOCI.map(
      ([label, focus]) =>
        `${label} ${(scoreAssuming(p, model, focus) ?? NaN).toFixed(2).padStart(6)}`,
    );
    console.log(
      `   ${p.id.padEnd(22)} @path=${kind.padEnd(8)} as ${model.padEnd(5)}  ${cells.join('   ')}`,
    );
  }
}

/* -------------------------------------------------------------------------- */
/* 10 - @flip, on a tile with no symmetry to hide behind                      */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('10. @flip on an asymmetric tile.');
console.log('    The first pass used a centred path tile, which is its own mirror on both axes.');
for (const kind of ['path', 'lin'] as const) {
  const rows = fixtureProbes.filter((p) => p.id.startsWith(`flipasym-${kind}-`));
  const baseGrid = gridOf(`flipasym-${kind}-none`);
  if (baseGrid === null) continue;
  for (const p of rows) {
    const g = gridOf(p.id);
    if (g === null) continue;
    let worst = 0;
    for (let i = 0; i < g.length; i++) {
      for (let ch = 0; ch < 3; ch++) {
        worst = Math.max(worst, Math.abs(g[i]![ch]! - baseGrid[i]![ch]!));
      }
    }
    console.log(`    ${p.id.padEnd(22)} worst byte against flip="none": ${String(worst)}`);
  }
}

/* -------------------------------------------------------------------------- */
/* 11 - the rounded corner: device pixels, or a share of the ramp?            */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('11. How wide the softened corner is.');
console.log('    A corner sampled at pixel centres is already blunt: the apex almost never');
console.log('    lands on one. So the candidates are box filters of a given width in device');
console.log('    pixels, and zero - point sampling - is one of them.');

const ARM_PCT = /soften-arm(\d+)/;

/**
 * The profile these probes ask for, before any rasteriser touches it: two
 * straight arms in sRGB meeting at the middle stop, flat beyond the end stops.
 */
function idealProfile(armPx: number, mid: number): (x: number) => number {
  const black = 0x00;
  const apex = 0xe0;
  const tail = 0x20;
  return (x) => {
    const d = x - mid;
    if (d <= -armPx) return black;
    if (d >= armPx) return tail;
    return d <= 0
      ? black + ((apex - black) * (d + armPx)) / armPx
      : apex + ((tail - apex) * d) / armPx;
  };
}

/** The ideal profile averaged over a box of `boxPx` centred on each pixel. */
function boxFiltered(ideal: (x: number) => number, boxPx: number, at: number): number {
  if (boxPx <= 0) return ideal(at);
  const steps = 64;
  let sum = 0;
  for (let k = 0; k < steps; k++) {
    sum += ideal(at + boxPx * ((k + 0.5) / steps - 0.5));
  }
  return sum / steps;
}

/** Widths in device pixels to score. 1.0 is the pixel's own footprint. */
const BOX_WIDTHS = [0, 0.5, 1, 1.6, 2.5, 4];
const softRows: { id: string; export: number; slope: number; rms: Record<string, number> }[] = [];
for (const p of fixtureProbes.filter((q) => q.group === 'soften')) {
  if (p.corner === undefined) continue;
  const armPct = Number(ARM_PCT.exec(p.id)?.[1] ?? 0) / 100;
  for (const [w, packed] of Object.entries(p.corner).sort((a, b) => Number(a[0]) - Number(b[0]))) {
    const value = samples(packed).map((px) => px[0]!);
    const mid = (value.length - 1) / 2;
    const shapePx = (p.cx / inputs.slideSize.cx) * Number(w);
    const armPx = armPct * shapePx;
    // The middle stop is at half the shape, which is a pixel boundary, and the
    // sample at `mid` is the pixel just past it - so the apex is half a pixel
    // to the left of that sample's centre.
    const ideal = idealProfile(armPx, mid - 0.5);
    const span = Math.min(mid, Math.round(armPx));
    const rms: Record<string, number> = {};
    for (const box of BOX_WIDTHS) {
      let sum = 0;
      let n = 0;
      for (let i = mid - span; i <= mid + span; i++) {
        const e = boxFiltered(ideal, box, i) - value[i]!;
        sum += e * e;
        n++;
      }
      rms[String(box)] = Math.sqrt(sum / n);
    }
    softRows.push({
      id: p.id,
      export: Number(w),
      slope: (0xe0 / armPx) * 1,
      rms,
    });
  }
}
console.log(
  `    ${'probe'.padEnd(15)}${'export'.padEnd(8)}${'bytes/px'.padEnd(10)}${BOX_WIDTHS.map((b) => `box ${b.toFixed(1)}`.padStart(10)).join('')}`,
);
for (const row of softRows) {
  console.log(
    `    ${row.id.padEnd(15)}${String(row.export).padEnd(8)}${row.slope.toFixed(1).padEnd(10)}${BOX_WIDTHS.map((b) => (row.rms[String(b)] ?? NaN).toFixed(2).padStart(10)).join('')}`,
  );
}
const bestBox = BOX_WIDTHS.map((b) => ({
  b,
  worst: Math.max(...softRows.map((r) => r.rms[String(b)] ?? Infinity)),
})).sort((a, b) => a.worst - b.worst);
console.log(
  `    best over all ${String(softRows.length)}: box ${bestBox[0]!.b.toFixed(1)} px at worst rms ${bestBox[0]!.worst.toFixed(2)};` +
    ` next ${bestBox[1]!.b.toFixed(1)} px at ${bestBox[1]!.worst.toFixed(2)}`,
);

/*
 * The pairing that settles it without any model at all.
 *
 * `arm1` exported at 3840 and `arm2` exported at 1920 are the same profile in
 * device pixels - 19.2 px per arm, 11.7 bytes per pixel - and different profiles
 * as fractions of the ramp, 1% against 2%. A rounding fixed in device pixels
 * makes them the same picture; a rounding that is a share of the ramp makes the
 * first one twice as blunt.
 */
console.log('');
console.log('    The same corner in device pixels, from two different ramps:');
const cornerOf = (id: string, width: number): number[] | null => {
  const packed = fixtureProbes.find((q) => q.id === id)?.corner?.[String(width)];
  return packed === undefined ? null : samples(packed).map((px) => px[0]!);
};
for (const [fine, coarse, armPx] of [
  ['soften-arm1', 'soften-arm2', 19.2],
  ['soften-arm2', 'soften-arm4', 38.4],
  ['soften-arm4', 'soften-arm8', 76.8],
] as const) {
  const a = cornerOf(fine, 3840);
  const b = cornerOf(coarse, 1920);
  if (a === null || b === null) continue;
  const mid = (a.length - 1) / 2;
  // Three quarters of an arm: the corner and its approaches, and not the outer
  // stops, where the two exports differ for a reason of their own - see below.
  const near = Math.min(mid, Math.round(armPx * 0.75));
  const far = Math.min(mid, Math.round(armPx * 1.2));
  const band = (from: number, to: number): string => {
    let worst = 0;
    let sum = 0;
    let n = 0;
    for (let i = mid - to; i <= mid + to; i++) {
      if (Math.abs(i - mid) < from) continue;
      const e = Math.abs(a[i]! - b[i]!);
      worst = Math.max(worst, e);
      sum += e * e;
      n++;
    }
    return `rms ${Math.sqrt(sum / n).toFixed(2)}, worst ${String(worst)}`;
  };
  console.log(
    `    ${fine}@3840 vs ${coarse}@1920: corner ${band(0, near)};  outer stops ${band(near, far)}`,
  );
}

/* -------------------------------------------------------------------------- */
/* 12 - a gradient on the slide background                                    */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('12. What a gradient in p:bg is laid out over.');
console.log('    Hypothesis: the slide, which is the only rectangle a background has.');
const LIN_ANG = /<a:lin ang="(\d+)"/;
for (const p of fixtureProbes.filter((q) => q.group === 'background')) {
  if (p.grid === undefined) continue;
  // The `-slide` probe is a noFill rectangle over the whole slide, so what it
  // shows is the background - scored against the same markup its `-shape`
  // sibling carries, with the sibling's own rectangle cut out of the samples.
  const sibling = fixtureProbes.find((q) => q.id === `${p.deck}-shape`);
  const onSlide = p.id.endsWith('-slide');
  const fill = sibling?.fill ?? p.fill;
  const exclude = onSlide && sibling !== undefined ? pixelRect(sibling, 1920, 1080) : undefined;
  const kind = PATH_KIND.exec(fill)?.[1];
  const ang = LIN_ANG.exec(fill)?.[1];
  let label: string;
  let score: {
    rms: number;
    worst: number;
    n: number;
    worstAt: { u: number; v: number; t: number } | null;
  } | null;
  if (kind !== undefined) {
    label = kind === 'circle' ? 'focal' : 'box';
    const focus = focusOf(fill) ?? [0.5, 0.5];
    const model = PATH_MODELS[label]!;
    score = scoreGrid(
      p,
      fill,
      (x, y, w, h) => model(focus[0] * w, focus[1] * h, w, h)(x, y),
      exclude,
    );
  } else {
    // Zero points along +x and the angle increases clockwise, measured in
    // section 5. The ramp runs from the box corner that projects lowest onto
    // that direction to the one that projects highest.
    const degrees = Number(ang ?? 0) / 60000;
    const radians = degrees * (Math.PI / 180);
    label = `lin ${degrees.toFixed(0)}deg`;
    const ux = Math.cos(radians);
    const uy = Math.sin(radians);
    score = scoreGrid(
      p,
      fill,
      (x, y, w, h) => {
        const at = (px: number, py: number): number => px * ux + py * uy;
        const ends = [at(0, 0), at(w, 0), at(0, h), at(w, h)];
        const lo = Math.min(...ends);
        return (at(x, y) - lo) / (Math.max(...ends) - lo);
      },
      exclude,
    );
  }
  console.log(
    `    ${p.id.padEnd(22)} ${label.padEnd(10)} rms ${(score?.rms ?? NaN).toFixed(2).padStart(6)} over ${String(score?.n ?? 0)} samples`,
  );
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

if (fixturePath !== null && fixturePath !== undefined) {
  const fixture = {
    experiment: 'C3 - gradients and pattern fills',
    generator: 'tools/ground-truth/paint/fills/analyse.ts',
    powerPoint: 'Microsoft PowerPoint 365, build 16.0.20326, Windows 11',
    slideSize: inputs.slideSize,
    exportWidth: 1920,
    exportHeight: 1080,
    stripSamples: STRIP_N,
    gridSamples: GRID_N,
    blendGamma: BLEND_GAMMA,
    clrScheme: inputs.clrScheme,
    probes: fixtureProbes,
  };
  writeFileSync(fixturePath, `${JSON.stringify(fixture, null, 2)}\n`);
  console.log('');
  console.log(
    `wrote ${fixturePath} (${String(fixtureProbes.length)} probes, ${String(Math.round(JSON.stringify(fixture).length / 1024))} KiB)`,
  );
}

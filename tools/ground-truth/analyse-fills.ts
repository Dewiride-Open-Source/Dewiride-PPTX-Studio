/**
 * Experiment C3, step 3 - read the bitmaps and say what PowerPoint did.
 *
 * ```
 * node tools/ground-truth/analyse-fills.ts <work-dir> [--fixture corpus/ground-truth/fills.json]
 * ```
 *
 * Each section states its hypothesis before it prints its evidence, so a reader
 * can disagree with the question rather than only with the answer. The fixture
 * it writes is the committed measurement; everything printed here is derived
 * from it and nothing in the package depends on this file having been run.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { hex, readBmp, type Bitmap } from './bmp.ts';

/* -------------------------------------------------------------------------- */
/* inputs                                                                     */
/* -------------------------------------------------------------------------- */

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse-fills.ts <work-dir> [--fixture <path>]');
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
  sample: 'strip' | 'grid' | 'tile';
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
    else if (probe.sample === 'grid') entry.grid = sampleGrid(bmp, probe);
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
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

if (fixturePath !== null && fixturePath !== undefined) {
  const fixture = {
    experiment: 'C3 - gradients and pattern fills',
    generator: 'tools/ground-truth/analyse-fills.ts',
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

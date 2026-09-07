import { describe, expect, it } from 'vitest';
import fills from '../../../../corpus/ground-truth/fills.json' with { type: 'json' };
import { PaintError } from '../errors.js';
import { resolveColor } from '../colors/resolve.js';
import {
  gradientColorAt,
  linearGradientVector,
  pathGeometry,
  pathPositionAt,
  resolveGradientStops,
  sortStops,
  svgStops,
  twoColorRamp,
} from './gradient.js';
import { RAMP_BLEND_GAMMA, TWO_STOP_RAMP } from './gradient-ramp.js';
import { ANTIALIASED_PATTERNS, PATTERN_TILES, PRESET_PATTERN_NAMES } from './pattern-tiles.js';
import { patternInk, resolvePattern } from './pattern.js';
import { toByte } from '../colors/transfer.js';
import type { GradientFill, GradientStop, PatternFill, RelativeRect } from './fill.js';
import type { ClrScheme, Color, ColorTransform, Rgba, SchemeSlot } from '../types.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface Probe {
  id: string;
  deck: string;
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
  opened: boolean;
  repaired: boolean | null;
  strip?: string;
  grid?: string;
  tile?: { period: [number, number] | null; rows: string[]; soft: number; fg: string; bg: string };
  tileByWidth?: Record<string, [number, number] | null>;
  com?: { stops: { position: number; rgb: string; transparency: number }[] };
}

const PROBES = fills.probes as unknown as Probe[];
const STRIP_N = fills.stripSamples;
const GRID_N = fills.gridSamples;
const EXPORT_W = fills.exportWidth;
const EXPORT_H = fills.exportHeight;

function probe(id: string): Probe {
  const found = PROBES.find((p) => p.id === id);
  if (found === undefined) throw new Error(`no probe ${id}`);
  return found;
}

/** Unpack a run of six-hex-digit samples into `[r, g, b]` bytes. */
function samples(packed: string): [number, number, number][] {
  const out: [number, number, number][] = [];
  for (let i = 0; i + 6 <= packed.length; i += 6) {
    out.push([
      parseInt(packed.slice(i, i + 2), 16),
      parseInt(packed.slice(i + 2, i + 4), 16),
      parseInt(packed.slice(i + 4, i + 6), 16),
    ]);
  }
  return out;
}

/**
 * The theme the probe decks were written against, from `tools/ground-truth/lib/pptx.ts`.
 * `dk1` and `lt1` are `sysClr` with a `lastClr`, as PowerPoint writes them.
 */
const SCHEME: ClrScheme = Object.fromEntries(
  Object.entries(fills.clrScheme).map(([slot, hex]) => [
    slot,
    slot === 'dk1' || slot === 'lt1'
      ? {
          space: 'sys',
          name: slot === 'dk1' ? 'windowText' : 'window',
          lastClr: hex,
          transforms: [],
        }
      : { space: 'srgb', hex, transforms: [] },
  ]),
) as unknown as ClrScheme;

const CTX = { scheme: SCHEME };

/* -------------------------------------------------------------------------- */
/* a small reader, for the probe markup only                                  */
/* -------------------------------------------------------------------------- */

/**
 * Enough of a parser to turn a probe's fill markup back into a `Fill`.
 *
 * Real parsing is 2.9's job and belongs to `@pptx-studio/model`. This exists so
 * the assertions below run against the exact markup PowerPoint was shown, rather
 * than against a hand-built object that might have drifted from it.
 */
function readColor(xml: string): Color {
  const transforms: ColorTransform[] = [];
  for (const m of xml.matchAll(/<a:(tint|shade|alpha|lumMod|satMod)\s+val="(-?\d+)"\s*\/>/g)) {
    transforms.push({ op: m[1] as 'tint', val: Number(m[2]) });
  }
  const srgb = /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(xml);
  if (srgb) return { space: 'srgb', hex: srgb[1]!.toUpperCase(), transforms };
  const scheme = /<a:schemeClr val="([a-zA-Z0-9]+)"/.exec(xml);
  if (scheme) return { space: 'scheme', name: scheme[1] as 'accent1', transforms };
  const prst = /<a:prstClr val="([a-zA-Z]+)"/.exec(xml);
  if (prst) return { space: 'prst', name: prst[1]!, transforms };
  throw new Error(`no colour in ${xml}`);
}

function readRect(attrs: string): RelativeRect {
  const one = (k: string): number => Number(new RegExp(`${k}="(-?\\d+)"`).exec(attrs)?.[1] ?? 0);
  return { l: one('l'), t: one('t'), r: one('r'), b: one('b') };
}

function readGradient(xml: string): GradientFill {
  const stops: GradientStop[] = [];
  for (const m of xml.matchAll(/<a:gs pos="(-?\d+)">([\s\S]*?)<\/a:gs>/g)) {
    stops.push({ pos: Number(m[1]), color: readColor(m[2]!) });
  }
  const linear = /<a:lin ang="(-?\d+)" scaled="(\d)"\/>/.exec(xml);
  // Both halves optional, and both absences mean something: no `@path` is the
  // box ramp and no `a:fillToRect` is the centre.
  const path =
    /<a:path(?: path="(shape|circle|rect)")?(?:\/>|>(?:<a:fillToRect ([^/]*)\/>)?<\/a:path>)/.exec(
      xml,
    );
  const tile = /<a:tileRect ([^/]*)\/>/.exec(xml);
  return {
    type: 'gradient',
    stops,
    shade: linear
      ? { kind: 'linear', ang: Number(linear[1]), scaled: linear[2] === '1' }
      : path
        ? {
            kind: 'path',
            path: (path[1] ?? 'rect') as 'circle',
            fillToRect: path[2] === undefined ? null : readRect(path[2]),
          }
        : null,
    tileRect: tile ? readRect(tile[1]!) : null,
    flip: (/flip="(none|x|y|xy)"/.exec(xml)?.[1] ?? 'none') as 'none',
    rotWithShape: /rotWithShape="0"/.exec(xml) === null,
  };
}

function readPattern(xml: string): PatternFill {
  const fg = /<a:fgClr>([\s\S]*?)<\/a:fgClr>/.exec(xml);
  const bg = /<a:bgClr>([\s\S]*?)<\/a:bgClr>/.exec(xml);
  return {
    type: 'pattern',
    prst: /<a:pattFill prst="([a-zA-Z0-9]+)"/.exec(xml)?.[1] ?? '',
    fg: fg ? readColor(fg[1]!) : null,
    bg: bg ? readColor(bg[1]!) : null,
  };
}

const bytes = (c: Rgba): [number, number, number] => [toByte(c.r), toByte(c.g), toByte(c.b)];

const resolveColorForTest = (color: Color): Rgba => resolveColor(color, CTX);

/**
 * How far along the shape strip sample `k` was taken, 0 to 1.
 *
 * The sampler walks pixels, so sample `k` is at the centre of a particular
 * pixel and not at `k / (n - 1)` of the way. On a gentle ramp the difference is
 * a fraction of a byte; on `ramp-steep`, which crosses from black to white in
 * two per cent of the width, it is eleven.
 */
function stripPosition(p: Probe, k: number): number {
  const sx = EXPORT_W / fills.slideSize.cx;
  const x0 = Math.round(p.x * sx);
  const x1 = Math.round((p.x + p.cx) * sx);
  const px = Math.min(x1 - 1, x0 + Math.round((k / (STRIP_N - 1)) * (x1 - x0 - 1)));
  return (px - x0 + 0.5) / (x1 - x0);
}

/**
 * The shape-space position of grid sample `(row, col)`.
 *
 * Mirrors `sampleGrid` in `tools/ground-truth/paint/fills/analyse.ts` exactly,
 * including its one-pixel inset - a prediction taken at a slightly different
 * point from the measurement is not a prediction of the measurement.
 */
function gridPoint(p: Probe, row: number, col: number): { x: number; y: number } {
  const sx = EXPORT_W / fills.slideSize.cx;
  const sy = EXPORT_H / fills.slideSize.cy;
  const x0 = Math.round(p.x * sx);
  const y0 = Math.round(p.y * sy);
  const x1 = Math.round((p.x + p.cx) * sx);
  const y1 = Math.round((p.y + p.cy) * sy);
  const px = x0 + 1 + Math.round((col / (GRID_N - 1)) * (x1 - 2 - (x0 + 1)));
  const py = y0 + 1 + Math.round((row / (GRID_N - 1)) * (y1 - 2 - (y0 + 1)));
  return { x: ((px - x0 + 0.5) / (x1 - x0)) * p.cx, y: ((py - y0 + 0.5) / (y1 - y0)) * p.cy };
}

/* -------------------------------------------------------------------------- */

describe('the C3 fixture', () => {
  it('holds every probe the experiment ran', () => {
    expect(PROBES).toHaveLength(182);
    expect(new Set(PROBES.map((p) => p.deck)).size).toBe(34);
    expect(PROBES.filter((p) => p.group === 'pattern')).toHaveLength(54);
    // The second pass, which closed ADR 0022's open questions.
    expect(PROBES.filter((p) => p.group === 'focus')).toHaveLength(14);
    expect(PROBES.filter((p) => p.group === 'background')).toHaveLength(8);
  });

  it('records which packages PowerPoint refused', () => {
    const repaired = PROBES.filter((p) => p.repaired === true).map((p) => p.id);
    // Seven of the fifteen hostile probes. Each is alone in its own package, so
    // each of these is attributed rather than inferred.
    expect(repaired.sort()).toEqual([
      'h-angbig',
      'h-angneg',
      'h-nostops',
      'h-onestop',
      'h-patbad',
      'h-poshi',
      'h-posneg',
    ]);
  });

  it('records the ones it accepted, which are the surprising half', () => {
    const ok = PROBES.filter((p) => p.group === 'hostile' && p.repaired === false).map((p) => p.id);
    expect(ok.sort()).toEqual([
      // A stop position spelled "50%", as 2.6 found legal on a:lumMod.
      'h-duppos',
      'h-ftr-inv',
      'h-nogslst',
      'h-noshade',
      'h-patnobg',
      'h-patnofg',
      'h-patnoprst',
      'h-pospct',
    ]);
  });

  it('is unanimous that an empty gsLst is refused while a missing one is not', () => {
    expect(probe('h-nostops').repaired).toBe(true);
    expect(probe('h-nogslst').repaired).toBe(false);
  });
});

describe('the two-stop curve', () => {
  it('is not linear, and is not linear light either', () => {
    const strip = samples(probe('ramp-bw').strip!);
    const mid = strip[(STRIP_N - 1) / 2]![0];
    expect(mid).toBe(186);
    // The three hypotheses this probe was built to separate.
    expect(mid).not.toBe(128); // linear in sRGB
    expect(mid).not.toBe(188); // linear in linear light
  });

  it('reproduces every two-stop ramp in the fixture to within two bytes', () => {
    const wrong: string[] = [];
    for (const p of PROBES) {
      if (p.strip === undefined || p.hasBackdrop) continue;
      const fill = readGradient(p.fill);
      if (fill.shade?.kind !== 'linear' || fill.shade.ang !== 0) continue;
      const actual = samples(p.strip);
      let worst = 0;
      for (let k = 0; k < actual.length; k++) {
        const got = bytes(gradientColorAt(fill, stripPosition(p, k), CTX));
        for (let i = 0; i < 3; i++) worst = Math.max(worst, Math.abs(got[i]! - actual[k]![i]!));
      }
      if (worst > 2) wrong.push(`${p.id}: worst ${String(worst)}`);
    }
    // Two exceptions, and they are one fact about PowerPoint's rasteriser
    // rather than two about the model: see the test below.
    expect(wrong).toEqual(['ramp-steep: worst 11', 'ramp-revorder: worst 4']);
  });

  it('softens every corner in the ramp profile by about a pixel and a half', () => {
    // Wherever the profile has a corner - a stop boundary, or the apex of a V -
    // PowerPoint rounds it off over roughly 1.6 px of the export. The size of
    // the resulting disagreement is the size of the slope change: eleven bytes
    // at `ramp-steep`, which crosses black to white in 2% of the width, and four
    // at `ramp-revorder`, whose two halves meet at 50%. Everywhere away from a
    // corner both are exact.
    //
    // Recorded rather than modelled. It is a sub-pixel property of the
    // rasteriser, it shrinks to nothing on any ramp a real deck contains, and an
    // SVG stop list has no way to express it.
    const corners: Record<string, number[]> = {
      'ramp-steep': [0.49, 0.51],
      'ramp-revorder': [0.5],
    };
    for (const [id, at] of Object.entries(corners)) {
      const p = probe(id);
      const actual = samples(p.strip!);
      const fill = readGradient(p.fill);
      let away = 0;
      for (let k = 0; k < actual.length; k++) {
        const t = stripPosition(p, k);
        if (at.some((c) => Math.abs(t - c) <= 0.004)) continue;
        away = Math.max(away, Math.abs(bytes(gradientColorAt(fill, t, CTX))[0] - actual[k]![0]));
      }
      expect(away, id).toBeLessThanOrEqual(2);
    }
  });

  it('applies the curve to two distinct colours anchored at both ends', () => {
    /** The probe's stops, resolved and sorted, without the curve expansion. */
    const stopsOf = (id: string) =>
      sortStops(readGradient(probe(id).fill).stops).map((s) => ({
        pos: s.pos / 100000,
        color: resolveColorForTest(s.color),
      }));

    // Two stops, both ends: curved.
    expect(twoColorRamp(stopsOf('ramp-bw'))).not.toBeNull();
    // THREE stops but only two colours, both ends: also curved. That is the
    // shape PowerPoint writes for variant 4 of every gradient in its gallery, so
    // "two stops" would have been the wrong rule for a great many real decks.
    expect(twoColorRamp(stopsOf('ramp-revorder'))).not.toBeNull();
    // Three distinct colours, same three positions: linear.
    expect(twoColorRamp(stopsOf('ramp-3stop'))).toBeNull();
    // Two colours, but no stop at either end: linear.
    expect(twoColorRamp(stopsOf('ramp-inset'))).toBeNull();

    // ...and the two that are not curved are linear in sRGB, to the byte.
    const linear = samples(probe('ramp-3stop').strip!);
    expect(linear[(STRIP_N - 1) / 4]![0]).toBe(0x40);
    expect(linear[(STRIP_N - 1) / 2]![0]).toBe(0x80);
  });

  it('holds the colour flat outside the first and last stop', () => {
    const fill = readGradient(probe('ramp-inset').fill);
    expect(bytes(gradientColorAt(fill, 0.05, CTX))).toEqual([0, 0, 0]);
    expect(bytes(gradientColorAt(fill, 0.95, CTX))).toEqual([255, 255, 255]);
  });

  it('pre-samples the curve into fifteen stops, not two and not thirty-three', () => {
    const fill = readGradient(probe('ramp-bw').fill);
    const emitted = resolveGradientStops(fill, CTX);
    expect(emitted).toHaveLength(15);
    expect(emitted[0]!.pos).toBe(0);
    expect(emitted[emitted.length - 1]!.pos).toBe(1);
  });

  it('reproduces the export when a viewer interpolates those stops linearly', () => {
    // What a browser actually paints, checked against what PowerPoint painted.
    const fill = readGradient(probe('ramp-bw').fill);
    const stops = resolveGradientStops(fill, CTX);
    const actual = samples(probe('ramp-bw').strip!);
    let worst = 0;
    for (let k = 0; k < actual.length; k++) {
      const t = k / (actual.length - 1);
      let painted = toByte(stops[stops.length - 1]!.color.r);
      for (let i = 1; i < stops.length; i++) {
        if (t <= stops[i]!.pos) {
          const a = stops[i - 1]!;
          const b = stops[i]!;
          const u = b.pos === a.pos ? 1 : (t - a.pos) / (b.pos - a.pos);
          painted = 255 * (a.color.r + (b.color.r - a.color.r) * u);
          break;
        }
      }
      worst = Math.max(worst, Math.abs(painted - actual[k]![0]));
    }
    expect(worst).toBeLessThan(2);
  });

  it('blends in a 2.2 power, which is neither sRGB nor linear light', () => {
    expect(RAMP_BLEND_GAMMA).toBe(2.2);
    expect(TWO_STOP_RAMP[0]![0]).toBe(0);
    expect(TWO_STOP_RAMP[TWO_STOP_RAMP.length - 1]![0]).toBe(1);
    // Symmetric: the weight a quarter of the way and three quarters sum to one.
    const quarter = TWO_STOP_RAMP.find(([p]) => p > 0.4)![1];
    expect(quarter).toBeGreaterThan(0.3);
    expect(quarter).toBeLessThan(0.4);
  });

  it('interpolates alpha linearly even when it curves the colour', () => {
    // White fading in over red. A quarter of the way is a linear quarter - 0x40
    // - and not the curve's 0.142, which would paint 0x24.
    const strip = samples(probe('ramp-alpha').strip!);
    expect(strip[(STRIP_N - 1) / 4]).toEqual([255, 0x40, 0x40]);
    expect(strip[(STRIP_N - 1) / 2]).toEqual([255, 0x7f, 0x7f]);

    // And the model agrees: the alpha channel is the factor, untouched by the
    // curve the colour channels go through.
    const fill = readGradient(probe('ramp-alpha').fill);
    for (const t of [0.25, 0.5, 0.75]) {
      expect(gradientColorAt(fill, t, CTX).a, String(t)).toBeCloseTo(t, 3);
    }

    // Composited over the red backdrop, that reproduces the measured pixels -
    // to within a byte, because 2.6 measured that PowerPoint composites the
    // *quantised* colour and rounds its ties down, neither of which this
    // package models.
    for (const [k, want] of [
      [(STRIP_N - 1) / 4, 0x40],
      [(STRIP_N - 1) / 2, 0x7f],
      [(3 * (STRIP_N - 1)) / 4, 0xbf],
    ] as const) {
      const c = gradientColorAt(fill, stripPosition(probe('ramp-alpha'), k), CTX);
      expect(Math.abs(toByte(c.g * c.a) - want), String(k)).toBeLessThanOrEqual(1);
    }
  });
});

describe('the stop list', () => {
  it('is sorted by position, not taken in document order', () => {
    const fill = readGradient(probe('ramp-revorder').fill);
    expect(fill.stops.map((s) => s.pos)).toEqual([50000, 0, 100000]);
    expect(sortStops(fill.stops).map((s) => s.pos)).toEqual([0, 50000, 100000]);

    const strip = samples(probe('ramp-revorder').strip!);
    expect(strip[0]).toEqual([255, 255, 255]);
    expect(strip[STRIP_N - 1]).toEqual([255, 255, 255]);
    expect(strip[(STRIP_N - 1) / 2]![0]).toBeLessThan(3);
  });

  it('and PowerPoint hands them back unsorted, so both facts have to be kept', () => {
    const com = probe('ramp-revorder').com;
    expect(com?.stops.map((s) => Math.round(s.position * 100) / 100)).toEqual([0.5, 0, 1]);
  });

  it('refuses to resolve a gradient with no stops', () => {
    const empty: GradientFill = {
      type: 'gradient',
      stops: [],
      shade: null,
      tileRect: null,
      flip: 'none',
      rotWithShape: true,
    };
    expect(() => resolveGradientStops(empty)).toThrowError(PaintError);
    try {
      resolveGradientStops(empty);
    } catch (error) {
      expect((error as PaintError).code).toBe('FILL_NO_STOPS');
    }
  });
});

describe('the linear ramp geometry', () => {
  it('predicts every angle probe from the shape and the two attributes', () => {
    const wrong: string[] = [];
    for (const p of PROBES.filter((q) => q.group === 'angle')) {
      const fill = readGradient(p.fill);
      const shade = fill.shade;
      if (shade?.kind !== 'linear') continue;
      const v = linearGradientVector(shade, p.cx, p.cy);
      const dx = v.x2 - v.x1;
      const dy = v.y2 - v.y1;
      const len2 = dx * dx + dy * dy;
      const actual = samples(p.grid!);
      let worst = 0;
      for (let r = 0; r < GRID_N; r++) {
        for (let c = 0; c < GRID_N; c++) {
          const { x, y } = gridPoint(p, r, c);
          const t = ((x - v.x1) * dx + (y - v.y1) * dy) / len2;
          const got = bytes(gradientColorAt(fill, t, CTX));
          worst = Math.max(worst, Math.abs(got[0] - actual[r * GRID_N + c]![0]));
        }
      }
      if (worst > 4) wrong.push(`${p.id}: worst ${String(worst)}`);
    }
    expect(wrong).toEqual([]);
  });

  it('leaves the angle alone when scaled is false', () => {
    const v = linearGradientVector({ kind: 'linear', ang: 2700000, scaled: false }, 300, 100);
    const deg = (Math.atan2(v.y2 - v.y1, v.x2 - v.x1) * 180) / Math.PI;
    expect(deg).toBeCloseTo(45, 6);
  });

  it('bends it towards the long axis when scaled is true, not away from it', () => {
    // The measured direction on a 3:1 shape is 71.6 degrees. The inverted
    // reading - the one GDI+'s isAngleScalable is usually described with, and
    // the one a careful derivation from the standard produces - gives 18.4.
    const v = linearGradientVector({ kind: 'linear', ang: 2700000, scaled: true }, 300, 100);
    const deg = (Math.atan2(v.y2 - v.y1, v.x2 - v.x1) * 180) / Math.PI;
    expect(deg).toBeCloseTo(71.565, 2);
  });

  it('spans the full projection of the shape, so a corner is exactly an end stop', () => {
    // At 45 degrees on a square, that projection is the diagonal: the ramp runs
    // corner to corner and neither overshoots nor falls short.
    const v = linearGradientVector({ kind: 'linear', ang: 2700000, scaled: false }, 100, 100);
    expect(v.x1).toBeCloseTo(0, 6);
    expect(v.y1).toBeCloseTo(0, 6);
    expect(v.x2).toBeCloseTo(100, 6);
    expect(v.y2).toBeCloseTo(100, 6);
    const dx = v.x2 - v.x1;
    const dy = v.y2 - v.y1;
    const at = (x: number, y: number): number =>
      ((x - v.x1) * dx + (y - v.y1) * dy) / (dx * dx + dy * dy);
    expect(at(0, 0)).toBeCloseTo(0, 6);
    expect(at(100, 100)).toBeCloseTo(1, 6);
    expect(at(100, 0)).toBeCloseTo(0.5, 6);
  });

  it('does not disagree with itself at a cardinal angle', () => {
    for (const ang of [0, 5400000, 10800000, 16200000]) {
      const a = linearGradientVector({ kind: 'linear', ang, scaled: false }, 300, 100);
      const b = linearGradientVector({ kind: 'linear', ang, scaled: true }, 300, 100);
      expect(b.x1).toBeCloseTo(a.x1, 6);
      expect(b.y1).toBeCloseTo(a.y1, 6);
    }
  });
});

describe('path gradients', () => {
  it('puts stop 0 at the focus, so nothing needs reversing for SVG', () => {
    const grid = samples(probe('path-circle-ctr').grid!);
    const centre = grid[Math.floor((GRID_N * GRID_N) / 2)]![0];
    const corner = grid[0]![0];
    expect(centre).toBeLessThan(20); // the first stop is black
    expect(corner).toBeGreaterThan(240); // the last is white
  });

  it('treats a focus rectangle as a point at its centre', () => {
    // A focus covering the middle half paints byte-for-byte what a point focus
    // at the centre paints, over all 169 samples. The rectangle's extent is not
    // a flat region and does not need modelling.
    const a = probe('path-circle-ctr').grid!;
    const b = probe('path-circle-focusrect').grid!;
    expect(b).toBe(a);
  });

  it('reads circle as a circle, not as an ellipse fitted to the box', () => {
    const p = probe('path-circle-wide');
    const geo = pathGeometry(
      { kind: 'path', path: 'circle', fillToRect: { l: 50000, t: 50000, r: 50000, b: 50000 } },
      p.cx,
      p.cy,
    );
    // The top edge's midpoint is much closer to the centre than the left edge's,
    // in shape units. An ellipse would put both at 1.
    const top = pathPositionAt(geo, p.cx / 2, 0);
    const left = pathPositionAt(geo, 0, p.cy / 2);
    expect(top).toBeCloseTo(0.315, 2);
    expect(left).toBeCloseTo(0.949, 2);
  });

  it('reads rect as concentric rectangles reaching the edge', () => {
    const geo = pathGeometry(
      { kind: 'path', path: 'rect', fillToRect: { l: 50000, t: 50000, r: 50000, b: 50000 } },
      100,
      100,
    );
    expect(pathPositionAt(geo, 50, 50)).toBeCloseTo(0, 6);
    expect(pathPositionAt(geo, 100, 50)).toBeCloseTo(1, 6);
    expect(pathPositionAt(geo, 75, 60)).toBeCloseTo(0.5, 6);
  });

  it('reads an all-zero fillToRect as the top-left corner, not as the whole box', () => {
    const geo = pathGeometry(
      { kind: 'path', path: 'rect', fillToRect: { l: 0, t: 0, r: 0, b: 0 } },
      100,
      100,
    );
    expect(geo.fx).toBe(0);
    expect(geo.fy).toBe(0);
    expect(pathPositionAt(geo, 0, 0)).toBe(0);
    expect(pathPositionAt(geo, 100, 100)).toBe(1);
  });

  it('reads an ABSENT fillToRect as the centre, which four zero insets are not', () => {
    // `pathdef-norect-rect` fits the centre at rms 0.45 and the corner at 72.6.
    // The two spellings are a different picture, so the field cannot be defaulted.
    const absent = pathGeometry({ kind: 'path', path: 'rect', fillToRect: null }, 100, 100);
    expect([absent.fx, absent.fy]).toEqual([50, 50]);
    const zeroes = pathGeometry(
      { kind: 'path', path: 'rect', fillToRect: { l: 0, t: 0, r: 0, b: 0 } },
      100,
      100,
    );
    expect([zeroes.fx, zeroes.fy]).not.toEqual([absent.fx, absent.fy]);
  });

  it('puts the outer circle on the shape, not on the focus', () => {
    // The refuted reading: concentric circles about the focus, scaled to the
    // farthest corner, which is what the first pass shipped. It is out by up to
    // 49 bytes on the eight off-centre probes.
    const geo = pathGeometry(
      { kind: 'path', path: 'circle', fillToRect: { l: 10000, t: 10000, r: 90000, b: 90000 } },
      100,
      100,
    );
    expect([geo.fx, geo.fy]).toEqual([10, 10]);
    expect([geo.cx, geo.cy]).toEqual([50, 50]);
    expect(geo.radius).toBeCloseTo(Math.hypot(100, 100) / 2, 9);
    // A point on the far side reaches the last stop at the shape's own circle,
    // not at a circle centred on the focus.
    expect(pathPositionAt(geo, 100, 100)).toBeCloseTo(1, 6);
    expect(pathPositionAt(geo, 10, 10)).toBeCloseTo(0, 6);
  });

  /**
   * The whole path measurement, re-derived.
   *
   * Twenty-seven probes over eight foci, three aspect ratios and every spelling
   * of `a:path`. This is the test that fails if the focal construction is wrong,
   * and the reason the individual cases above can stay short.
   */
  it.each(
    PROBES.filter(
      (p) =>
        (p.group === 'focus' || p.group === 'path' || p.group === 'pathdef') &&
        p.grid !== undefined &&
        // `shape` on an ellipse follows the outline, which needs the geometry
        // package; `@pptx-studio/paint` has only the box.
        !p.id.includes('ellipse'),
    ).map((p) => [p.id, p] as const),
  )('predicts every sample of %s', (_id, p) => {
    const fill = readGradient(p.fill);
    const shade = fill.shade;
    if (shade?.kind !== 'path') throw new Error('expected a path shade');
    const geo = pathGeometry(shade, p.cx, p.cy);
    const grid = samples(p.grid!);
    let worst = 0;
    let sum = 0;
    let n = 0;
    for (let r = 0; r < GRID_N; r++) {
      for (let c = 0; c < GRID_N; c++) {
        const { x, y } = gridPoint(p, r, c);
        const t = pathPositionAt(geo, x, y);
        const got = bytes(gradientColorAt(fill, t, CTX));
        const want = grid[r * GRID_N + c]!;
        for (let ch = 0; ch < 3; ch++) {
          const e = Math.abs(got[ch]! - want[ch]!);
          // Every probe's worst sample sits on the focus, where the two-colour
          // curve climbs eight bytes in a hundredth of the ramp - so a lattice
          // point half a pixel off reads several bytes out. That is a limit of
          // sampling a near-vertical curve, and it is excluded by position
          // rather than by raising the tolerance for everyone.
          if (t > 0.05) worst = Math.max(worst, e);
          sum += e * e;
          n++;
        }
      }
    }
    expect(Math.sqrt(sum / n)).toBeLessThan(2);
    // Against the raw measurement the geometry is worst by 2.0 bytes on a circle
    // and 4.0 on a box. This path adds the shipped fifteen-knot ramp on top,
    // which the ramp tests bound at 1.9.
    expect(worst).toBeLessThan(7);
  });

  it('all but ignores a tileRect on a path gradient', () => {
    // PowerPoint's own from-corner fill carries tileRect l="-100000" t="-100000",
    // which doubles the tile and should therefore show only half the ramp. With
    // it and without it, every one of the 169 samples agrees to within one byte -
    // so it is decorative, and a renderer that implements it faithfully for a
    // path gradient is implementing something PowerPoint does not do.
    const withTile = samples(probe('path-rect-corner-pp').grid!);
    const without = samples(probe('path-rect-corner-notile').grid!);
    let worst = 0;
    for (let i = 0; i < withTile.length; i++) {
      worst = Math.max(worst, Math.abs(withTile[i]![0] - without[i]![0]));
    }
    expect(worst).toBeLessThanOrEqual(1);
  });
});

describe('@flip', () => {
  /**
   * The first pass could not answer this: its four flip probes were a *centred*
   * path tile, which is its own mirror image on both axes, so agreement between
   * them said nothing. These tiles have an off-centre focus and a 30-degree
   * ramp, and they still agree - so the attribute really is inert.
   */
  it.each(['path', 'lin'] as const)('changes nothing on an asymmetric %s tile', (kind) => {
    const base = samples(probe(`flipasym-${kind}-none`).grid!);
    for (const flip of ['x', 'y', 'xy'] as const) {
      const other = samples(probe(`flipasym-${kind}-${flip}`).grid!);
      let worst = 0;
      for (let i = 0; i < base.length; i++) {
        for (let ch = 0; ch < 3; ch++) {
          worst = Math.max(worst, Math.abs(base[i]![ch]! - other[i]![ch]!));
        }
      }
      expect(worst, `${kind} flip="${flip}"`).toBeLessThanOrEqual(1);
    }
  });

  it('is asymmetric enough for that to mean something', () => {
    // The guard on the test above: if the tile were symmetric, agreement would
    // be free. A mirrored copy of it has to disagree with the original.
    const base = samples(probe('flipasym-lin-none').grid!);
    let worst = 0;
    for (let r = 0; r < GRID_N; r++) {
      for (let c = 0; c < GRID_N; c++) {
        const there = base[r * GRID_N + c]![0];
        const mirrored = base[r * GRID_N + (GRID_N - 1 - c)]![0];
        worst = Math.max(worst, Math.abs(there - mirrored));
      }
    }
    expect(worst).toBeGreaterThan(60);
  });
});

describe('a gradient on the slide background', () => {
  /**
   * A background has no shape box, so its ramp needs a rectangle from somewhere.
   * It is the slide: each deck carries the same fill on its background and on a
   * shape in one quadrant, and predicting the background over the slide's own
   * rectangle lands within a byte of what PowerPoint painted.
   */
  it.each(['bg-lin', 'bg-lin-vert', 'bg-path', 'bg-path-off'])(
    'runs %s over the slide, not over some smaller box',
    (deck) => {
      const slide = probe(`${deck}-slide`);
      const shape = probe(`${deck}-shape`);
      const fill = readGradient(shape.fill);
      const shade = fill.shade;
      if (shade === null) throw new Error('expected a shade');
      const geo = shade.kind === 'path' ? pathGeometry(shade, slide.cx, slide.cy) : null;
      const vector =
        shade.kind === 'linear' ? linearGradientVector(shade, slide.cx, slide.cy) : null;
      const grid = samples(slide.grid!);
      let worst = 0;
      for (let r = 0; r < GRID_N; r++) {
        for (let c = 0; c < GRID_N; c++) {
          const { x, y } = gridPoint(slide, r, c);
          // The comparison shape sits on top of the background in one quadrant.
          if (
            x >= shape.x - slide.x &&
            x < shape.x - slide.x + shape.cx &&
            y >= shape.y - slide.y &&
            y < shape.y - slide.y + shape.cy
          ) {
            continue;
          }
          let t: number;
          if (geo !== null) {
            t = pathPositionAt(geo, x, y);
          } else {
            const dx = vector!.x2 - vector!.x1;
            const dy = vector!.y2 - vector!.y1;
            t = ((x - vector!.x1) * dx + (y - vector!.y1) * dy) / (dx * dx + dy * dy);
          }
          const got = bytes(gradientColorAt(fill, Math.min(1, Math.max(0, t)), CTX));
          const want = grid[r * GRID_N + c]!;
          for (let ch = 0; ch < 3; ch++) worst = Math.max(worst, Math.abs(got[ch]! - want[ch]!));
        }
      }
      expect(worst).toBeLessThan(7);
    },
  );
});

describe('rotWithShape', () => {
  it('turns the ramp with the shape when set, and leaves it in slide space when not', () => {
    const withShape = samples(probe('mix-fliph-rws1').grid!);
    const without = samples(probe('mix-fliph-rws0').grid!);
    const row = (g: [number, number, number][], r: number): number[] =>
      Array.from({ length: GRID_N }, (_, c) => g[r * GRID_N + c]![0]);
    // On a horizontally flipped shape the ramp runs the other way with
    // rotWithShape="1" and the original way with "0". The written gradient is
    // black on the left, so with the flip applied to it the left edge is white.
    expect(row(withShape, 6)[0]).toBeGreaterThan(240);
    expect(row(withShape, 6)[GRID_N - 1]).toBeLessThan(60);
    expect(row(without, 6)[0]).toBeLessThan(60);
    expect(row(without, 6)[GRID_N - 1]).toBeGreaterThan(240);
  });
});

describe('the 54 pattern tiles', () => {
  it('are the names PowerPoint itself wrote, in MsoPatternType order', () => {
    expect(PRESET_PATTERN_NAMES).toHaveLength(54);
    expect(PRESET_PATTERN_NAMES[0]).toBe('pct5');
    expect(PRESET_PATTERN_NAMES[53]).toBe('diagCross');
    expect(Object.keys(PATTERN_TILES)).toEqual([...PRESET_PATTERN_NAMES]);
  });

  it('re-derive from the fixture, so the table cannot drift from the measurement', () => {
    for (const p of PROBES.filter((q) => q.group === 'pattern')) {
      const name = p.id.slice('pat-'.length);
      const tile = PATTERN_TILES[name];
      expect(tile, name).toBeDefined();
      expect([tile!.w, tile!.h], name).toEqual(p.tile!.period);
      expect(tile!.cover, name).toBe(p.tile!.rows.join(''));
    }
  });

  it('are a fixed physical size: eight pixels at 96 DPI, whatever the export', () => {
    const p = probe('pat-pct5');
    expect(p.tileByWidth).toEqual({
      '640': [4, 4],
      '1280': [8, 8],
      '1920': [12, 12],
      '2560': [16, 16],
    });
    const resolved = resolvePattern(readPattern(p.fill));
    // Eight pixels at 96 DPI: six points, 76200 EMU. That is the tile a
    // renderer must give an SVG <pattern> in *slide* units, never as a fraction
    // of the shape.
    expect(resolved.width).toBe(8 * 9525);
    expect(resolved.height).toBe(8 * 9525);
    expect(8 * 9525).toBe(76200);
  });

  it('do not scale with the shape', () => {
    const periods = PROBES.filter((q) => q.group === 'patsize').map((q) => q.tile!.period);
    // Six shape sizes from 15% to 100% of the cell, two patterns; every tile the
    // same period, because the tile is not a fraction of anything.
    expect(new Set(periods.map((x) => String(x))).size).toBe(2);
  });

  it('keep their antialiasing through resolvePattern', () => {
    // A renderer that thresholds the three soft tiles to a bitmap draws visibly
    // harsher hatching than PowerPoint does, so the greys have to survive the
    // whole way to the coverage array.
    for (const name of ANTIALIASED_PATTERNS) {
      const resolved = resolvePattern({ type: 'pattern', prst: name, fg: null, bg: null });
      const soft = resolved.coverage.filter((a) => a > 0 && a < 1);
      expect(soft.length, name).toBeGreaterThan(0);
    }
  });

  it('are one-bit except the three PowerPoint antialiases', () => {
    expect([...ANTIALIASED_PATTERNS]).toEqual(['dnDiag', 'upDiag', 'diagCross']);
    for (const name of PRESET_PATTERN_NAMES) {
      if (ANTIALIASED_PATTERNS.includes(name)) continue;
      const cover = PATTERN_TILES[name]!.cover;
      for (let i = 0; i + 2 <= cover.length; i += 2) {
        expect(['00', 'FF'], name).toContain(cover.slice(i, i + 2));
      }
    }
  });

  it('contain exactly one pair that measures identical', () => {
    const full = (name: string): string => {
      const t = PATTERN_TILES[name]!;
      let s = '';
      for (let r = 0; r < 8; r++) {
        for (let c = 0; c < 8; c++) {
          const i = ((r % t.h) * t.w + (c % t.w)) * 2;
          s += t.cover.slice(i, i + 2);
        }
      }
      return s;
    };
    const byBits = new Map<string, string[]>();
    for (const name of PRESET_PATTERN_NAMES) {
      const key = full(name);
      const bucket = byBits.get(key);
      if (bucket) bucket.push(name);
      else byBits.set(key, [name]);
    }
    const pairs = [...byBits.values()].filter((v) => v.length > 1);
    // GDI+ defines LargeGrid as an alias of Cross, and the pixels agree.
    expect(pairs).toEqual([['lgGrid', 'cross']]);
  });

  it('have names that are labels rather than measurements', () => {
    expect(Math.round(patternInk('pct50') * 100)).toBe(50);
    expect(Math.round(patternInk('pct75') * 100)).toBe(88);
    expect(Math.round(patternInk('pct20') * 100)).toBe(13);
  });

  it('default a missing foreground to black and a missing background to white', () => {
    const fg = resolvePattern({ type: 'pattern', prst: 'lgCheck', fg: null, bg: null });
    expect(bytes(fg.fg)).toEqual([0, 0, 0]);
    expect(bytes(fg.bg)).toEqual([255, 255, 255]);
  });

  it('reject a name outside the enumeration, as PowerPoint does', () => {
    expect(() =>
      resolvePattern({ type: 'pattern', prst: 'notAPattern', fg: null, bg: null }),
    ).toThrowError(PaintError);
  });
});

describe('emitting SVG', () => {
  it('keeps opacity out of the colour so a fade does not darken', () => {
    const fill = readGradient(probe('ramp-alpha').fill);
    const stops = svgStops(fill, CTX);
    expect(stops[0]!.color).toBe('#FFFFFF');
    expect(stops[0]!.opacity).toBe(0);
    expect(stops[stops.length - 1]!.color).toBe('#FFFFFF');
    expect(stops[stops.length - 1]!.opacity).toBe(1);
  });
});

describe('the theme colours the probes used', () => {
  it('resolve through the scheme the decks were written with', () => {
    const fill = readGradient(probe('ramp-scheme').fill);
    const strip = samples(probe('ramp-scheme').strip!);
    expect(bytes(gradientColorAt(fill, 0, CTX))).toEqual([...strip[0]!]);
  });

  it('and the slots are the ones tools/ground-truth/lib/pptx.ts writes', () => {
    expect((SCHEME.accent1 as { hex: string }).hex).toBe('4472C4');
    expect((SCHEME['dk1' as SchemeSlot] as { space: string }).space).toBe('sys');
  });
});

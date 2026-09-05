import { describe, expect, it } from 'vitest';
import lines from '../../../../corpus/ground-truth/lines.json' with { type: 'json' };
import { PaintError } from '../errors.js';
import { COMPOUND_RUNS } from './compound-table.js';
import { PRESET_DASHES, PRESET_DASH_NAMES, isPresetDashName } from './dash-table.js';
import { MARKERS, MARKER_SIZE } from './marker-table.js';
import {
  DEFAULT_LINE_CAP,
  DEFAULT_LINE_JOIN,
  DEFAULT_LINE_WIDTH,
  DEFAULT_MITER_LIMIT,
  compoundRails,
  dashArray,
  dashSegments,
  markerGeometry,
  markerOvershoot,
  resolveLine,
  svgStroke,
  type Line,
} from './line.js';
import {
  BLUR_RADIUS_TO_SIGMA,
  DEFAULT_SHADOW_ALIGN,
  EDGE_GROWTH_PER_RADIUS,
  GLOW_RADIUS_SCALE,
  anchorPoint,
  blurSigma,
  effectFilter,
  effectMargin,
  glowGeometry,
  shadowBox,
  shadowMatrix,
  shadowOffset,
  softEdgeGeometry,
  type Box,
  type Effect,
  type OuterShadow,
} from '../effect.js';
import type { Color } from '../types.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface Edge {
  at: number;
  rising: boolean;
}

interface Probe {
  id: string;
  deck: string;
  group: string;
  question: string;
  opened: boolean;
  repaired: boolean | null;
  rect: { x: number; y: number; cx: number; cy: number };
  edges: Edge[];
  total: number;
  cover: string | null;
  columns: { ink: string; span: string; x0: number; scale: number } | null;
  boxes: {
    ink: Box2 | null;
    red: Box2 | null;
    green: Box2 | null;
  } | null;
  export: { pxPerPoint: number; samples: number; start: number; step: number } | null;
  com: { weight: number | null; insetPen: number | null } | null;
}

interface Box2 {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

const PROBES = lines.probes as unknown as Probe[];
const FINDINGS = lines.findings as unknown as Record<string, unknown>;
const DASH_W = lines.strokeWidths.dash;
const AH_W = lines.strokeWidths.arrowhead;
const JOIN_W = lines.strokeWidths.join;
const CMPD_W = lines.strokeWidths.cmpd;

function probe(id: string): Probe {
  const found = PROBES.find((p) => p.id === id);
  if (found === undefined) throw new Error(`no probe ${id}`);
  return found;
}

function finding<T>(key: string): T {
  const value = FINDINGS[key];
  if (value === undefined) throw new Error(`no finding ${key}`);
  return value as T;
}

/** Coverage back out of the packed profile. */
function unpack(hex: string): number[] {
  const out: number[] = [];
  for (let i = 0; i + 1 < hex.length; i += 2) out.push(parseInt(hex.slice(i, i + 2), 16) / 255);
  return out;
}

/* -------------------------------------------------------------------------- */
/* the dash arrays                                                            */
/* -------------------------------------------------------------------------- */

describe('the eleven preset dash arrays', () => {
  /**
   * Re-derived from the crossings, exactly as the generator does it.
   *
   * The point is not to test the generator twice. It is that the table shipped
   * in `dash-table.ts` is a *measurement*, and the only thing that keeps it one
   * is a test that goes back to the pixels rather than to the table.
   */
  function derive(name: string): number[] {
    const edges = probe(`dash-${name}`).edges.slice(1, -1);
    const period =
      finding<Record<string, { sequence: number[] }>>('dashArrays')[name]!.sequence.length;
    let start = 0;
    while (start < edges.length && !edges[start]!.rising) start++;
    const pairs: [number, number][] = [];
    for (let k = 0; k < period / 2; k++) {
      const a = edges[start + k * 2]!;
      const b = edges[start + k * 2 + 1]!;
      const c = edges[start + k * 2 + 2]!;
      pairs.push([
        Math.round(((b.at - a.at) / DASH_W) * 1000) / 1000,
        Math.round(((c.at - b.at) / DASH_W) * 1000) / 1000,
      ]);
    }
    let lead = 0;
    for (let i = 1; i < pairs.length; i++) if (pairs[i]![0] > pairs[lead]![0]) lead = i;
    return [...pairs.slice(lead), ...pairs.slice(0, lead)].flat();
  }

  it('has exactly the eleven names PowerPoint spelled for itself', () => {
    expect([...PRESET_DASH_NAMES].sort()).toEqual(
      [
        'dash',
        'dashDot',
        'dot',
        'lgDash',
        'lgDashDot',
        'lgDashDotDot',
        'solid',
        'sysDash',
        'sysDashDot',
        'sysDashDotDot',
        'sysDot',
      ].sort(),
    );
  });

  it('re-derives every shipped array from the measured crossings', () => {
    for (const name of PRESET_DASH_NAMES) {
      if (name === 'solid') {
        expect(PRESET_DASHES[name]).toEqual([]);
        continue;
      }
      const measured = derive(name);
      const shipped = PRESET_DASHES[name]!.flatMap((s) => [s.d, s.sp]);
      expect(shipped.length, name).toBe(measured.length);
      shipped.forEach((value, i) => {
        expect(Math.abs(value - measured[i]!), `${name}[${String(i)}]`).toBeLessThan(0.08);
      });
    }
  });

  it('is whole numbers of stroke widths, every one of them', () => {
    for (const segments of Object.values(PRESET_DASHES)) {
      for (const { d, sp } of segments) {
        expect(Number.isInteger(d)).toBe(true);
        expect(Number.isInteger(sp)).toBe(true);
      }
    }
  });

  it('paints the same as the a:custDash everyone quotes, for all ten', () => {
    const folklore = finding<Record<string, { same: boolean }>>('folklore');
    const disagreed = Object.entries(folklore)
      .filter(([, v]) => !v.same)
      .map(([k]) => k);
    expect(disagreed).toEqual([]);
    expect(Object.keys(folklore)).toHaveLength(10);
  });

  it('scales with @w rather than being an absolute length', () => {
    const scaling = finding<Record<string, number[]>>('widthScaling');
    for (const [w, [, , onPerW, offPerW]] of Object.entries(scaling)) {
      expect(Math.abs(onPerW! - 4), `w=${w} on`).toBeLessThan(0.05);
      expect(Math.abs(offPerW! - 3), `w=${w} off`).toBeLessThan(0.05);
    }
    expect(Object.keys(scaling).sort()).toEqual(['24', '3', '6']);
  });

  it('rejects a name outside the eleven with a typed error', () => {
    expect(isPresetDashName('dotDotDash')).toBe(false);
    try {
      dashSegments({ kind: 'preset', val: 'dotDotDash' });
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(PaintError);
      expect((error as PaintError).code).toBe('LINE_DASH_UNKNOWN');
      expect((error as PaintError).token).toBe('dotDotDash');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* the cap, and what it does to a dash                                        */
/* -------------------------------------------------------------------------- */

describe('the cap', () => {
  it('defaults to flat, not to the square ECMA-376 specifies', () => {
    const extents = finding<Record<string, { overshoot: number }>>('capExtents');
    expect(extents['absent']!.overshoot).toBe(0);
    expect(extents['flat']!.overshoot).toBe(0);
    expect(DEFAULT_LINE_CAP).toBe('flat');
    // The measurement that makes it a fact and not a preference: a square cap on
    // the same line overshoots by exactly half a stroke width, and the absent
    // case does not.
    expect(extents['sq']!.overshoot / 24).toBeCloseTo(0.5, 3);
    expect(extents['rnd']!.overshoot / 24).toBeCloseTo(0.5, 3);
  });

  it('leaves the painted dash period alone, whichever cap is asked for', () => {
    const capDash = finding<Record<string, { on: number; off: number; period: number }>>('capDash');
    for (const name of ['dot', 'dash']) {
      const flat = capDash[`${name}-flat`]!;
      for (const cap of ['sq', 'rnd']) {
        const other = capDash[`${name}-${cap}`]!;
        expect(Math.abs(other.on - flat.on), `${name} ${cap} on`).toBeLessThan(0.03);
        expect(Math.abs(other.off - flat.off), `${name} ${cap} off`).toBeLessThan(0.03);
      }
    }
  });

  it('compensates a round cap so a dot paints one width and not two', () => {
    const w = 12700;
    const dot = PRESET_DASHES['dot']!;
    expect(dashArray(dot, w, 'flat')).toEqual([w, 3 * w]);
    // Shortened by a width and the gap lengthened by one, so the round cap adds
    // the width back and the period is unchanged.
    expect(dashArray(dot, w, 'rnd')).toEqual([0, 4 * w]);
    expect(dashArray(dot, w, 'sq')).toEqual([0, 4 * w]);
    const period = (a: number[]): number => a.reduce((x, y) => x + y, 0);
    expect(period(dashArray(dot, w, 'rnd'))).toBe(period(dashArray(dot, w, 'flat')));
  });

  it('compensates a longer dash the same way', () => {
    const w = 12700;
    expect(dashArray(PRESET_DASHES['dash']!, w, 'rnd')).toEqual([3 * w, 4 * w]);
    expect(dashArray(PRESET_DASHES['dash']!, w, 'flat')).toEqual([4 * w, 3 * w]);
  });

  it('clamps rather than dropping a segment a cap would make negative', () => {
    const w = 12700;
    // sysDot is [1, 1]: with a cap the dash goes to zero and cannot go below.
    expect(dashArray(PRESET_DASHES['sysDot']!, w, 'rnd')).toEqual([0, 2 * w]);
  });
});

/* -------------------------------------------------------------------------- */
/* widths and alignment                                                       */
/* -------------------------------------------------------------------------- */

describe('width and alignment', () => {
  it('defaults an a:ln with no @w to three quarters of a point', () => {
    expect(DEFAULT_LINE_WIDTH).toBe(9525);
    const widths = finding<Record<string, { comWeight: number | null }>>('widths');
    expect(widths['w-absent']!.comWeight).toBe(0.75);
    expect(widths['w-9525']!.comWeight).toBe(0.75);
    expect(resolveLine(blankLine()).width).toBe(9525);
  });

  it('paints w="0" as a hairline rather than as nothing', () => {
    const widths =
      finding<Record<string, { thickness: number; comWeight: number | null }>>('widths');
    expect(widths['w-zero']!.comWeight).toBe(0);
    // Half a point is one pixel at the export's two pixels to the point.
    expect(widths['w-zero']!.thickness).toBe(0.5);
    expect(probe('w-zero').total).toBeGreaterThan(0);
  });

  it('accepts the 1584pt maximum and covers the slide with it', () => {
    const widths = finding<Record<string, { comWeight: number | null }>>('widths');
    expect(widths['w-max']!.comWeight).toBe(1584);
    // The read column is 300pt tall and every point of it is ink.
    expect(probe('w-max').total).toBeGreaterThan(299);
  });

  it('centres the stroke on the geometry when @algn is absent', () => {
    const alignment = finding<Record<string, { outside: number }>>('alignment');
    expect(alignment['algn-absent']!.outside).toBe(10);
    expect(alignment['algn-ctr']!.outside).toBe(10);
    expect(resolveLine(blankLine()).algn).toBe('ctr');
  });

  it('puts algn="in" wholly inside the shape', () => {
    const alignment = finding<Record<string, { outside: number }>>('alignment');
    // Half a point is one pixel of antialiasing at the export resolution; the
    // centred cases are out by ten points, which is half the stroke.
    expect(alignment['algn-in']!.outside).toBeLessThanOrEqual(0.5);
    expect(probe('algn-in').com?.insetPen).toBe(-1);
    expect(probe('algn-ctr').com?.insetPen).toBe(0);
  });
});

/* -------------------------------------------------------------------------- */
/* the compound strokes                                                       */
/* -------------------------------------------------------------------------- */

describe('the compound strokes', () => {
  it('re-derives every shipped division from the measured crossings', () => {
    const compound = finding<Record<string, { at: number[] }>>('compound');
    for (const [name, runs] of Object.entries(COMPOUND_RUNS)) {
      const at = compound[`cmpd-${name}`]!.at;
      expect(runs.length, name).toBe(at.length - 1);
      runs.forEach((run, i) => {
        expect(Math.abs(run - (at[i + 1]! - at[i]!)), `${name}[${String(i)}]`).toBeLessThan(0.01);
      });
    }
  });

  it('subdivides @w rather than adding to it - every form sums to one', () => {
    for (const [name, runs] of Object.entries(COMPOUND_RUNS)) {
      const total = runs.reduce((a, b) => a + b, 0);
      expect(Math.abs(total - 1), name).toBeLessThan(0.001);
    }
  });

  it('puts the thick rail outside for thickThin and inside for thinThick', () => {
    expect(COMPOUND_RUNS['thickThin']![0]).toBeCloseTo(0.6, 3);
    expect(COMPOUND_RUNS['thinThick']![0]).toBeCloseTo(0.2, 3);
    expect(COMPOUND_RUNS['thickThin']!.at(-1)).toBeCloseTo(0.2, 3);
    expect(COMPOUND_RUNS['thinThick']!.at(-1)).toBeCloseTo(0.6, 3);
  });

  it('lays a centred double stroke across the geometry, not beside it', () => {
    const w = 36000;
    const rails = compoundRails('dbl', w, 'ctr');
    expect(rails).toHaveLength(2);
    expect(rails[0]!.width).toBeCloseTo(w / 3, 6);
    expect(rails[1]!.width).toBeCloseTo(w / 3, 6);
    // Outer rail centred a third of the width outward, inner a third inward.
    expect(rails[0]!.offset).toBeCloseTo(-w / 3, 6);
    expect(rails[1]!.offset).toBeCloseTo(w / 3, 6);
    const span = rails[1]!.offset + rails[1]!.width / 2 - (rails[0]!.offset - rails[0]!.width / 2);
    expect(span).toBeCloseTo(w, 6);
  });

  it('lays an inset double stroke from the edge inwards', () => {
    const w = 36000;
    const rails = compoundRails('dbl', w, 'in');
    expect(rails[0]!.offset).toBeCloseTo(w / 6, 6);
    expect(rails[1]!.offset).toBeCloseTo((5 * w) / 6, 6);
    // Which is exactly what the measurement saw: crossings at 0, 1/3, 2/3 and 1
    // of the width, all inside the shape.
    const compound = finding<Record<string, { at: number[] }>>('compound');
    const at = compound['cmpd-dbl-in']!.at;
    expect(at[0]).toBeCloseTo(0, 1);
    expect(at.at(-1)).toBeCloseTo(1, 1);
  });

  it('gives tri five runs, thickest in the middle', () => {
    const runs = COMPOUND_RUNS['tri']!;
    expect(runs).toHaveLength(5);
    expect(runs[2]).toBeCloseTo(1 / 3, 3);
    expect(compoundRails('tri', 36000, 'ctr')).toHaveLength(3);
  });

  it('rejects a compound name outside the five', () => {
    try {
      compoundRails('quad', 12700, 'ctr');
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(PaintError);
      expect((error as PaintError).code).toBe('LINE_COMPOUND_UNKNOWN');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* joins                                                                      */
/* -------------------------------------------------------------------------- */

describe('joins', () => {
  it('defaults to round, which is not what SVG defaults to', () => {
    const joins = finding<Record<string, { inWidths: number }>>('joins');
    expect(joins['join-absent']!.inWidths).toBe(joins['join-round']!.inWidths);
    expect(joins['join-absent']!.inWidths).toBeCloseTo(0.5, 2);
    expect(joins['join-bevel']!.inWidths).toBeLessThan(0.4);
    expect(DEFAULT_LINE_JOIN).toBe('round');
    expect(resolveLine(blankLine()).join.kind).toBe('round');
  });

  it('has a default miter limit of eight, bracketed to a tenth', () => {
    const limits = finding<Record<string, { ratio: number; clipped: boolean }>>('miterLimits');
    const full = Object.values(limits)
      .filter((v) => !v.clipped)
      .map((v) => v.ratio);
    const clipped = Object.values(limits)
      .filter((v) => v.clipped)
      .map((v) => v.ratio);
    expect(Math.max(...full)).toBeLessThan(8);
    expect(Math.min(...clipped)).toBeGreaterThanOrEqual(8);
    // The bracket is tight: 7.91 is full and 8.0 is not.
    expect(Math.max(...full)).toBeGreaterThan(7.9);
    expect(DEFAULT_MITER_LIMIT).toBe(8);
    expect(resolveLine(blankLine()).join.limit).toBe(8);
  });

  it('measures @lim against the whole stroke width, not half of it', () => {
    const units = finding<Record<string, { limitPct: number; clipped: boolean }>>('miterUnits');
    // The corner's own miter ratio is 5.
    expect(units['mlim-400']!.clipped).toBe(true);
    expect(units['mlim-600']!.clipped).toBe(false);
    // In half-width units the change would have fallen between 800% and 1200%.
    expect(units['mlim-200']!.clipped).toBe(true);
    expect(units['mlim-1000']!.clipped).toBe(false);
  });

  it('writes both the join and the limit into SVG rather than relying on defaults', () => {
    const stroke = svgStroke(resolveLine(blankLine()));
    expect(stroke['stroke-linejoin']).toBe('round');
    expect(stroke['stroke-miterlimit']).toBe(8);
    expect(stroke['stroke-linecap']).toBe('butt');
  });
});

/* -------------------------------------------------------------------------- */
/* arrowheads                                                                 */
/* -------------------------------------------------------------------------- */

describe('arrowheads', () => {
  interface Head {
    outline: string;
    length: number;
    width: number;
    notch: number;
    fillRatio: number;
  }
  const HEADS = finding<Record<string, Head>>('arrowheads');

  it('has three sizes, the same three for @len and @w', () => {
    expect(MARKER_SIZE).toEqual({ sm: 2, med: 3, lg: 5 });
    for (const [size, expected] of Object.entries(MARKER_SIZE)) {
      for (const other of ['sm', 'med', 'lg']) {
        const byLength = HEADS[`triangle-l${size}-w${other}`]!;
        const byWidth = HEADS[`triangle-l${other}-w${size}`]!;
        expect(Math.abs(byLength.length - expected), `len ${size}`).toBeLessThan(0.1);
        expect(Math.abs(byWidth.width - expected), `w ${size}`).toBeLessThan(0.15);
      }
    }
  });

  it('keeps the same size in stroke widths at 4, 8 and 16 points', () => {
    const scale =
      finding<Record<string, { w: number; length: number; width: number }>>('headScale');
    for (const record of Object.values(scale)) {
      expect(Math.abs(record.length - 3), `w=${String(record.w)}`).toBeLessThan(0.1);
      expect(Math.abs(record.width - 3), `w=${String(record.w)}`).toBeLessThan(0.1);
    }
    expect(Object.keys(scale)).toHaveLength(2);
  });

  it('centres a diamond and an oval on the endpoint and tips the rest', () => {
    const anchors = finding<Record<string, string>>('headAnchors');
    expect(anchors['diamond']).toBe('centre');
    expect(anchors['oval']).toBe('centre');
    expect(anchors['triangle']).toBe('tip');
    expect(anchors['stealth']).toBe('tip');
    expect(anchors['arrow']).toBe('tip');
    expect(MARKERS['diamond']!.anchor).toBe('centre');
    expect(MARKERS['oval']!.anchor).toBe('centre');
    expect(MARKERS['triangle']!.anchor).toBe('tip');
  });

  it('puts the endpoint at the right place along the marker', () => {
    const w = 12700;
    const triangle = markerGeometry({ type: 'triangle', w: 'med', len: 'med' }, w)!;
    expect(triangle.length).toBe(3 * w);
    expect(triangle.width).toBe(3 * w);
    // A tip-anchored head sits entirely behind the endpoint.
    expect(triangle.refX).toBe(3 * w);
    expect(markerOvershoot({ type: 'triangle', w: 'med', len: 'med' }, w)).toBe(0);
    // A centred one hangs half its length past it, which is what the ah-both
    // probe measured: an oval head at a line's start reached 40pt before it.
    const oval = markerGeometry({ type: 'oval', w: 'lg', len: 'lg' }, w)!;
    expect(oval.refX).toBe(oval.length / 2);
    expect(markerOvershoot({ type: 'oval', w: 'lg', len: 'lg' }, w)).toBe(oval.length / 2);
  });

  it('agrees with the both-ends probe about where an oval head reaches', () => {
    // A 16pt line from 300 to 700 with a large oval at the head: 5 x 16 = 80pt
    // long, centred, so 40pt before the start.
    const both = probe('ah-both');
    const first = both.edges[0]!.at;
    expect(300 - first).toBeCloseTo(markerOvershoot({ type: 'oval', w: 'lg', len: 'lg' }, 16), 0);
  });

  it('separates a stealth from the triangle it shares a silhouette with', () => {
    const triangle = Object.entries(HEADS).filter(([k]) => k.startsWith('triangle-'));
    const stealth = Object.entries(HEADS).filter(([k]) => k.startsWith('stealth-'));
    const mean = (rows: [string, Head][], pick: (h: Head) => number): number =>
      rows.reduce((a, [, h]) => a + pick(h), 0) / rows.length;
    // Same silhouette...
    expect(mean(triangle, (h) => h.length)).toBeCloseTo(
      mean(stealth, (h) => h.length),
      0,
    );
    // ...and a quarter of it not inked.
    expect(mean(triangle, (h) => h.fillRatio)).toBeGreaterThan(0.97);
    expect(mean(stealth, (h) => h.fillRatio)).toBeLessThan(0.9);
    expect(MARKERS['stealth']!.notch).toBeCloseTo(0.25, 2);
    expect(MARKERS['triangle']!.notch).toBe(0);
    expect(MARKERS['stealth']!.path).not.toBe(MARKERS['triangle']!.path);
  });

  it('treats an open arrow as stroked rather than filled', () => {
    expect(MARKERS['arrow']!.filled).toBe(false);
    expect(MARKERS['triangle']!.filled).toBe(true);
    expect(markerGeometry({ type: 'arrow', w: 'med', len: 'med' }, 12700)!.filled).toBe(false);
    // Its silhouette is wider than its nominal width, because the V is stroked.
    expect(HEADS['arrow-lmed-wmed']!.width).toBeGreaterThan(MARKER_SIZE['med']!);
  });

  it('is nothing at all for type="none"', () => {
    expect(markerGeometry({ type: 'none', w: 'med', len: 'med' }, 12700)).toBeNull();
    expect(markerGeometry(null, 12700)).toBeNull();
    for (const [key, head] of Object.entries(HEADS)) {
      if (key.startsWith('none-')) expect(head.width).toBe(0);
    }
  });

  it('rejects a type or a size outside the enumerations', () => {
    for (const [end, code] of [
      [{ type: 'chevron', w: 'med', len: 'med' }, 'LINE_END_UNKNOWN'],
      [{ type: 'triangle', w: 'xl', len: 'med' }, 'LINE_END_SIZE'],
    ] as const) {
      try {
        markerGeometry(end as never, 12700);
        expect.unreachable();
      } catch (error) {
        expect(error).toBeInstanceOf(PaintError);
        expect((error as PaintError).code).toBe(code);
      }
    }
  });
});

/* -------------------------------------------------------------------------- */
/* the blur, and the three things that use it                                 */
/* -------------------------------------------------------------------------- */

describe('the blur', () => {
  interface Rule {
    element: string;
    declared: number;
    r: number;
    sigmaOverR: number;
    offsetOverR: number;
  }
  const RULES = finding<{ rules: Rule[] }>('blurRule').rules;

  it('is one third of the radius, on all four elements', () => {
    expect(RULES.length).toBeGreaterThanOrEqual(14);
    for (const rule of RULES) {
      expect(
        Math.abs(rule.sigmaOverR - BLUR_RADIUS_TO_SIGMA),
        `${rule.element} ${String(rule.declared)}`,
      ).toBeLessThan(0.005);
    }
    expect(BLUR_RADIUS_TO_SIGMA).toBeCloseTo(1 / 3, 10);
    expect(blurSigma(30000)).toBe(10000);
  });

  it('covers all four elements and not just one', () => {
    expect(new Set(RULES.map((r) => r.element))).toEqual(
      new Set(['outerShdw', 'blur', 'glow', 'softEdge']),
    );
  });

  it('is a Gaussian, not merely something that spreads', () => {
    const fits = finding<Record<string, { worst: number; blurRad: number }>>('blurFits');
    for (const [id, fit] of Object.entries(fits)) {
      if (fit.blurRad === 0) continue;
      // An error function fitted to the measured edge, in units of full scale.
      expect(fit.worst * 255, id).toBeLessThan(1.5);
    }
  });

  it('is isotropic, to the byte', () => {
    const isotropy = finding<Record<string, { samples: number; worst: number }>>('blurIsotropy');
    for (const [rad, record] of Object.entries(isotropy)) {
      expect(record.samples, rad).toBeGreaterThan(20);
      expect(record.worst, rad).toBe(0);
    }
  });

  it('leaves the edge where it was for a shadow and a plain blur', () => {
    for (const rule of RULES.filter((r) => r.element === 'outerShdw' || r.element === 'blur')) {
      expect(Math.abs(rule.offsetOverR), rule.element).toBeLessThan(0.05);
    }
  });

  it('moves the edge nine tenths of a radius for a glow and a soft edge', () => {
    // The tolerance is the measurement's own resolution and not a round number.
    // The edge position is recovered from a 1920-wide export, so it is good to
    // about a quarter of a point; as a fraction of r that is 0.25/r, which is a
    // tenth at r=2 and a hundredth at r=32. A flat tolerance would either pass
    // the small radii on nothing or fail them on noise.
    for (const rule of RULES.filter((r) => r.element === 'glow' || r.element === 'softEdge')) {
      const resolution = 0.25 / rule.r;
      expect(
        Math.abs(Math.abs(rule.offsetOverR) - EDGE_GROWTH_PER_RADIUS),
        `${rule.element} ${String(rule.declared)}pt, r=${String(rule.r)}`,
      ).toBeLessThan(0.02 + resolution);
    }
    // And the largest radius, where the resolution is finest, pins it hardest.
    const finest = RULES.filter((r) => r.element === 'softEdge').sort((a, b) => b.r - a.r)[0]!;
    expect(finest.r).toBe(32);
    expect(Math.abs(Math.abs(finest.offsetOverR) - EDGE_GROWTH_PER_RADIUS)).toBeLessThan(0.01);
    // Outward for a glow, inward for a soft edge.
    expect(RULES.filter((r) => r.element === 'glow').every((r) => r.offsetOverR > 0)).toBe(true);
    expect(RULES.filter((r) => r.element === 'softEdge').every((r) => r.offsetOverR < 0)).toBe(
      true,
    );
  });

  it('halves a glow radius before applying either constant', () => {
    expect(GLOW_RADIUS_SCALE).toBe(0.5);
    for (const rule of RULES.filter((r) => r.element === 'glow')) {
      expect(rule.r).toBe(rule.declared / 2);
    }
    const glow = glowGeometry(16000);
    expect(glow.sigma).toBeCloseTo(8000 / 3, 6);
    expect(glow.grow).toBeCloseTo(8000 * 0.9, 6);
    const soft = softEdgeGeometry(16000);
    expect(soft.sigma).toBeCloseTo(16000 / 3, 6);
    expect(soft.shrink).toBeCloseTo(16000 * 0.9, 6);
  });

  it('refuses a negative radius with a typed error', () => {
    try {
      blurSigma(-1);
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(PaintError);
      expect((error as PaintError).code).toBe('EFFECT_NEGATIVE_RADIUS');
    }
  });
});

/* -------------------------------------------------------------------------- */
/* the shadow as a transform                                                  */
/* -------------------------------------------------------------------------- */

describe('the shadow', () => {
  const BOX: Box = { x: 0, y: 0, w: 100, h: 100 };

  it('points dir=0 along +x and turns it toward +y', () => {
    const dirs = finding<Record<string, { dx: number | null; dy: number | null }>>('shadowDirs');
    for (const [deg, measured] of Object.entries(dirs)) {
      const { dx, dy } = shadowOffset(30, Number(deg) * 60000);
      expect(Math.abs(dx - measured.dx!), `dir=${deg} dx`).toBeLessThan(0.6);
      expect(Math.abs(dy - measured.dy!), `dir=${deg} dy`).toBeLessThan(0.6);
    }
    expect(Object.keys(dirs)).toHaveLength(8);
  });

  it('anchors an untransformed shadow nowhere in particular', () => {
    const shadow = plainShadow({ dist: 200, dir: 0 });
    expect(shadowBox(shadow, BOX)).toEqual({ x: 200, y: 0, w: 100, h: 100 });
  });

  it('defaults @algn to bottom centre, which is what the schema says too', () => {
    expect(DEFAULT_SHADOW_ALIGN).toBe('b');
    expect(anchorPoint('b', BOX)).toEqual({ x: 50, y: 100 });
    // sx doubles about the horizontal centre; sy doubles upward from the bottom.
    const doubled = shadowBox(plainShadow({ sx: 200000, sy: 200000 }), BOX);
    expect(doubled).toEqual({ x: -50, y: -100, w: 200, h: 200 });
  });

  it('reproduces every affine the experiment measured', () => {
    const affine =
      finding<Record<string, { dx0: number; dy0: number; dw: number; dh: number }>>('shadowAffine');
    // The probes were a 100 x 55pt block with the shadow thrown 200pt right.
    const shapeBox: Box = { x: 0, y: 0, w: 100, h: 55 };
    const cases: Record<string, Partial<Parameters<typeof plainShadow>[0]>> = {
      'shdw-plain': {},
      'shdw-sx': { sx: 200000 },
      'shdw-sy': { sy: 200000 },
      'shdw-syneg': { sy: -100000 },
      'shdw-kx': { kx: 20 * 60000 },
      'shdw-ky': { ky: 20 * 60000 },
      'shdw-algn-bl': { sx: 200000, sy: 200000, algn: 'bl' },
      'shdw-algn-ctr': { sx: 200000, sy: 200000, algn: 'ctr' },
      'shdw-algn-tl': { sx: 200000, sy: 200000, algn: 'tl' },
      'shdw-algn-absent': { sx: 200000, sy: 200000 },
    };
    const control = shadowBox(plainShadow({ dist: 200, dir: 0 }), shapeBox);
    for (const [id, options] of Object.entries(cases)) {
      const measured = affine[id];
      if (measured === undefined) continue;
      const box = shadowBox(plainShadow({ dist: 200, dir: 0, ...options }), shapeBox);
      expect(Math.abs(box.x - control.x - measured.dx0), `${id} dx0`).toBeLessThan(1);
      expect(Math.abs(box.y - control.y - measured.dy0), `${id} dy0`).toBeLessThan(1);
      expect(Math.abs(box.w - control.w - measured.dw), `${id} dw`).toBeLessThan(1);
      expect(Math.abs(box.h - control.h - measured.dh), `${id} dh`).toBeLessThan(1);
    }
  });

  it('builds a matrix that is the identity when nothing is asked of it', () => {
    // The one case where the answer is knowable without measuring anything, and
    // therefore the one that catches an anchor term left in by mistake: with no
    // scale, no skew and no offset, every anchor must give the identity.
    for (const algn of ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'] as const) {
      const matrix = shadowMatrix(plainShadow({ algn }), BOX);
      expect(matrix.map((v) => Math.round(v * 1e9) / 1e9)).toEqual([1, 0, 0, 1, 0, 0]);
    }
  });

  it('widens a skewed shadow by the height times the tangent', () => {
    // kx=20 degrees on a 100 x 55pt block widened the measured box by 20pt, and
    // 55 * tan(20 degrees) is 20.02.
    const box: Box = { x: 0, y: 0, w: 100, h: 55 };
    const skewed = shadowBox(plainShadow({ kx: 20 * 60000 }), box);
    expect(skewed.w - box.w).toBeCloseTo(box.h * Math.tan((20 * Math.PI) / 180), 6);
    expect(skewed.h).toBeCloseTo(box.h, 6);
  });

  it('mirrors about the anchor for sy=-100%', () => {
    const box = shadowBox(plainShadow({ sy: -100000 }), { x: 0, y: 0, w: 100, h: 55 });
    // Bottom-anchored, so the mirror leaves the bottom edge and moves the box
    // down by its own height.
    expect(box.y).toBeCloseTo(55, 6);
    expect(box.h).toBeCloseTo(55, 6);
  });

  it('turns with the shape only when rotWithShape says so', () => {
    const rot = finding<Record<string, { centre: [number, number] | null }>>('rotWithShape');
    const still = rot['shdw-rotwith0']!.centre!;
    const turned = rot['shdw-rotwith1']!.centre!;
    // The shape is at 45 degrees and the shadow is thrown 80pt at dir=0.
    // Without rotWithShape it stays horizontal; with it, it follows the shape.
    const shape0: [number, number] = [120 + 60, 160 + 30];
    const shape1: [number, number] = [520 + 60, 160 + 30];
    expect(Math.abs(still[0] - shape0[0] - 80)).toBeLessThan(2);
    expect(Math.abs(still[1] - shape0[1])).toBeLessThan(2);
    expect(Math.abs(turned[0] - shape1[0] - 80 * Math.cos(Math.PI / 4))).toBeLessThan(2);
    expect(Math.abs(turned[1] - shape1[1] - 80 * Math.sin(Math.PI / 4))).toBeLessThan(2);
  });

  it('darkens the inside edge that dir points at, not the opposite one', () => {
    const inner = finding<Record<string, { peakAt: number; extent: number }>>('innerShadow');
    // The shape is 200pt wide. dir=0 darkens the right inside edge at 180pt in;
    // dir=180 darkens the left one at 0.
    expect(inner['inner-dist']!.peakAt).toBeGreaterThan(170);
    expect(inner['inner-dist180']!.peakAt).toBeLessThan(2);
  });
});

/* -------------------------------------------------------------------------- */
/* the filter graph                                                           */
/* -------------------------------------------------------------------------- */

describe('the filter graph', () => {
  const BOX: Box = { x: 0, y: 0, w: 100000, h: 100000 };
  const BLACK: Color = { space: 'rgb', hex: '000000', transforms: [] } as unknown as Color;
  const colorOf = (): { css: string; alpha: number } => ({ css: '#000000', alpha: 1 });

  it('reserves room for a shadow that lands outside the shape', () => {
    const shadow = plainShadow({ dist: 50000, dir: 0, blurRad: 30000 });
    const margin = effectMargin(shadow, BOX);
    // Thrown 50000 right, blurred with sigma 10000, three sigmas of reach.
    expect(margin.right).toBeCloseTo(80000, 0);
    expect(margin.left).toBeCloseTo(0, 0);
    expect(margin.top).toBeCloseTo(30000, 0);
  });

  it('reserves nothing for effects that only paint inside', () => {
    for (const effect of [
      { kind: 'softEdge', rad: 30000 },
      // The spread goes first. Written the other way round, plainShadow's own
      // `kind` overwrites the one being asked for and this silently tests an
      // outer shadow instead.
      { ...plainShadow({ blurRad: 30000 }), kind: 'innerShdw' },
    ] as unknown as Effect[]) {
      const margin = effectMargin(effect, BOX);
      expect(margin.left + margin.top + margin.right + margin.bottom).toBe(0);
    }
  });

  /**
   * Which effect each merge input came from, traced back through the graph.
   *
   * Asserting only the number of layers and which one is the shape does not
   * distinguish "glow above shadow" from "shadow above glow" - both give three
   * layers ending in the shape, and a mutation that swapped them killed no test
   * at all until this walked the provenance instead.
   */
  function layerOrigins(graph: ReturnType<typeof effectFilter>): string[] {
    const byResult = new Map(
      graph.primitives
        .filter((p) => 'result' in p)
        .map((p) => [(p as { result: string }).result, p]),
    );
    const originOf = (name: string, depth = 0): string => {
      if (name === 'SourceGraphic') return 'shape';
      if (name === 'SourceAlpha') return 'alpha';
      if (depth > 12) return 'deep';
      const node = byResult.get(name);
      if (node === undefined) return 'unknown';
      switch (node.op) {
        case 'morphology':
          return node.operator === 'dilate' ? 'glow' : 'softEdge';
        case 'matrix':
          return 'outerShdw';
        case 'composite':
          // A tint takes its shape from in2, so follow that rather than the flood.
          return originOf(node.in2, depth + 1);
        case 'gaussian':
        case 'offset':
          return originOf(node.in, depth + 1);
        default:
          return node.op;
      }
    };
    return graph.primitives
      .filter((p) => p.op === 'merge')
      .flatMap((p) => (p as { inputs: readonly string[] }).inputs)
      .map((name) => originOf(name));
  }

  it('paints the glow over the outer shadow, which is not the schema order', () => {
    const glow: Effect = { kind: 'glow', rad: 20000, color: BLACK };
    const shadow: Effect = plainShadow({ dist: 40000, dir: 0 });
    const graph = effectFilter([glow, shadow], BOX, colorOf);
    // Back to front: the shadow, then the glow over it, then the shape on top.
    expect(layerOrigins(graph)).toEqual(['outerShdw', 'glow', 'shape']);
    // Which is what the stacking probe saw: green over red in the band where a
    // 16pt glow and a 30pt shadow both fall.
    const stacking = finding<{ sequence: string[] }>('stacking');
    expect(stacking.sequence[0]).toBe('green');
    expect(stacking.sequence.at(-1)).toBe('red');
  });

  it('keeps that order however the file happens to list them', () => {
    const glow: Effect = { kind: 'glow', rad: 20000, color: BLACK };
    const shadow: Effect = plainShadow({ dist: 40000, dir: 0 });
    // a:effectLst is an xsd:sequence, so a file cannot express an order anyway -
    // which is exactly why the order has to come from the renderer and be the
    // same either way round.
    expect(layerOrigins(effectFilter([shadow, glow], BOX, colorOf))).toEqual(
      layerOrigins(effectFilter([glow, shadow], BOX, colorOf)),
    );
  });

  it('builds a soft edge as an erosion followed by a blur', () => {
    const graph = effectFilter([{ kind: 'softEdge', rad: 30000 }], BOX, colorOf);
    const ops = graph.primitives.map((p) => p.op);
    expect(ops).toContain('morphology');
    const erode = graph.primitives.find((p) => p.op === 'morphology');
    expect((erode as { operator: string; radius: number }).operator).toBe('erode');
    expect((erode as { radius: number }).radius).toBeCloseTo(27000, 6);
    const blur = graph.primitives.find((p) => p.op === 'gaussian');
    expect((blur as { stdDeviation: number }).stdDeviation).toBeCloseTo(10000, 6);
  });

  it('builds a glow as a dilation followed by a blur, at half the radius', () => {
    const graph = effectFilter([{ kind: 'glow', rad: 30000, color: BLACK }], BOX, colorOf);
    const dilate = graph.primitives.find((p) => p.op === 'morphology');
    expect((dilate as { operator: string }).operator).toBe('dilate');
    expect((dilate as { radius: number }).radius).toBeCloseTo(15000 * 0.9, 6);
    const blur = graph.primitives.find((p) => p.op === 'gaussian');
    expect((blur as { stdDeviation: number }).stdDeviation).toBeCloseTo(5000, 6);
  });

  it('offsets an inner shadow along dir and clips it to the shape', () => {
    const inner = {
      ...plainShadow({ dist: 20000, dir: 0 }),
      kind: 'innerShdw',
    } as unknown as Effect;
    const graph = effectFilter([inner], BOX, colorOf);
    const offset = graph.primitives.find((p) => p.op === 'offset');
    expect((offset as { dx: number; dy: number }).dx).toBeCloseTo(20000, 6);
    // Clipped back to the shape, which is what makes it inner.
    const composites = graph.primitives.filter((p) => p.op === 'composite');
    expect(composites.some((p) => (p as { in2: string }).in2 === 'SourceAlpha')).toBe(true);
  });

  it('names one result and merges every layer into it', () => {
    const graph = effectFilter(
      [
        plainShadow({ dist: 20000, dir: 0, blurRad: 10000 }),
        { kind: 'glow', rad: 10000, color: BLACK },
      ],
      BOX,
      colorOf,
    );
    const last = graph.primitives.at(-1)!;
    expect(last.op).toBe('merge');
    expect((last as { result: string }).result).toBe(graph.result);
  });
});

/* -------------------------------------------------------------------------- */
/* what PowerPoint refuses                                                    */
/* -------------------------------------------------------------------------- */

describe('what PowerPoint refuses', () => {
  it('refuses thirteen of the fifteen hostile probes', () => {
    const refusals =
      finding<Record<string, { opened: boolean; repaired: boolean | null }>>('refusals');
    const accepted = Object.entries(refusals)
      .filter(([, v]) => v.opened && v.repaired === false)
      .map(([k]) => k)
      .sort();
    expect(accepted).toEqual(['h-dashempty', 'h-dashzero']);
    expect(Object.keys(refusals)).toHaveLength(15);
  });

  it('validates the stroke vocabulary far more strictly than the fill one', () => {
    const refusals =
      finding<Record<string, { opened: boolean; repaired: boolean | null }>>('refusals');
    // Every enumeration - dash name, cap, compound, alignment, line-end type and
    // line-end size - is refused when it is out of range.
    for (const id of ['h-dashbad', 'h-capbad', 'h-cmpdbad', 'h-algnbad', 'h-endbad', 'h-endsize']) {
      expect(refusals[id]!.repaired, id).toBe(true);
    }
  });

  it('accepts an empty custDash and a zero-length segment', () => {
    // Both of which a renderer therefore has to cope with, and neither of which
    // the schema forbids.
    expect(probe('h-dashempty').repaired).toBe(false);
    expect(probe('h-dashzero').repaired).toBe(false);
    expect(dashSegments({ kind: 'custom', stops: [] })).toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* a:prstShdw, recorded rather than modelled                                  */
/* -------------------------------------------------------------------------- */

describe('a:prstShdw', () => {
  it('records all twenty presets PowerPoint writes', () => {
    const presets = finding<Record<string, { box: unknown; w: number | null }>>('presetShadows');
    expect(Object.keys(presets)).toHaveLength(20);
    for (const [id, record] of Object.entries(presets)) {
      expect(record.box, id).not.toBeNull();
    }
  });

  it('is not simply an outer shadow under another name', () => {
    const presets =
      finding<Record<string, { w: number | null; h: number | null }>>('presetShadows');
    // The shape is 70 x 60pt. A plain offset shadow would show a 60 x 60pt
    // sliver; several of these are stretched, squashed or skewed instead.
    const shapes = new Set(Object.values(presets).map((p) => `${String(p.w)}x${String(p.h)}`));
    expect(shapes.size).toBeGreaterThan(5);
  });
});

/* -------------------------------------------------------------------------- */
/* helpers                                                                    */
/* -------------------------------------------------------------------------- */

function blankLine(): Line {
  return {
    w: null,
    cap: null,
    cmpd: null,
    algn: null,
    fill: null,
    dash: null,
    join: null,
    headEnd: null,
    tailEnd: null,
  };
}

function plainShadow(
  options: {
    blurRad?: number;
    dist?: number;
    dir?: number;
    sx?: number;
    sy?: number;
    kx?: number;
    ky?: number;
    algn?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  } = {},
): OuterShadow {
  return {
    kind: 'outerShdw',
    blurRad: options.blurRad ?? 0,
    dist: options.dist ?? 0,
    dir: options.dir ?? 0,
    sx: options.sx ?? 100000,
    sy: options.sy ?? 100000,
    kx: options.kx ?? 0,
    ky: options.ky ?? 0,
    algn: options.algn ?? DEFAULT_SHADOW_ALIGN,
    rotWithShape: false,
    color: { space: 'rgb', hex: '000000', transforms: [] } as unknown as Color,
  };
}

/* Referenced so the constants above are not silently unused. */
void unpack;
void AH_W;
void JOIN_W;
void CMPD_W;

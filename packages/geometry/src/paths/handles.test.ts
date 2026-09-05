import { describe, expect, it } from 'vitest';
import { evaluateGuides, resolveOperand, type ShapeSize } from '../formula/evaluate.js';
import { FULL_CIRCLE } from '../formula/formula.js';
import {
  HANDLE_SAMPLES,
  dragHandle,
  handleCentre,
  invertAxis,
  readHandleAxis,
  resolveHandles,
  roundAdjust,
  type HandleAxisKind,
  type ResolvedHandleAxis,
} from './handles.js';
import { getPreset, presetNames } from '../presets/index.js';
import { custGeom, type Geometry, type Point, type PresetShape } from '../types.js';

/**
 * Two things are checked here and they are not the same thing.
 *
 * The first is that the inverse is an inverse: put a handle somewhere, ask what
 * adjust value would put it there, and get an answer that puts it there. That
 * is arithmetic and it is checked exhaustively, over every axis of every preset
 * at every size below.
 *
 * The second is that the *readouts* are the right ones - that a polar handle's
 * angle is an angle about the shape centre and not its `pos.y`. That cannot be
 * checked by round tripping, because a wrong readout round trips perfectly
 * well. It is checked against the presets' own arithmetic, in the manner of 2.3
 * and 2.4: ten of the eighteen angle axes state where their handle goes twice
 * over, and the two statements have to agree.
 */

/** The sizes every census below is taken at. Two square, three not, one huge. */
const SIZES: readonly ShapeSize[] = [
  { w: 200, h: 100 },
  { w: 100, h: 100 },
  { w: 100, h: 400 },
  { w: 914400, h: 685800 },
  { w: 40, h: 900 },
];

function preset(name: string): PresetShape {
  const shape = getPreset(name);
  if (shape === undefined) throw new Error(`no preset ${name}`);
  return shape;
}

function everyPreset(): PresetShape[] {
  return presetNames().map((name) => preset(name));
}

/** Every (preset, handle index, axis) triple in the corpus, at one size. */
function everyAxis(size: ShapeSize): {
  name: string;
  index: number;
  axis: ResolvedHandleAxis;
}[] {
  const out: { name: string; index: number; axis: ResolvedHandleAxis }[] = [];
  for (const shape of everyPreset()) {
    if (shape.ahLst.length === 0) continue;
    const resolved = resolveHandles(shape, size);
    resolved.forEach((handle, index) => {
      for (const axis of handle.axes) out.push({ name: shape.name, index, axis });
    });
  }
  return out;
}

const LITERAL = /^[-+]?\d+$/;

describe('what a handle is, counted', () => {
  it('120 of the 187 presets carry one, and there are 243 of them', () => {
    const shapes = everyPreset();
    const withHandles = shapes.filter((s) => s.ahLst.length > 0);
    expect(shapes).toHaveLength(187);
    expect(withHandles).toHaveLength(120);
    expect(withHandles.reduce((n, s) => n + s.ahLst.length, 0)).toBe(243);

    const histogram = new Map<number, number>();
    for (const shape of shapes) {
      const n = shape.ahLst.length;
      histogram.set(n, (histogram.get(n) ?? 0) + 1);
    }
    expect([...histogram].sort((a, b) => a[0] - b[0])).toEqual([
      [0, 67],
      [1, 49],
      [2, 36],
      [3, 19],
      [4, 15],
      [5, 1],
    ]);
  });

  it('23 handles are polar, across ten shapes', () => {
    const polar = everyPreset().filter((s) => s.ahLst.some((h) => h.kind === 'polar'));
    expect(polar.map((s) => s.name)).toEqual([
      'arc',
      'blockArc',
      'chord',
      'circularArrow',
      'donut',
      'leftCircularArrow',
      'leftRightCircularArrow',
      'mathNotEqual',
      'noSmoking',
      'pie',
    ]);
    const counts = everyPreset()
      .flatMap((s) => s.ahLst)
      .reduce(
        (acc, h) => ({ ...acc, [h.kind]: (acc[h.kind] ?? 0) + 1 }),
        {} as Record<string, number>,
      );
    expect(counts).toEqual({ xy: 220, polar: 23 });
  });

  it('287 axes, and every one names an adjust value the preset declares', () => {
    const axes = everyAxis(SIZES[0] as ShapeSize);
    expect(axes).toHaveLength(287);

    const names = new Map<string, number>();
    for (const { name, axis } of axes) {
      names.set(axis.guide, (names.get(axis.guide) ?? 0) + 1);
      // The thing a drag writes is always an `a:avLst` entry. Never a computed
      // guide, which `a:avLst` could not hold.
      expect(preset(name).avLst.map((gd) => gd.name)).toContain(axis.guide);
    }
    expect([...names].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))).toEqual([
      ['adj1', 81],
      ['adj2', 75],
      ['adj3', 40],
      ['adj', 39],
      ['adj4', 24],
      ['adj5', 12],
      ['adj6', 8],
      ['adj7', 4],
      ['adj8', 4],
    ]);
  });

  it('92 of the 287 axes take at least one bound from a guide rather than a literal', () => {
    let fromGuide = 0;
    for (const shape of everyPreset()) {
      for (const handle of shape.ahLst) {
        const tokens =
          handle.kind === 'xy'
            ? ([
                [handle.gdRefX, handle.minX, handle.maxX],
                [handle.gdRefY, handle.minY, handle.maxY],
              ] as const)
            : ([
                [handle.gdRefR, handle.minR, handle.maxR],
                [handle.gdRefAng, handle.minAng, handle.maxAng],
              ] as const);
        for (const [ref, lo, hi] of tokens) {
          if (ref === null) continue;
          const named = [lo, hi].some((t) => t !== null && !LITERAL.test(t));
          if (named) fromGuide++;
        }
      }
    }
    expect(fromGuide).toBe(92);
  });

  it('a bound is a bound on the guide, which is why some are not coordinates at all', () => {
    // `arc` bounds an angle by one turn less a unit, in 60000ths of a degree.
    const arc = preset('arc');
    expect(arc.ahLst[0]).toMatchObject({ kind: 'polar', minAng: '0', maxAng: '21599999' });
    // The callouts bound nothing: that is the largest `ST_Coordinate` there is.
    const callout = preset('accentCallout1');
    expect(callout.ahLst[0]).toMatchObject({ minX: '-2147483647', maxX: '2147483647' });
  });
});

describe('the four readouts', () => {
  it('reads a coordinate, a radius and an angle off one point', () => {
    const centre = handleCentre({ w: 200, h: 100 });
    expect(centre).toEqual({ x: 100, y: 50 });
    const pos: Point = { x: 130, y: 90 };
    expect(readHandleAxis('x', pos, centre)).toBe(130);
    expect(readHandleAxis('y', pos, centre)).toBe(90);
    expect(readHandleAxis('r', pos, centre)).toBeCloseTo(50, 10);
    // y grows downwards, so a point below and right of the centre is a positive
    // angle. 53.13 degrees, in 60000ths.
    expect(readHandleAxis('ang', pos, centre)).toBeCloseTo(53.1301 * 60000, 0);
  });

  it('the centre is the shape centre at any aspect ratio', () => {
    for (const size of SIZES) {
      expect(handleCentre(size)).toEqual({ x: size.w / 2, y: size.h / 2 });
    }
  });
});

describe('the presets state the polar origin themselves', () => {
  /** Every polar angle axis, with the worst gap between the ray and the guide. */
  function rayResiduals(): { key: string; worst: number }[] {
    const extra: readonly ShapeSize[] = [...SIZES, { w: 3, h: 1 }];
    const out: { key: string; worst: number }[] = [];
    for (const shape of everyPreset()) {
      shape.ahLst.forEach((handle, index) => {
        if (handle.kind !== 'polar' || handle.gdRefAng === null) return;
        const ref = handle.gdRefAng;
        let worst = 0;
        for (const size of extra) {
          const base = evaluateGuides(shape, size);
          const lo = handle.minAng === null ? 0 : resolveOperand(handle.minAng, base);
          const hi = handle.maxAng === null ? FULL_CIRCLE : resolveOperand(handle.maxAng, base);
          for (let k = 0; k < 49; k++) {
            const v = Math.round(lo + ((hi - lo) * k) / 48);
            const guides = evaluateGuides(shape, size, { adjust: { [ref]: v } });
            const pos: Point = {
              x: resolveOperand(handle.pos.x, guides),
              y: resolveOperand(handle.pos.y, guides),
            };
            const measured = readHandleAxis('ang', pos, handleCentre(size));
            let gap = (measured - v) % FULL_CIRCLE;
            if (gap > FULL_CIRCLE / 2) gap -= FULL_CIRCLE;
            if (gap <= -FULL_CIRCLE / 2) gap += FULL_CIRCLE;
            worst = Math.max(worst, Math.abs(gap));
          }
        }
        out.push({ key: `${shape.name}[${String(index)}]`, worst });
      });
    }
    return out;
  }

  it('ten of the eighteen angle axes put pos exactly on their own ray', () => {
    const residuals = rayResiduals();
    expect(residuals).toHaveLength(18);
    const exact = residuals.filter((r) => r.worst < 1e-6);
    expect(exact.map((r) => r.key)).toEqual([
      'arc[0]',
      'arc[1]',
      'blockArc[0]',
      'blockArc[1]',
      'chord[0]',
      'chord[1]',
      'circularArrow[1]',
      'leftCircularArrow[1]',
      'pie[0]',
      'pie[1]',
    ]);
    // One unit in the last place of an arctangent, over 49 angles at six aspect
    // ratios. Not a fit, not a tolerance chosen to pass - an identity.
    for (const r of exact) expect(r.worst).toBeLessThan(1e-7);
  });

  it('an ellipse still lands on the ray, which is what makes the origin the centre', () => {
    // At 200x100 `pie`'s handle traces an ellipse: its distance from the centre
    // runs from 50 to 100 across the sweep. The angle is exact anyway, because
    // a DrawingML angle names a ray and not an ellipse parameter.
    const pie = preset('pie');
    const size: ShapeSize = { w: 200, h: 100 };
    const radii: number[] = [];
    for (let k = 0; k < 24; k++) {
      const v = Math.round((FULL_CIRCLE * k) / 24);
      const guides = evaluateGuides(pie, size, { adjust: { adj1: v } });
      const pos: Point = {
        x: resolveOperand(pie.ahLst[0]?.pos.x ?? '', guides),
        y: resolveOperand(pie.ahLst[0]?.pos.y ?? '', guides),
      };
      expect(readHandleAxis('ang', pos, handleCentre(size))).toBeCloseTo(
        v > FULL_CIRCLE / 2 ? v - FULL_CIRCLE : v,
        3,
      );
      radii.push(readHandleAxis('r', pos, handleCentre(size)));
    }
    expect(Math.min(...radii)).toBeCloseTo(50, 5);
    expect(Math.max(...radii)).toBeCloseTo(100, 5);
  });

  it('the other eight are nowhere near their ray, so the identity is never assumed', () => {
    const off = rayResiduals().filter((r) => r.worst >= 1e-6);
    expect(off.map((r) => r.key)).toEqual([
      'circularArrow[0]',
      'circularArrow[2]',
      'leftCircularArrow[0]',
      'leftCircularArrow[2]',
      'leftRightCircularArrow[0]',
      'leftRightCircularArrow[1]',
      'leftRightCircularArrow[2]',
      'mathNotEqual[1]',
    ]);
    // Reading the angle off the mouse and writing it into the guide would be
    // right for the ten above and out by this much for these.
    const worst = Math.max(...off.map((r) => r.worst));
    expect(worst / 60000).toBeGreaterThan(150);
  });
});

describe('the forward map, with the right readout', () => {
  /** Classify one axis at one size by sampling its readout across the range. */
  function classify(
    shape: PresetShape,
    size: ShapeSize,
    index: number,
    axis: ResolvedHandleAxis,
  ): 'affine' | 'monotone' | 'non-monotone' {
    const handle = shape.ahLst[index];
    if (handle === undefined) throw new Error('no handle');
    const centre = handleCentre(size);
    const N = 41;
    const vs: number[] = [];
    const ys: number[] = [];
    for (let k = 0; k < N; k++) {
      const v = Math.round(axis.min + ((axis.max - axis.min) * k) / (N - 1));
      if (vs.includes(v)) continue;
      const guides = evaluateGuides(shape, size, { adjust: { [axis.guide]: v } });
      const pos: Point = {
        x: resolveOperand(handle.pos.x, guides),
        y: resolveOperand(handle.pos.y, guides),
      };
      vs.push(v);
      ys.push(readHandleAxis(axis.kind, pos, centre));
    }
    if (axis.kind === 'ang') {
      let turns = 0;
      for (let i = 1; i < ys.length; i++) {
        const previous = (ys[i - 1] as number) - turns * FULL_CIRCLE;
        const raw = ys[i] as number;
        if (raw - previous > FULL_CIRCLE / 2) turns--;
        else if (raw - previous < -FULL_CIRCLE / 2) turns++;
        ys[i] = raw + turns * FULL_CIRCLE;
      }
    }
    const scale = axis.kind === 'ang' ? FULL_CIRCLE / 360 : Math.max(size.w, size.h);
    const M = ys.length;
    const v0 = vs[0] as number;
    const v1 = vs[M - 1] as number;
    const y0 = ys[0] as number;
    const y1 = ys[M - 1] as number;
    const span = Math.max(...ys) - Math.min(...ys);
    let affine = v1 !== v0;
    for (let k = 0; affine && k < M; k++) {
      const want = y0 + ((y1 - y0) * ((vs[k] as number) - v0)) / (v1 - v0);
      if (Math.abs((ys[k] as number) - want) > 1e-6 * Math.max(scale, span)) affine = false;
    }
    if (affine) return 'affine';
    let up = 0;
    let down = 0;
    for (let k = 1; k < M; k++) {
      const d = (ys[k] as number) - (ys[k - 1] as number);
      if (d > 1e-9 * scale) up++;
      else if (d < -1e-9 * scale) down++;
    }
    return up > 0 && down > 0 ? 'non-monotone' : 'monotone';
  }

  it('1308 of the 1435 axis-and-size pairs are affine and only five fold back', () => {
    const counts = { affine: 0, monotone: 0, 'non-monotone': 0 };
    const folds: string[] = [];
    let affineEverywhere = 0;
    let monotoneEverywhere = 0;
    for (const shape of everyPreset()) {
      if (shape.ahLst.length === 0) continue;
      const perAxis = new Map<string, string[]>();
      for (const size of SIZES) {
        resolveHandles(shape, size).forEach((handle, index) => {
          for (const axis of handle.axes) {
            const verdict = classify(shape, size, index, axis);
            counts[verdict]++;
            const key = `${shape.name}[${String(index)}].${axis.kind}`;
            perAxis.set(key, [...(perAxis.get(key) ?? []), verdict]);
            if (verdict === 'non-monotone')
              folds.push(`${key} ${String(size.w)}x${String(size.h)}`);
          }
        });
      }
      for (const verdicts of perAxis.values()) {
        if (verdicts.every((v) => v === 'affine')) affineEverywhere++;
        if (verdicts.every((v) => v !== 'non-monotone')) monotoneEverywhere++;
      }
    }
    expect(counts).toEqual({ affine: 1308, monotone: 122, 'non-monotone': 5 });
    expect(`${String(affineEverywhere)}/${String(monotoneEverywhere)}`).toBe('243/282');
    expect(folds.sort()).toEqual([
      'circularArrow[2].ang 40x900',
      'curvedLeftArrow[0].y 100x100',
      'curvedRightArrow[0].y 100x100',
      'leftCircularArrow[2].ang 40x900',
      'leftRightCircularArrow[2].ang 40x900',
    ]);
  }, 60_000);

  it('reading a polar angle as pos.y instead would break twelve shapes', () => {
    // The point of the readout, stated as a measurement. `mathNotEqual`'s angle
    // handle has `pos.y` of `t`, so read that way it never moves at all and no
    // drag on it could mean anything.
    const shape = preset('mathNotEqual');
    const handle = shape.ahLst[1];
    if (handle === undefined || handle.kind !== 'polar') throw new Error('expected polar');
    expect(handle.pos.y).toBe('t');
    const ys = new Set<number>();
    const angles = new Set<number>();
    for (let k = 0; k <= 8; k++) {
      const v = 4200000 + ((6600000 - 4200000) * k) / 8;
      const guides = evaluateGuides(shape, { w: 200, h: 100 }, { adjust: { adj2: v } });
      const pos: Point = {
        x: resolveOperand(handle.pos.x, guides),
        y: resolveOperand(handle.pos.y, guides),
      };
      ys.add(pos.y);
      angles.add(Math.round(readHandleAxis('ang', pos, handleCentre({ w: 200, h: 100 }))));
    }
    expect(ys).toEqual(new Set([0]));
    expect(angles.size).toBe(9);
  });
});

describe('every handle inverts', () => {
  /**
   * The plan's verification - "drag each preset's handles through its full
   * `ahLst` range without NaN" - in the strongest form available.
   *
   * Absence of NaN is the floor. What is actually asserted is that the answer
   * *works*: put the handle at the position a known adjust value gives, ask for
   * the value back, and the value that comes back has to put the handle in the
   * same place. Not the same number - a flat region legitimately answers with a
   * different one - the same place.
   */
  it('through the full range of every axis at every size, and lands exactly', () => {
    let probes = 0;
    let worst = 0;
    let where = '';
    for (const shape of everyPreset()) {
      if (shape.ahLst.length === 0) continue;
      for (const size of SIZES) {
        const centre = handleCentre(size);
        const resolved = resolveHandles(shape, size);
        resolved.forEach((handle, index) => {
          const declared = shape.ahLst[index];
          if (declared === undefined) return;
          for (const axis of handle.axes) {
            for (let k = 0; k <= 8; k++) {
              const v = axis.min + ((axis.max - axis.min) * k) / 8;
              const at = resolveHandles(shape, size, { adjust: { [axis.guide]: v } })[index];
              if (at === undefined || !at.finite) continue;
              const target = readHandleAxis(axis.kind, at.pos, centre);
              const back = invertAxis(shape, size, declared, { ...axis, value: v }, target, {
                adjust: { [axis.guide]: v },
              });
              probes++;
              expect(Number.isNaN(back)).toBe(false);
              expect(Number.isFinite(back)).toBe(true);

              const after = resolveHandles(shape, size, { adjust: { [axis.guide]: back } })[index];
              if (after === undefined) continue;
              const got = readHandleAxis(axis.kind, after.pos, centre);
              const unit = axis.kind === 'ang' ? FULL_CIRCLE / 360 : Math.max(size.w, size.h);
              let error = Math.abs(got - target);
              if (axis.kind === 'ang') {
                error = Math.abs(((error + FULL_CIRCLE / 2) % FULL_CIRCLE) - FULL_CIRCLE / 2);
              }
              if (error / unit > worst) {
                worst = error / unit;
                where = `${shape.name}[${String(index)}].${axis.kind} ${String(size.w)}x${String(size.h)}`;
              }
            }
          }
        });
      }
    }
    expect(probes).toBeGreaterThan(12000);
    // Not "close enough". Every one of them, to the last bit the readout has.
    expect(`${worst.toExponential(1)} ${where}`).toBe('0.0e+0 ');
  }, 120_000);

  it('touching a handle without moving it changes nothing, on all 187', () => {
    const moved: string[] = [];
    for (const shape of everyPreset()) {
      if (shape.ahLst.length === 0) continue;
      for (const size of SIZES) {
        const resolved = resolveHandles(shape, size);
        resolved.forEach((handle, index) => {
          const declared = shape.ahLst[index];
          if (declared === undefined || !handle.finite) return;
          const next = dragHandle(shape, size, declared, handle.pos);
          for (const axis of handle.axes) {
            const wrote = next[axis.guide];
            if (wrote !== undefined && Math.abs(wrote - axis.value) > 0.5) {
              moved.push(
                `${shape.name}[${String(index)}].${axis.kind} ${String(size.w)}x${String(size.h)}: ${String(axis.value)} -> ${String(wrote)}`,
              );
            }
          }
        });
      }
    }
    expect(moved).toEqual([]);
  }, 120_000);
});

describe('dragging', () => {
  const size: ShapeSize = { w: 200, h: 100 };

  it('is what the README says it is', () => {
    // The worked example in packages/geometry/README.md, so documentation
    // cannot drift away from behaviour without a test going red.
    const roundRect = preset('roundRect');
    expect(resolveHandles(roundRect, size)).toEqual([
      {
        kind: 'xy',
        pos: { x: 16.667, y: 0 },
        axes: [
          {
            kind: 'x',
            guide: 'adj',
            value: 16667,
            min: 0,
            max: 50000,
            widened: false,
            readout: 16.667,
          },
        ],
        finite: true,
      },
    ]);
    expect(dragHandle(roundRect, size, roundRect.ahLst[0] as never, { x: 25, y: 0 })).toEqual({
      adj: 25000,
    });
  });

  it('a rounded corner follows the pointer', () => {
    const shape = preset('roundRect');
    const handle = shape.ahLst[0];
    if (handle === undefined) throw new Error('no handle');
    // `x1` is `ss * a / 100000` and `ss` is 100 here, so a quarter of the way
    // along the short side is 25000.
    expect(dragHandle(shape, size, handle, { x: 25, y: 0 })).toEqual({ adj: 25000 });
    expect(dragHandle(shape, size, handle, { x: 10, y: 0 })).toEqual({ adj: 10000 });
  });

  it('clamps rather than running past the end of the range', () => {
    const shape = preset('roundRect');
    const handle = shape.ahLst[0];
    if (handle === undefined) throw new Error('no handle');
    expect(dragHandle(shape, size, handle, { x: -500, y: 0 })).toEqual({ adj: 0 });
    expect(dragHandle(shape, size, handle, { x: 5000, y: 0 })).toEqual({ adj: 50000 });
  });

  it('writes only the adjust values the handle it was given controls', () => {
    const shape = preset('accentCallout2');
    const handle = shape.ahLst[1];
    if (handle === undefined) throw new Error('no handle');
    const next = dragHandle(shape, size, handle, { x: 40, y: 40 });
    expect(Object.keys(next).sort()).toEqual(['adj3', 'adj4']);
  });

  it('the two axes of a handle are independent, on all 220 of them', () => {
    for (const shape of everyPreset()) {
      for (const handle of shape.ahLst) {
        if (handle.kind !== 'xy' || handle.gdRefX === null || handle.gdRefY === null) continue;
        const base = evaluateGuides(shape, size);
        const lo = handle.minY === null ? 0 : resolveOperand(handle.minY, base);
        const hi = handle.maxY === null ? 0 : resolveOperand(handle.maxY, base);
        if (!Number.isFinite(lo) || !Number.isFinite(hi)) continue;
        const xAt = (v: number): number =>
          resolveOperand(
            handle.pos.x,
            evaluateGuides(shape, size, { adjust: { [handle.gdRefY as string]: v } }),
          );
        expect(xAt(lo)).toBeCloseTo(xAt(hi), 9);
      }
    }
  });

  it('an angle winds on past twelve o clock rather than snapping back', () => {
    const shape = preset('pie');
    const handle = shape.ahLst[0];
    if (handle === undefined) throw new Error('no handle');
    const square: ShapeSize = { w: 100, h: 100 };
    // Straight up is -90 degrees out of `atan2`, which is 270 as a guide.
    const up = dragHandle(shape, square, handle, { x: 50, y: 0 });
    expect(up['adj1']).toBe(16200000);
    // And straight down is 90.
    const down = dragHandle(shape, square, handle, { x: 50, y: 100 });
    expect(down['adj1']).toBe(5400000);
  });

  it('a fold is resolved towards the value the shape already holds', () => {
    // A V, built on purpose: `pos.x` is the distance of `adj` from 50000, so
    // every target in the range has two answers. Which one a drag should get is
    // not a matter of taste - a handle that teleports to the far branch when the
    // pointer crosses the bottom of the V is a bug the user sees - so the answer
    // nearer the value the shape already holds wins.
    const vee = custGeom({
      avLst: [{ name: 'adj', fmla: ['val', '10000'] }],
      gdLst: [
        { name: 'd', fmla: ['+-', 'adj', '0', '50000'] },
        { name: 'm', fmla: ['abs', 'd'] },
        { name: 'x1', fmla: ['*/', 'm', '1', '1000'] },
      ],
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: '0',
          maxX: '100000',
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'x1', y: 't' },
        },
      ],
    });
    const handle = vee.ahLst[0];
    if (handle === undefined) throw new Error('no handle');
    const square: ShapeSize = { w: 100, h: 100 };
    const axis = resolveHandles(vee, square)[0]?.axes[0];
    if (axis === undefined) throw new Error('no axis');

    const low = invertAxis(vee, square, handle, axis, 20, { near: 0 });
    const high = invertAxis(vee, square, handle, axis, 20, { near: 100000 });
    expect(low).toBeCloseTo(30000, 6);
    expect(high).toBeCloseTo(70000, 6);
    // And a drag reads the same way through the public entry point.
    expect(dragHandle(vee, square, handle, { x: 20, y: 0 })).toEqual({ adj: 30000 });
  });

  it('a fold narrower than one sample step is not seen, and that is the safe way round', () => {
    // `curvedLeftArrow` reverses over the last 1250 of a 50000 range - 2.5% of
    // it, against a scan that steps 3.1%. So the second branch is invisible
    // here, and invisible means the handle cannot jump onto it. Recorded
    // because it is a real limit of a fixed scan, not because it is desirable.
    const shape = preset('curvedLeftArrow');
    const handle = shape.ahLst[0];
    if (handle === undefined || handle.kind !== 'xy' || handle.gdRefY === null) {
      throw new Error('expected an xy handle');
    }
    const square: ShapeSize = { w: 100, h: 100 };
    const yAt = (v: number): number =>
      resolveOperand(
        handle.pos.y,
        evaluateGuides(shape, square, { adjust: { [handle.gdRefY as string]: v } }),
      );
    let turn = 0;
    for (let v = 40000; v <= 50000; v += 250) if (yAt(v) < yAt(turn)) turn = v;
    expect(turn).toBeGreaterThan(48000);
    expect((50000 - turn) / 50000).toBeLessThan(1 / (HANDLE_SAMPLES - 1));
  });
});

describe('a polar handle has to be solved in the right order', () => {
  /** Where the handle ends up if both axes are solved against the pre-drag state. */
  function simultaneous(
    name: string,
    index: number,
    size: ShapeSize,
    to: Point,
  ): Record<string, number> {
    const shape = preset(name);
    const handle = shape.ahLst[index];
    const resolved = resolveHandles(shape, size)[index];
    if (handle === undefined || resolved === undefined) throw new Error('no handle');
    const centre = handleCentre(size);
    const out: Record<string, number> = {};
    for (const axis of resolved.axes) {
      const target = readHandleAxis(axis.kind, to, centre);
      out[axis.guide] = roundAdjust(
        invertAxis(shape, size, handle, axis, target),
        axis.min,
        axis.max,
      );
    }
    return out;
  }

  function landedAt(name: string, index: number, size: ShapeSize, adjust: Record<string, number>) {
    const at = resolveHandles(preset(name), size, { adjust })[index];
    if (at === undefined) throw new Error('unresolved');
    return at.pos;
  }

  it('blockArc lands on the pointer only if the angle is solved before the radius', () => {
    // Its handle sits on the ray of the angle guide whatever the radius guide
    // says, and its distance from the centre depends on both. Solving them
    // against the same pre-drag state puts it a fifth of the shape away.
    for (const [size, to] of [
      [
        { w: 200, h: 100 },
        { x: 70, y: 20 },
      ],
      [
        { w: 100, h: 400 },
        { x: 35, y: 80 },
      ],
    ] as const) {
      const shape = preset('blockArc');
      const handle = shape.ahLst[1];
      if (handle === undefined) throw new Error('no handle');
      const scale = Math.max(size.w, size.h);

      const sequential = landedAt('blockArc', 1, size, dragHandle(shape, size, handle, to));
      const together = landedAt('blockArc', 1, size, simultaneous('blockArc', 1, size, to));
      const miss = (p: Point): number => Math.hypot(p.x - to.x, p.y - to.y) / scale;

      expect(miss(sequential)).toBeLessThan(1e-5);
      expect(miss(together)).toBeGreaterThan(0.1);
    }
  });

  it('and on a square shape the two are identical, which is why size matters here', () => {
    const shape = preset('blockArc');
    const handle = shape.ahLst[1];
    if (handle === undefined) throw new Error('no handle');
    const square: ShapeSize = { w: 100, h: 100 };
    const to: Point = { x: 35, y: 20 };
    expect(dragHandle(shape, square, handle, to)).toEqual(simultaneous('blockArc', 1, square, to));
  });

  it('the angle axis is offered first, because that is the one nothing depends on', () => {
    const axes = resolveHandles(preset('blockArc'), { w: 200, h: 100 })[1]?.axes;
    expect(axes?.map((a) => a.kind)).toEqual(['ang', 'r']);
  });
});

describe('bounds that are not bounds', () => {
  it('star24 and star32 write their handle bound in the wrong units', () => {
    // Five of the seven stars bound `adj` at 50000 - fifty per cent, in the
    // adjust's own units, and the same number their own `pin` uses. Two write
    // `ssd2`, which is half the shorter side and a coordinate.
    const bound = (name: string): string | null => {
      const handle = preset(name).ahLst[0];
      return handle !== undefined && handle.kind === 'xy' ? handle.maxY : null;
    };
    expect(['star4', 'star5', 'star8', 'star16'].map(bound)).toEqual([
      '50000',
      '50000',
      '50000',
      '50000',
    ]);
    expect([bound('star24'), bound('star32')]).toEqual(['ssd2', 'ssd2']);
    for (const name of ['star24', 'star32']) {
      expect(preset(name).gdLst[0]).toEqual({ name: 'a', fmla: ['pin', '0', 'adj', '50000'] });
      expect(preset(name).avLst[0]).toEqual({ name: 'adj', fmla: ['val', '37500'] });
    }
  });

  it('a range that excludes the shape own value is widened to admit it', () => {
    const widened: string[] = [];
    for (const shape of everyPreset()) {
      if (shape.ahLst.length === 0) continue;
      for (const size of SIZES) {
        resolveHandles(shape, size).forEach((handle, index) => {
          for (const axis of handle.axes) {
            if (axis.widened) {
              widened.push(`${shape.name}[${String(index)}].${axis.kind}`);
            }
          }
        });
      }
    }
    expect(widened).toHaveLength(16);
    expect([...new Set(widened)].sort()).toEqual([
      'leftRightRibbon[1].x',
      'mathDivide[1].y',
      'mathDivide[2].x',
      'star24[0].y',
      'star32[0].y',
    ]);
    // Taken literally, `star24`'s bound confines a handle whose guide runs to
    // 50000 within a range of 50. Widening is what keeps it usable.
    const star = resolveHandles(preset('star24'), { w: 100, h: 100 })[0]?.axes[0];
    expect(star).toMatchObject({ guide: 'adj', value: 37500, min: 0, max: 37500, widened: true });
  });

  it('mathDivide keeps its own number rather than being canonicalised', () => {
    // `adj2` is declared 5880 and the shape clamps it to 2930, so every value
    // in between draws identically. Touching the handle must not rewrite it.
    const shape = preset('mathDivide');
    const handle = shape.ahLst[1];
    if (handle === undefined) throw new Error('no handle');
    const size: ShapeSize = { w: 200, h: 100 };
    const resolved = resolveHandles(shape, size)[1];
    if (resolved === undefined) throw new Error('unresolved');
    expect(resolved.axes[0]).toMatchObject({ guide: 'adj2', value: 5880, widened: true });
    expect(dragHandle(shape, size, handle, resolved.pos)).toEqual({ adj2: 5880 });
  });
});

describe('what gets written', () => {
  it('rounds to an integer, because that is what the markup holds', () => {
    expect(roundAdjust(16666.6, 0, 50000)).toBe(16667);
    expect(roundAdjust(-0.4, 0, 50000)).toBe(0);
  });

  it('a fractional bound is not rounded past', () => {
    // `trapezoid`'s `maxAdj` is `50000*w/ss`, which is 66666.67 on an EMU-sized
    // shape, and rounding a clamped value would step outside it.
    expect(roundAdjust(70000, 0, 66666.67)).toBe(66666);
    expect(roundAdjust(0.2, 0.4, 66666.67)).toBe(1);
  });

  it('every drag writes an integer, on all 187', () => {
    for (const shape of everyPreset()) {
      if (shape.ahLst.length === 0) continue;
      for (const size of [SIZES[0] as ShapeSize, SIZES[3] as ShapeSize]) {
        resolveHandles(shape, size).forEach((handle, index) => {
          const declared = shape.ahLst[index];
          if (declared === undefined || !handle.finite) return;
          const next = dragHandle(shape, size, declared, { x: size.w * 0.3, y: size.h * 0.7 });
          for (const value of Object.values(next)) {
            expect(Number.isInteger(value)).toBe(true);
          }
        });
      }
    }
  }, 60_000);
});

describe('a shape with no size', () => {
  it('reports a handle it cannot place rather than inventing one', () => {
    const shape = preset('roundRect');
    const flat = resolveHandles(shape, { w: 0, h: 0 });
    expect(flat[0]?.pos).toEqual({ x: 0, y: 0 });
    expect(flat[0]?.finite).toBe(true);

    // A preset whose bounds divide by the size is the case that is not finite.
    const bad = custGeom({
      avLst: [{ name: 'adj', fmla: ['val', '1'] }],
      gdLst: [{ name: 'lim', fmla: ['*/', '100000', '100000', 'w'] }],
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: '0',
          maxX: 'lim',
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'adj', y: 't' },
        },
      ],
    });
    const resolved = resolveHandles(bad, { w: 0, h: 10 });
    expect(resolved[0]?.finite).toBe(false);
    expect(dragHandle(bad, { w: 0, h: 10 }, bad.ahLst[0] as never, { x: 5, y: 0 })).toEqual({});
  });
});

describe('custGeom', () => {
  const handmade: Geometry = custGeom({
    avLst: [{ name: 'adj1', fmla: ['val', '25000'] }],
    gdLst: [{ name: 'x1', fmla: ['*/', 'w', 'adj1', '100000'] }],
    ahLst: [
      {
        kind: 'xy',
        gdRefX: 'adj1',
        minX: '0',
        maxX: '100000',
        gdRefY: null,
        minY: null,
        maxY: null,
        pos: { x: 'x1', y: 'vc' },
      },
    ],
    pathLst: [],
  });

  it('a hand-authored handle goes down the same path as a preset', () => {
    const size: ShapeSize = { w: 400, h: 100 };
    const resolved = resolveHandles(handmade, size);
    expect(resolved[0]?.pos).toEqual({ x: 100, y: 50 });
    expect(resolved[0]?.axes[0]).toMatchObject({ kind: 'x', guide: 'adj1', value: 25000 });
    expect(dragHandle(handmade, size, handmade.ahLst[0] as never, { x: 300, y: 50 })).toEqual({
      adj1: 75000,
    });
  });

  it('a handle with no gdRef at all has nothing to adjust', () => {
    const inert = custGeom({
      ahLst: [
        {
          kind: 'xy',
          gdRefX: null,
          minX: null,
          maxX: null,
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'hc', y: 'vc' },
        },
      ],
    });
    const size: ShapeSize = { w: 100, h: 100 };
    expect(resolveHandles(inert, size)[0]?.axes).toEqual([]);
    expect(dragHandle(inert, size, inert.ahLst[0] as never, { x: 0, y: 0 })).toEqual({});
  });

  it('an omitted bound freezes the axis rather than freeing it', () => {
    // ISO 29500-1 on `maxX`: "If this attribute is omitted, then it is assumed
    // that this adjust handle cannot move in the x direction. That is the maxX
    // and minX are equal." No preset omits one, so this is a `custGeom` case
    // only - but freeing an axis the author froze is the wrong way to be wrong.
    for (const shape of everyPreset()) {
      for (const handle of shape.ahLst) {
        if (handle.kind === 'xy') {
          expect(handle.gdRefX === null || (handle.minX !== null && handle.maxX !== null)).toBe(
            true,
          );
          expect(handle.gdRefY === null || (handle.minY !== null && handle.maxY !== null)).toBe(
            true,
          );
        }
      }
    }

    const pinned = custGeom({
      avLst: [{ name: 'adj', fmla: ['val', '30000'] }],
      gdLst: [{ name: 'x1', fmla: ['*/', 'w', 'adj', '100000'] }],
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: null,
          maxX: null,
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'x1', y: 'vc' },
        },
      ],
    });
    const size: ShapeSize = { w: 100, h: 100 };
    expect(resolveHandles(pinned, size)[0]?.axes[0]).toMatchObject({ min: 30000, max: 30000 });
    expect(dragHandle(pinned, size, pinned.ahLst[0] as never, { x: 90, y: 50 })).toEqual({
      adj: 30000,
    });
  });

  it('an axis nothing downstream reads keeps its value instead of collapsing to the minimum', () => {
    // `gdRefX` may name a guide no part of `pos` depends on - nothing in the
    // schema ties them, and `pos` is required while every `gdRef` is optional.
    // In no preset does it happen; in a `custGeom` it can, and the readout is
    // then constant, so no drag can mean anything.
    const deaf = custGeom({
      avLst: [{ name: 'adj', fmla: ['val', '12345'] }],
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: '0',
          maxX: '100000',
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'hc', y: 'vc' },
        },
      ],
    });
    const size: ShapeSize = { w: 100, h: 100 };
    expect(dragHandle(deaf, size, deaf.ahLst[0] as never, { x: 0, y: 0 })).toEqual({ adj: 12345 });
  });

  it('bounds stated the wrong way round are read as a range, not as a failure', () => {
    const backwards = custGeom({
      avLst: [{ name: 'adj', fmla: ['val', '20000'] }],
      gdLst: [{ name: 'x1', fmla: ['*/', 'w', 'adj', '100000'] }],
      ahLst: [
        {
          kind: 'xy',
          gdRefX: 'adj',
          minX: '80000',
          maxX: '10000',
          gdRefY: null,
          minY: null,
          maxY: null,
          pos: { x: 'x1', y: 't' },
        },
      ],
    });
    const axis = resolveHandles(backwards, { w: 100, h: 100 })[0]?.axes[0];
    expect(axis).toMatchObject({ min: 10000, max: 80000 });
  });
});

describe('the search itself', () => {
  it('samples 33 points, which is what finds a fold a scan of two would miss', () => {
    expect(HANDLE_SAMPLES).toBe(33);
  });

  it('an axis whose readout cannot reach the target clamps to the nearest end', () => {
    const shape = preset('roundRect');
    const handle = shape.ahLst[0];
    const axis = resolveHandles(shape, { w: 200, h: 100 })[0]?.axes[0];
    if (handle === undefined || axis === undefined) throw new Error('no handle');
    const kinds: HandleAxisKind[] = ['x'];
    expect(kinds).toContain(axis.kind);
    expect(invertAxis(shape, { w: 200, h: 100 }, handle, axis, 1e9)).toBe(axis.max);
    expect(invertAxis(shape, { w: 200, h: 100 }, handle, axis, -1e9)).toBe(axis.min);
  });
});

describe('the evaluator takes a number as a number', () => {
  it('a fractional or very large override no longer goes through the integer grammar', () => {
    const shape = preset('roundRect');
    const size: ShapeSize = { w: 100, h: 100 };
    // Both of these threw before 2.5: the record form used to be serialised to
    // a string and re-parsed against `ST_AdjCoordinate`'s integer grammar, and
    // `String(1e21)` is "1e+21".
    expect(evaluateGuides(shape, size, { adjust: { adj: 16667.5 } }).get('adj')).toBe(16667.5);
    expect(evaluateGuides(shape, size, { adjust: { adj: 1e21 } }).get('adj')).toBe(1e21);
    // The guide-list form still evaluates its formula, which is the whole
    // reason the two forms are not the same thing.
    expect(
      evaluateGuides(shape, size, {
        adjust: [{ name: 'adj', fmla: ['*/', '100000', '1', '4'] }],
      }).get('adj'),
    ).toBe(25000);
  });
});

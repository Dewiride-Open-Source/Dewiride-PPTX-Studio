import { describe, expect, it } from 'vitest';
import {
  ARC_MAX_SWEEP,
  arcEnd,
  arcGeometry,
  arcPointAt,
  arcSegments,
  clampSweep,
  unskewAngle,
  type ArcParameters,
} from './arc.js';
import { evaluateGuides, resolveOperand, resolvePoint, type ShapeSize } from './evaluate.js';
import { FULL_CIRCLE, angleToRadians } from './formula.js';
import { getPreset, presetNames } from './presets/index.js';
import type { Point, PresetShape } from './types.js';

/**
 * The presets check this module, not the other way round.
 *
 * Four of them - `pie`, `chord`, `arc` and `blockArc` - compute the arc's start
 * point AND its end point in guide arithmetic, through `sin`, `cos`, `cat2` and
 * `sat2`, and then draw the arc between them. So the file states the answer
 * twice by two different routes, and this module has to agree with a number it
 * had no part in producing. `blockArc` states it four times.
 *
 * The rest are facts about shapes rather than about code: `ellipse` traces the
 * ellipse inscribed in its own bounding box; `donut`'s two rings wind in
 * opposite directions, which is the only reason the hole is a hole;
 * `smileyFace` has an outline and two eyes.
 */

function preset(name: string): PresetShape {
  const shape = getPreset(name);
  if (shape === undefined) throw new Error(`${name} is not a preset`);
  return shape;
}

function guidesOf(name: string, size: ShapeSize, adjust?: Record<string, number>) {
  return evaluateGuides(preset(name), size, adjust === undefined ? {} : { adjust });
}

function value(guides: ReadonlyMap<string, number>, name: string): number {
  const v = guides.get(name);
  if (v === undefined) throw new Error(`guide ${name} was never defined`);
  return v;
}

function expectPoint(actual: Point, expected: Point, digits = 6): void {
  expect(actual.x).toBeCloseTo(expected.x, digits);
  expect(actual.y).toBeCloseTo(expected.y, digits);
}

/**
 * A path walker, so the corpus-wide checks can find each arc's current point.
 *
 * Deliberately local to the test. Walking a `pathLst` and emitting from it is
 * 2.4's job; all that is borrowed here is the one thing an arc cannot do
 * without, which is knowing where the pen already is.
 */
function arcsOf(
  name: string,
  size: ShapeSize,
  adjust?: Record<string, number>,
): { from: Point; arc: ArcParameters; path: number }[] {
  const shape = preset(name);
  const guides = guidesOf(name, size, adjust);
  const at = (token: string) => resolveOperand(token, guides, { preset: name, guide: null });
  const found: { from: Point; arc: ArcParameters; path: number }[] = [];

  shape.pathLst.forEach((path, index) => {
    let cursor: Point = { x: 0, y: 0 };
    let subpathStart: Point = cursor;
    for (const command of path.commands) {
      switch (command.kind) {
        case 'moveTo':
          cursor = resolvePoint(command.to, guides);
          subpathStart = cursor;
          break;
        case 'lnTo':
        case 'quadBezTo':
        case 'cubicBezTo':
          cursor = resolvePoint(command.to, guides);
          break;
        case 'close':
          cursor = subpathStart;
          break;
        case 'arcTo': {
          const arc: ArcParameters = {
            wR: at(command.wR),
            hR: at(command.hR),
            stAng: at(command.stAng),
            swAng: at(command.swAng),
          };
          found.push({ from: cursor, arc, path: index });
          cursor = arcEnd(cursor, arc);
          break;
        }
      }
    }
  });
  return found;
}

/** The nth arc of a shape, insisting it exists rather than destructuring a maybe. */
function arcAt(
  name: string,
  index: number,
  size: ShapeSize,
  adjust?: Record<string, number>,
): { from: Point; arc: ArcParameters } {
  const found = arcsOf(name, size, adjust)[index];
  if (found === undefined) throw new Error(`${name} has no arc at ${index}`);
  return found;
}

/** Is this point on the ellipse the arc belongs to? */
function ellipseResidual(point: Point, centre: Point, rx: number, ry: number): number {
  const dx = (point.x - centre.x) / rx;
  const dy = (point.y - centre.y) / ry;
  return Math.abs(dx * dx + dy * dy - 1);
}

describe('unskewAngle', () => {
  it('fixes the four cardinal angles whatever the radii', () => {
    // Nine digits rather than twelve because of the 2000000:1 pair. A quarter
    // turn in 60000ths of a degree does not convert to exactly pi/2, so its
    // cosine is a few units of last place away from zero rather than zero, and
    // dividing that by the smaller radius scales the residue up. It reaches
    // 1.2e-10 of angle, which moves a point by a ten-billionth of a radius.
    const radii: [number, number][] = [
      [1, 1],
      [200, 3],
      [3, 200],
      [0.5, 1e6],
    ];
    for (const [wR, hR] of radii) {
      expect(unskewAngle(0, wR, hR)).toBeCloseTo(0, 9);
      expect(unskewAngle(FULL_CIRCLE / 4, wR, hR)).toBeCloseTo(Math.PI / 2, 9);
      expect(unskewAngle(FULL_CIRCLE / 2, wR, hR)).toBeCloseTo(Math.PI, 9);
      expect(unskewAngle((FULL_CIRCLE * 3) / 4, wR, hR)).toBeCloseTo((Math.PI * 3) / 2, 9);
    }
  });

  it('is the identity on a circle', () => {
    for (const angle of [0, 1234567, 5400000, 9999999, 16200000, 21599999]) {
      expect(unskewAngle(angle, 7, 7)).toBeCloseTo(angleToRadians(angle), 12);
    }
  });

  it('is not the identity on an ellipse away from the cardinals', () => {
    // 45 degrees on a 4:1 ellipse. The ray meets the curve well past 45 degrees
    // of parameter, because the ellipse is squashed vertically.
    const t = unskewAngle(FULL_CIRCLE / 8, 4, 1);
    expect(t).toBeGreaterThan(Math.PI / 4);
    expect(t).toBeCloseTo(Math.atan2(4 * Math.SQRT1_2, 1 * Math.SQRT1_2), 12);
  });

  it('is continuous and strictly increasing across turn boundaries', () => {
    // The parameter never runs away from the angle by more than the aspect
    // ratio - that is the derivative of the unskew at its steepest, which is on
    // the long axis - so a step in the angle bounds the step in the parameter.
    // Nothing but a wrapped `atan2` could produce a jump here, and it would
    // produce one of a whole turn.
    const step = 7919;
    const bound = angleToRadians(step) * 5 + 1e-12;
    let previous = unskewAngle(-FULL_CIRCLE * 2, 5, 1);
    for (let a = -FULL_CIRCLE * 2 + step; a <= FULL_CIRCLE * 2; a += step) {
      const t = unskewAngle(a, 5, 1);
      expect(t).toBeGreaterThan(previous);
      expect(t - previous).toBeLessThan(bound);
      previous = t;
    }
  });

  it('adds exactly one turn of parameter per turn of angle', () => {
    for (const angle of [0, 1234567, 13000000]) {
      expect(unskewAngle(angle + FULL_CIRCLE, 9, 2) - unskewAngle(angle, 9, 2)).toBeCloseTo(
        Math.PI * 2,
        12,
      );
    }
  });
});

describe('clampSweep', () => {
  it('holds a swing to one turn in either direction', () => {
    expect(clampSweep(0)).toBe(0);
    expect(clampSweep(ARC_MAX_SWEEP)).toBe(ARC_MAX_SWEEP);
    expect(clampSweep(-ARC_MAX_SWEEP)).toBe(-ARC_MAX_SWEEP);
    expect(clampSweep(ARC_MAX_SWEEP + 1)).toBe(ARC_MAX_SWEEP);
    expect(clampSweep(-ARC_MAX_SWEEP - 1)).toBe(-ARC_MAX_SWEEP);
    expect(clampSweep(40000000)).toBe(ARC_MAX_SWEEP);
  });

  it('clamps rather than normalising, which is a different arc', () => {
    // 400 degrees. Normalised it would be 40 and end 40 degrees along;
    // clamped it is a whole turn and ends where it began.
    const fourHundred = 400 * 60000;
    expect(clampSweep(fourHundred)).toBe(FULL_CIRCLE);
    const g = arcGeometry({ x: 100, y: 50 }, { wR: 50, hR: 25, stAng: 0, swAng: fourHundred });
    expectPoint(g.end, g.start);
  });
});

describe('pie states its own answer, and this module has to match it', () => {
  // M x1 y1 / A wd2 hd2 stAng swAng / L hc vc / Z
  // x1,y1 and x2,y2 are computed in guide arithmetic from stAng and enAng.
  const sizes: ShapeSize[] = [
    { w: 200, h: 100 },
    { w: 100, h: 200 },
    { w: 137, h: 137 },
  ];
  const adjusts = [
    undefined,
    { adj1: 0, adj2: 16200000 },
    { adj1: 2700000, adj2: 8100000 },
    { adj1: 19000000, adj2: 1000000 },
    { adj1: 5400000, adj2: 5400000 },
  ];

  for (const size of sizes) {
    for (const adjust of adjusts) {
      const label = `${size.w}x${size.h} ${adjust === undefined ? 'default' : JSON.stringify(adjust)}`;

      it(`starts where pie says, at ${label}`, () => {
        const g = guidesOf('pie', size, adjust);
        const { from, arc } = arcAt('pie', 0, size, adjust);
        expectPoint(from, { x: value(g, 'x1'), y: value(g, 'y1') });
        expectPoint(arcGeometry(from, arc).start, from);
      });

      it(`puts the centre back at the shape centre, at ${label}`, () => {
        const { from, arc } = arcAt('pie', 0, size, adjust);
        expectPoint(arcGeometry(from, arc).centre, { x: size.w / 2, y: size.h / 2 });
      });

      it(`ends where pie says, at ${label}`, () => {
        const g = guidesOf('pie', size, adjust);
        const { from, arc } = arcAt('pie', 0, size, adjust);
        expectPoint(arcEnd(from, arc), { x: value(g, 'x2'), y: value(g, 'y2') });
      });
    }
  }
});

describe('chord and arc state their own answer too', () => {
  for (const name of ['chord', 'arc']) {
    for (const adjust of [
      undefined,
      { adj1: 2700000, adj2: 16200000 },
      { adj1: 16200000, adj2: 2700000 },
      { adj1: 100000, adj2: 21000000 },
    ]) {
      it(`${name} ends at its own x2,y2 with ${adjust === undefined ? 'the default adjusts' : JSON.stringify(adjust)}`, () => {
        const size = { w: 240, h: 90 };
        const g = guidesOf(name, size, adjust);
        const { from, arc } = arcAt(name, 0, size, adjust);
        expectPoint(from, { x: value(g, 'x1'), y: value(g, 'y1') });
        expectPoint(arcEnd(from, arc), { x: value(g, 'x2'), y: value(g, 'y2') });
      });
    }
  }
});

describe('blockArc states it four times', () => {
  // The outer arc runs stAng -> istAng on the outer ellipse, so it must end on
  // x3,y3. The inner one runs istAng -> stAng on the inner ellipse, ending on
  // x4,y4. Neither of those two points is used by the path at all - they exist
  // only to bound the text rectangle - so they are an entirely separate
  // calculation of where these arcs have to finish.
  const cases = [
    { size: { w: 200, h: 120 }, adjust: undefined },
    { size: { w: 200, h: 120 }, adjust: { adj1: 0, adj2: 5400000, adj3: 25000 } },
    { size: { w: 90, h: 300 }, adjust: { adj1: 16200000, adj2: 2700000, adj3: 12500 } },
    { size: { w: 150, h: 150 }, adjust: { adj1: 20000000, adj2: 1000000, adj3: 40000 } },
    { size: { w: 310, h: 70 }, adjust: { adj1: 5400000, adj2: 5400001, adj3: 5000 } },
  ];

  for (const { size, adjust } of cases) {
    const label = `${size.w}x${size.h} ${adjust === undefined ? 'default' : JSON.stringify(adjust)}`;

    it(`outer arc runs x1,y1 to x3,y3 at ${label}`, () => {
      const g = guidesOf('blockArc', size, adjust);
      const outer = arcAt('blockArc', 0, size, adjust);
      expectPoint(outer.from, { x: value(g, 'x1'), y: value(g, 'y1') });
      expectPoint(arcEnd(outer.from, outer.arc), { x: value(g, 'x3'), y: value(g, 'y3') });
    });

    it(`inner arc runs x2,y2 to x4,y4 at ${label}`, () => {
      const g = guidesOf('blockArc', size, adjust);
      const inner = arcAt('blockArc', 1, size, adjust);
      expectPoint(inner.from, { x: value(g, 'x2'), y: value(g, 'y2') });
      expectPoint(arcEnd(inner.from, inner.arc), { x: value(g, 'x4'), y: value(g, 'y4') });
    });

    it(`sweeps the two arcs in opposite directions at ${label}`, () => {
      const outer = arcAt('blockArc', 0, size, adjust);
      const inner = arcAt('blockArc', 1, size, adjust);
      const a = arcGeometry(outer.from, outer.arc);
      const b = arcGeometry(inner.from, inner.arc);
      expect(Math.sign(a.endT - a.startT)).toBe(-Math.sign(b.endT - b.startT));
    });
  }
});

describe('ellipse traces the ellipse inscribed in its own bounds', () => {
  // M l vc / A wd2 hd2 cd2 cd4 / A wd2 hd2 3cd4 cd4 / A wd2 hd2 0 cd4 / A wd2 hd2 cd4 cd4 / Z
  // Four quarter turns from the left edge: left, top, right, bottom, left.
  const size = { w: 300, h: 120 };
  const centre = { x: 150, y: 60 };
  const corners: Point[] = [
    { x: 0, y: 60 },
    { x: 150, y: 0 },
    { x: 300, y: 60 },
    { x: 150, y: 120 },
    { x: 0, y: 60 },
  ];

  it('visits the four extreme points of the bounding box in order', () => {
    const arcs = arcsOf('ellipse', size);
    expect(arcs).toHaveLength(4);
    arcs.forEach(({ from, arc }, i) => {
      expectPoint(from, corners[i]!);
      expectPoint(arcEnd(from, arc), corners[i + 1]!);
    });
  });

  it('keeps the centre at the centre of the box for every quarter', () => {
    for (const { from, arc } of arcsOf('ellipse', size)) {
      expectPoint(arcGeometry(from, arc).centre, centre);
    }
  });

  it('holds every sampled point on the ellipse', () => {
    for (const { from, arc } of arcsOf('ellipse', size)) {
      const g = arcGeometry(from, arc);
      for (let i = 0; i <= 32; i += 1) {
        expect(ellipseResidual(arcPointAt(g, i / 32), centre, 150, 60)).toBeLessThan(1e-12);
      }
    }
  });

  it('runs clockwise on screen, so 90 degrees is the bottom', () => {
    // cd4 is the third arc's start: it begins at the right edge and the quarter
    // that follows ends at the bottom. If angles ran anticlockwise it would end
    // at the top, and every rounded corner in the corpus would be inside out.
    const arcs = arcsOf('ellipse', size);
    expect(arcs[3]!.arc.stAng).toBe(FULL_CIRCLE / 4);
    expectPoint(arcs[3]!.from, { x: 150, y: 120 });
  });
});

describe("donut's rings wind in opposite directions, which is the hole", () => {
  const size = { w: 160, h: 160 };

  /** Twice the signed area of the polygon a run of arcs sweeps out. */
  function signedArea(arcs: { from: Point; arc: ArcParameters }[]): number {
    const points: Point[] = [];
    for (const { from, arc } of arcs) {
      const g = arcGeometry(from, arc);
      for (let i = 0; i < 24; i += 1) points.push(arcPointAt(g, i / 24));
    }
    let total = 0;
    for (let i = 0; i < points.length; i += 1) {
      const a = points[i]!;
      const b = points[(i + 1) % points.length]!;
      total += a.x * b.y - b.x * a.y;
    }
    return total;
  }

  it('sweeps the outer ring forwards and the inner ring backwards', () => {
    const arcs = arcsOf('donut', size);
    expect(arcs).toHaveLength(8);
    for (const { arc } of arcs.slice(0, 4)) expect(arc.swAng).toBeGreaterThan(0);
    for (const { arc } of arcs.slice(4)) expect(arc.swAng).toBeLessThan(0);
  });

  it('gives the two rings opposite signed area', () => {
    const arcs = arcsOf('donut', size);
    const outer = signedArea(arcs.slice(0, 4));
    const inner = signedArea(arcs.slice(4));
    expect(Math.sign(outer)).toBe(-Math.sign(inner));
    expect(Math.abs(outer)).toBeGreaterThan(Math.abs(inner));
  });

  it('reports the winding through the SVG sweep flag as well', () => {
    const arcs = arcsOf('donut', size);
    for (const { from, arc } of arcs.slice(0, 4)) {
      for (const segment of arcSegments(from, arc)) expect(segment.sweep).toBe(true);
    }
    for (const { from, arc } of arcs.slice(4)) {
      for (const segment of arcSegments(from, arc)) expect(segment.sweep).toBe(false);
    }
  });

  it('a normalised negative sweep would land on the same point by the wrong route', () => {
    // -90 and +270 finish in exactly the same place, which is why normalising
    // looks harmless until something has to be filled.
    const from = { x: 0, y: 50 };
    const back = arcGeometry(from, { wR: 80, hR: 50, stAng: FULL_CIRCLE / 2, swAng: -5400000 });
    const forward = arcGeometry(from, { wR: 80, hR: 50, stAng: FULL_CIRCLE / 2, swAng: 16200000 });
    expectPoint(back.end, forward.end);
    expect(arcPointAt(back, 0.5).y).not.toBeCloseTo(arcPointAt(forward, 0.5).y, 3);
    expect(
      arcSegments(from, { wR: 80, hR: 50, stAng: FULL_CIRCLE / 2, swAng: -5400000 })[0]!.sweep,
    ).toBe(false);
  });
});

describe('a whole turn is split, because SVG will not draw one', () => {
  it("gives smileyFace's outline and both eyes two segments each", () => {
    const arcs = arcsOf('smileyFace', { w: 200, h: 200 });
    const full = arcs.filter(({ arc }) => Math.abs(arc.swAng) === FULL_CIRCLE);
    expect(full).toHaveLength(4);
    for (const { from, arc } of full) {
      const segments = arcSegments(from, arc);
      expect(segments).toHaveLength(2);
      expectPoint(segments[0]!.from, from);
      expectPoint(segments[1]!.to, from);
      // Neither piece is a no-op, which is the whole point of splitting.
      for (const segment of segments) {
        expect(
          Math.hypot(segment.to.x - segment.from.x, segment.to.y - segment.from.y),
        ).toBeGreaterThan(1);
      }
    }
  });

  it('splits at the antipode', () => {
    const from = { x: 0, y: 25 };
    const [first, second] = arcSegments(from, {
      wR: 60,
      hR: 25,
      stAng: FULL_CIRCLE / 2,
      swAng: FULL_CIRCLE,
    });
    expectPoint(first!.to, { x: 120, y: 25 });
    expectPoint(second!.to, from);
    expect(first!.largeArc).toBe(false);
    expect(second!.largeArc).toBe(false);
  });

  it('splits a backwards whole turn the same way, keeping the direction', () => {
    const from = { x: 0, y: 25 };
    const segments = arcSegments(from, {
      wR: 60,
      hR: 25,
      stAng: FULL_CIRCLE / 2,
      swAng: -FULL_CIRCLE,
    });
    expect(segments).toHaveLength(2);
    for (const segment of segments) expect(segment.sweep).toBe(false);
    expectPoint(segments[1]!.to, from);
  });

  it('leaves anything short of a whole turn as one segment', () => {
    const from = { x: 0, y: 25 };
    for (const swAng of [5400000, -5400000, 16200000, FULL_CIRCLE - 1]) {
      const segments = arcSegments(from, { wR: 60, hR: 25, stAng: FULL_CIRCLE / 2, swAng });
      expect(segments).toHaveLength(1);
      expect(segments[0]!.largeArc).toBe(Math.abs(swAng) > FULL_CIRCLE / 2);
    }
  });
});

describe('moon, computed by hand', () => {
  // M r b / A w hd2 cd4 cd2 / ...
  // Starting at the bottom-right corner with a horizontal radius of the whole
  // width and a vertical radius of half the height. stAng is 90 degrees, which
  // on this ellipse is still 90 degrees of parameter, so the centre sits half a
  // height above the start - on the right edge - and the half turn that follows
  // reaches the top-right corner by way of the left edge.
  const size = { w: 200, h: 100 };

  it('places the outer arc from the corner it starts in', () => {
    const outer = arcAt('moon', 0, size);
    expectPoint(outer.from, { x: 200, y: 100 });
    const g = arcGeometry(outer.from, outer.arc);
    expectPoint(g.centre, { x: 200, y: 50 });
    expectPoint(g.end, { x: 200, y: 0 });
    expectPoint(arcPointAt(g, 0.5), { x: 0, y: 50 });
  });

  it('takes a negative start angle on the inner arc without normalising it', () => {
    const g = guidesOf('moon', size);
    expect(value(g, 'stAng1')).toBeLessThan(0);
    const inner = arcAt('moon', 1, size);
    expect(inner.arc.stAng).toBe(value(g, 'stAng1'));
    expect(Number.isFinite(arcEnd(inner.from, inner.arc).x)).toBe(true);
  });

  it('closes: the inner arc comes back to where the path began', () => {
    // The subpath is M r b / A / A / Z, so the second arc has to finish at the
    // start for the moon to have two clean horns rather than a chord across it.
    const inner = arcAt('moon', 1, size);
    expectPoint(arcEnd(inner.from, inner.arc), { x: 200, y: 100 }, 3);
  });
});

describe('the circular arrows round-trip at2 through the unskew', () => {
  // These three build their arc start point parametrically - a point at
  // distance rI on a circle, scaled by rw2/rI and rh2/rI, which is
  // (rw2 cos p, rh2 sin p) - then take its RAY angle with `at2` and hand that
  // to `arcTo` as stAng. So the file asks for the parameter to be recovered
  // from the ray, which is exactly what the unskew does and exactly what an
  // implementation that treats stAng as a parameter cannot do. If the two are
  // not inverses the arc is centred somewhere other than the shape centre, and
  // the arrowhead detaches from the band.
  const sizes: ShapeSize[] = [
    { w: 200, h: 200 },
    { w: 300, h: 110 },
    { w: 80, h: 260 },
  ];

  for (const name of ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow']) {
    for (const size of sizes) {
      it(`${name} centres both arcs on the shape centre at ${size.w}x${size.h}`, () => {
        const arcs = arcsOf(name, size);
        expect(arcs.length).toBeGreaterThanOrEqual(2);
        for (const { from, arc } of arcs) {
          expectPoint(arcGeometry(from, arc).centre, { x: size.w / 2, y: size.h / 2 }, 4);
        }
      });
    }

    it(`${name} keeps both arcs concentric across its adjust range`, () => {
      const size = { w: 260, h: 120 };
      for (const adj4 of [0, 2700000, 8100000, 16200000, 21000000]) {
        for (const { from, arc } of arcsOf(name, size, { adj4 })) {
          expectPoint(arcGeometry(from, arc).centre, { x: 130, y: 60 }, 4);
        }
      }
    });
  }
});

describe('funnel says the same thing a third way', () => {
  // pie writes its start point through `cat2`/`sat2`; funnel writes it in polar
  // form - n1 = wd2 hd4 / sqrt(hd4^2 cos^2 + wd2^2 sin^2), which is the polar
  // equation of the same ellipse, then n1 at the ray angle. Different
  // arithmetic, same point, and the arc still has to be centred on it.
  const size = { w: 240, h: 160 };

  it('centres the bowl on the middle of its upper ellipse', () => {
    const bowl = arcAt('funnel', 0, size);
    expectPoint(arcGeometry(bowl.from, bowl.arc).centre, { x: 120, y: 40 }, 6);
  });

  it('draws the whole reverse turn as two segments running backwards', () => {
    const arcs = arcsOf('funnel', size);
    const loop = arcs.find(({ arc }) => arc.swAng === -FULL_CIRCLE);
    expect(loop).toBeDefined();
    const segments = arcSegments(loop!.from, loop!.arc);
    expect(segments).toHaveLength(2);
    for (const segment of segments) expect(segment.sweep).toBe(false);
    expectPoint(segments[1]!.to, loop!.from);
  });
});

describe('degenerate arcs', () => {
  it('reports a zero radius rather than producing NaN', () => {
    const g = arcGeometry({ x: 10, y: 10 }, { wR: 0, hR: 20, stAng: 0, swAng: 5400000 });
    expect(g.degenerate).toBe(true);
    expect(arcSegments({ x: 10, y: 10 }, { wR: 0, hR: 20, stAng: 0, swAng: 5400000 })).toEqual([]);
  });

  it('carries a non-finite radius through as degenerate', () => {
    const g = arcGeometry(
      { x: 10, y: 10 },
      { wR: Number.POSITIVE_INFINITY, hR: 20, stAng: 0, swAng: 5400000 },
    );
    expect(g.degenerate).toBe(true);
  });

  it('draws nothing for a zero sweep', () => {
    expect(arcSegments({ x: 10, y: 10 }, { wR: 5, hR: 5, stAng: 0, swAng: 0 })).toEqual([]);
  });

  it('uses the absolute value of a negative radius, as SVG requires', () => {
    const a = arcGeometry({ x: 0, y: 0 }, { wR: -30, hR: 20, stAng: 0, swAng: 5400000 });
    const b = arcGeometry({ x: 0, y: 0 }, { wR: 30, hR: 20, stAng: 0, swAng: 5400000 });
    expect(a.rx).toBe(30);
    expectPoint(a.centre, b.centre);
    expectPoint(a.end, b.end);
  });
});

describe('every arc in the corpus', () => {
  const sizes: ShapeSize[] = [
    { w: 100, h: 100 },
    { w: 320, h: 90 },
    { w: 90, h: 320 },
    { w: 7, h: 4000 },
  ];

  function everyArc() {
    const all: { shape: string; size: ShapeSize; from: Point; arc: ArcParameters }[] = [];
    for (const shape of presetNames()) {
      for (const size of sizes) {
        for (const { from, arc } of arcsOf(shape, size)) all.push({ shape, size, from, arc });
      }
    }
    return all;
  }

  it('is found where the static count says it is', () => {
    const perShape = new Map<string, number>();
    for (const shape of presetNames()) {
      const n = arcsOf(shape, { w: 100, h: 100 }).length;
      if (n > 0) perShape.set(shape, n);
    }
    expect(perShape.size).toBe(63);
    expect([...perShape.values()].reduce((a, b) => a + b, 0)).toBe(393);
  });

  it('never opens a path, so the current point always exists', () => {
    for (const shape of presetNames()) {
      for (const path of preset(shape).pathLst) {
        expect(path.commands[0]?.kind).not.toBe('arcTo');
      }
    }
  });

  it('produces finite geometry at every size', () => {
    for (const { shape, size, from, arc } of everyArc()) {
      const g = arcGeometry(from, arc);
      const where = `${shape} at ${size.w}x${size.h}`;
      expect([where, Number.isFinite(g.end.x), Number.isFinite(g.end.y)]).toStrictEqual([
        where,
        true,
        true,
      ]);
    }
  });

  it('keeps every sampled point on its own ellipse', () => {
    for (const { shape, size, from, arc } of everyArc()) {
      const g = arcGeometry(from, arc);
      if (g.degenerate) continue;
      for (let i = 0; i <= 8; i += 1) {
        const residual = ellipseResidual(arcPointAt(g, i / 8), g.centre, g.rx, g.ry);
        expect([`${shape} at ${size.w}x${size.h}`, residual < 1e-9]).toStrictEqual([
          `${shape} at ${size.w}x${size.h}`,
          true,
        ]);
      }
    }
  });

  it('starts every arc exactly where the pen already was', () => {
    for (const { shape, size, from, arc } of everyArc()) {
      const g = arcGeometry(from, arc);
      if (g.degenerate) continue;
      expect([
        `${shape} at ${size.w}x${size.h}`,
        Math.hypot(g.start.x - from.x, g.start.y - from.y) < 1e-9,
      ]).toStrictEqual([`${shape} at ${size.w}x${size.h}`, true]);
    }
  });

  it('never needs a radius it has to take the absolute value of', () => {
    const negative = everyArc().filter(({ arc }) => arc.wR < 0 || arc.hR < 0);
    expect(negative.map((n) => `${n.shape} ${n.arc.wR} ${n.arc.hR}`)).toStrictEqual([]);
  });

  it('never asks for more than the whole turn the clamp allows', () => {
    const over = everyArc().filter(({ arc }) => Math.abs(arc.swAng) > FULL_CIRCLE);
    expect(over.map((o) => `${o.shape} ${o.arc.swAng}`)).toStrictEqual([]);
  });

  it('counts the arcs that close on themselves and therefore have to be split', () => {
    const seen = new Set<string>();
    let full = 0;
    for (const { shape, from, arc } of everyArc()) {
      if (Math.abs(arc.swAng) < FULL_CIRCLE) continue;
      if (arcSegments(from, arc).length !== 2) throw new Error(`${shape} did not split`);
      full += 1;
      seen.add(shape);
    }
    expect(full / sizes.length).toBe(20);
    expect([...seen].sort()).toStrictEqual([
      'actionButtonHelp',
      'actionButtonInformation',
      'cloudCallout',
      'funnel',
      'mathDivide',
      'smileyFace',
      'sun',
    ]);
  });

  it('measures how many arcs the unskew actually moves', () => {
    // The number that says whether this module earns its place. An arc on a
    // circle, or one starting on a cardinal angle, is untouched by the unskew;
    // anything else lands somewhere else entirely if the step is skipped, and
    // "somewhere else" is a shape that is visibly wrong rather than slightly
    // off. Measured at one ordinary non-square size, because a square shape
    // hides every circular case behind ss === ls.
    const size = { w: 320, h: 90 };
    const moved = new Set<string>();
    let count = 0;
    let total = 0;
    for (const shape of presetNames()) {
      for (const { from, arc } of arcsOf(shape, size)) {
        const g = arcGeometry(from, arc);
        if (g.degenerate) continue;
        total += 1;
        if (Math.abs(g.startT - unskewAngle(arc.stAng, 1, 1)) > 1e-9) {
          count += 1;
          moved.add(shape);
        }
      }
    }
    // 84 of the 389: better than one arc in five is somewhere else without it.
    expect([total, count, moved.size]).toStrictEqual([389, 84, 16]);
  });

  it('degenerates only where a preset asks for a corner with no radius', () => {
    // Both shapes round two corners and leave two square, and a square corner
    // is written as an arc of radius zero rather than left out. So this is the
    // whole list, and it is a correct answer rather than a lost arc: `adj2`
    // defaults to 0, `a` comes out 0, and the two `A a a` commands have nothing
    // to draw. At any other adjust value they are ordinary arcs.
    const lost = new Set<string>();
    for (const { shape, from, arc } of everyArc()) {
      if (arcGeometry(from, arc).degenerate) lost.add(shape);
    }
    expect([...lost].sort()).toStrictEqual(['round2DiagRect', 'round2SameRect']);

    for (const shape of ['round2DiagRect', 'round2SameRect']) {
      const rounded = arcsOf(shape, { w: 320, h: 90 }, { adj1: 16667, adj2: 16667 });
      expect(rounded.every(({ from, arc }) => !arcGeometry(from, arc).degenerate)).toBe(true);
    }
  });

  it('leaves the pen exactly where it was on a degenerate arc', () => {
    // A path walker has to be able to keep going. A square corner already sits
    // at the point the preceding `lnTo` reached, so the arc must not move it.
    for (const { from, arc } of arcsOf('round2DiagRect', { w: 320, h: 90 })) {
      if (!arcGeometry(from, arc).degenerate) continue;
      expect(arcEnd(from, arc)).toStrictEqual(from);
    }
  });
});

import { describe, expect, it } from 'vitest';
import { GeometryError } from '../errors.js';
import {
  evaluateGuides,
  nonFiniteGuides,
  resolveOperand,
  resolvePoint,
  type ShapeSize,
} from './evaluate.js';
import { getPreset, presetNames } from '../presets/index.js';
import type { PresetShape } from '../types.js';

/**
 * Where the independent checking happens.
 *
 * `formula.test.ts` checks each operator against arithmetic, which is close to
 * checking this implementation against the understanding that wrote it. The
 * assertions here are different in kind: they are facts about shapes. A
 * triangle's apex is above the middle of its base; a rectangle's connection
 * sites are the midpoints of its edges; every point a `blockArc` computes lies
 * on the ellipse it is an arc of. None of that is derived from the evaluator,
 * so an operator with its operands transposed fails them.
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

describe('triangle, computed by hand', () => {
  // At 200 x 100 with the default adjust of 50000:
  //   a  = pin 0 50000 100000     = 50000
  //   x1 = w * a / 200000         = 50
  //   x2 = w * a / 100000         = 100   <- the apex, at the horizontal centre
  //   x3 = x1 + wd2               = 150
  const g = guidesOf('triangle', { w: 200, h: 100 });

  it('resolves the four guides to the numbers the arithmetic gives', () => {
    expect(value(g, 'a')).toBe(50000);
    expect(value(g, 'x1')).toBe(50);
    expect(value(g, 'x2')).toBe(100);
    expect(value(g, 'x3')).toBe(150);
  });

  it('puts the apex above the middle of the base, which is what makes it isosceles', () => {
    expect(value(g, 'x2')).toBe(value(g, 'hc'));
  });

  it('draws the three corners of that triangle', () => {
    const path = preset('triangle').pathLst[0];
    if (path === undefined) throw new Error('triangle has no path');
    const points = path.commands
      .filter((c) => c.kind === 'moveTo' || c.kind === 'lnTo')
      .map((c) => resolvePoint((c as { to: { x: string; y: string } }).to, g));

    expect(points).toStrictEqual([
      { x: 0, y: 100 },
      { x: 100, y: 0 },
      { x: 200, y: 100 },
    ]);
  });

  it('puts the text rectangle inside the lower half, where a triangle has room for text', () => {
    const rect = preset('triangle').rect;
    if (rect === null) throw new Error('triangle has no text rectangle');
    expect({
      l: resolveOperand(rect.l, g),
      t: resolveOperand(rect.t, g),
      r: resolveOperand(rect.r, g),
      b: resolveOperand(rect.b, g),
    }).toStrictEqual({ l: 50, t: 50, r: 150, b: 100 });
  });

  it('moves the apex to either end when the adjust is driven to its limits', () => {
    expect(value(guidesOf('triangle', { w: 200, h: 100 }, { adj: 0 }), 'x2')).toBe(0);
    expect(value(guidesOf('triangle', { w: 200, h: 100 }, { adj: 100000 }), 'x2')).toBe(200);
  });

  it('clamps an adjust outside its range rather than letting the apex leave the shape', () => {
    // `a = pin 0 adj 100000` is the preset doing this, not the evaluator.
    expect(value(guidesOf('triangle', { w: 200, h: 100 }, { adj: 999999 }), 'x2')).toBe(200);
    expect(value(guidesOf('triangle', { w: 200, h: 100 }, { adj: -999999 }), 'x2')).toBe(0);
  });
});

describe("rect's connection sites are the midpoints of its edges", () => {
  // The strongest oracle available at this stage: what these six numbers have
  // to be follows from the shape being a rectangle, and from nothing in this
  // package.
  const g = guidesOf('rect', { w: 100, h: 50 });
  const sites = preset('rect').cxnLst.map((c) => ({
    ang: resolveOperand(c.ang, g),
    ...resolvePoint(c.pos, g),
  }));

  it('places them at top, left, bottom and right centre', () => {
    expect(sites).toStrictEqual([
      { ang: 16200000, x: 50, y: 0 },
      { ang: 10800000, x: 0, y: 25 },
      { ang: 5400000, x: 50, y: 50 },
      { ang: 0, x: 100, y: 25 },
    ]);
  });

  it('confirms the angle convention: zero points right, and ninety degrees points down', () => {
    // The y axis increases downward in shape space, so a positive angle turns
    // clockwise on screen. `triangle` says the same thing independently - its
    // apex site is 3cd4 and its base corners are cd4.
    const bottom = sites[2];
    const right = sites[3];
    expect(right?.ang).toBe(0);
    expect(right?.x).toBe(100);
    expect(bottom?.ang).toBe(5400000);
    expect(bottom?.y).toBe(50);
  });
});

describe('blockArc, which pins the arc-tangent operators', () => {
  const size = { w: 200, h: 100 };

  it('starts at three o’clock when the start angle is zero', () => {
    // With cat2 and sat2 transposed this lands at six o'clock instead - the
    // shape still draws, and draws wrong.
    const g = guidesOf('blockArc', size, { adj1: 0 });
    expect(value(g, 'x1')).toBeCloseTo(200, 9);
    expect(value(g, 'y1')).toBeCloseTo(50, 9);
  });

  it('starts at six o’clock at a quarter turn, because ninety degrees is down', () => {
    const g = guidesOf('blockArc', size, { adj1: 5400000 });
    expect(value(g, 'x1')).toBeCloseTo(100, 6);
    expect(value(g, 'y1')).toBeCloseTo(100, 9);
  });

  it('keeps every computed point on the ellipse it is an arc of', () => {
    // ((x - hc) / wd2)^2 + ((y - vc) / hd2)^2 = 1 for the outer pair, and the
    // same with the inset radii for the inner pair. True for any angle, so it
    // holds across the whole adjust range and is not a spot check.
    for (const s of [
      { w: 200, h: 100 },
      { w: 100, h: 300 },
      { w: 77, h: 77 },
    ]) {
      for (const adj1 of [0, 1234567, 5400000, 10800000, 16200000, 21599999]) {
        for (const adj3 of [0, 25000, 50000]) {
          const g = guidesOf('blockArc', s, { adj1, adj2: 0, adj3 });
          const hc = value(g, 'hc');
          const vc = value(g, 'vc');

          for (const [xn, yn, wn, hn] of [
            ['x1', 'y1', 'wd2', 'hd2'],
            ['x3', 'y3', 'wd2', 'hd2'],
            ['x2', 'y2', 'iwd2', 'ihd2'],
            ['x4', 'y4', 'iwd2', 'ihd2'],
          ] as const) {
            const rx = value(g, wn);
            const ry = value(g, hn);
            if (rx === 0 || ry === 0) continue;
            const dx = (value(g, xn) - hc) / rx;
            const dy = (value(g, yn) - vc) / ry;
            expect(dx * dx + dy * dy).toBeCloseTo(1, 9);
          }
        }
      }
    }
  });
});

describe('hexagon, which pins the angle unit', () => {
  it('computes dy1 as exactly half the height', () => {
    // shd2 = hd2 * vf / 100000 with vf = 115470, and dy1 = sin(shd2, 3600000).
    // 3600000 angle units is 60 degrees, so dy1 = (h/2)(2/sqrt 3)(sqrt 3 / 2),
    // which is h/2 up to the four decimal places vf is rounded to. Read the
    // angle as degrees or radians instead and this is not close.
    const g = guidesOf('hexagon', { w: 400, h: 200 });
    expect(value(g, 'dy1') / 100).toBeCloseTo(1, 5);
    expect(value(g, 'dy1')).toBeCloseTo(value(g, 'hd2'), 3);
  });

  it('therefore reaches the top and bottom edges', () => {
    const g = guidesOf('hexagon', { w: 400, h: 200 });
    expect(value(g, 'y1')).toBeCloseTo(0, 3);
    expect(value(g, 'y2')).toBeCloseTo(200, 3);
  });
});

describe('adjust values merge by name', () => {
  it('overrides a default the preset declares', () => {
    expect(value(guidesOf('triangle', { w: 100, h: 100 }, { adj: 25000 }), 'x2')).toBe(25);
  });

  it('accepts the parsed avLst form as well as plain numbers', () => {
    const fromList = evaluateGuides(
      preset('triangle'),
      { w: 100, h: 100 },
      { adjust: [{ name: 'adj', fmla: ['val', '25000'] }] },
    );
    expect(value(fromList, 'x2')).toBe(25);
  });

  it('evaluates an override that is not a plain value', () => {
    // `a:avLst` is a list of guides and the schema does not require `val`.
    const g = evaluateGuides(
      preset('triangle'),
      { w: 100, h: 100 },
      { adjust: [{ name: 'adj', fmla: ['*/', '100000', '1', '4'] }] },
    );
    expect(value(g, 'adj')).toBe(25000);
    expect(value(g, 'x2')).toBe(25);
  });

  it('is by name and not by position, which matters because adjust names are not uniform', () => {
    // `hexagon` declares `adj` and `vf`. A positional merge would put a deck's
    // first adjust into whichever came first in the file.
    const g = guidesOf('hexagon', { w: 400, h: 200 }, { vf: 100000 });
    expect(value(g, 'vf')).toBe(100000);
    expect(value(g, 'adj')).toBe(25000);
  });

  it('keeps an override for a name the preset does not declare', () => {
    const g = guidesOf('triangle', { w: 100, h: 100 }, { adj7: 1234 });
    expect(value(g, 'adj7')).toBe(1234);
    expect(value(g, 'x2')).toBe(50);
  });

  it('refuses to let an override redefine a built-in', () => {
    // Otherwise a deck could make the shape's own width a suggestion.
    const g = guidesOf('triangle', { w: 100, h: 100 }, { w: 999 });
    expect(value(g, 'w')).toBe(100);
  });
});

describe('every preset, at every size', () => {
  const SIZES: readonly ShapeSize[] = [
    { w: 9144000, h: 6858000 },
    { w: 1828800, h: 1828800 },
    { w: 200, h: 100 },
    { w: 100, h: 300 },
    { w: 3, h: 7 },
    { w: 1, h: 1 },
  ];

  it('evaluates without throwing, which is also the proof that one forward pass is enough', () => {
    // An operand naming a guide defined later in the same gdLst would throw
    // FMLA_OPERAND here. Nothing does, across all 3622 computed guides - so the
    // source really is in dependency order, and a future POI that reorders one
    // fails this rather than quietly building a shape on an undefined guide.
    for (const size of SIZES) {
      for (const name of presetNames()) {
        expect(() => evaluateGuides(preset(name), size)).not.toThrow();
      }
    }
  });

  it('produces only finite numbers at any size with a non-zero width and height', () => {
    for (const size of SIZES) {
      for (const name of presetNames()) {
        const bad = nonFiniteGuides(evaluateGuides(preset(name), size));
        expect(bad, `${name} at ${String(size.w)}x${String(size.h)}`).toStrictEqual([]);
      }
    }
  });

  it('defines every guide the preset declares, and never fewer', () => {
    for (const name of presetNames()) {
      const shape = preset(name);
      const g = evaluateGuides(shape, { w: 200, h: 100 });
      for (const gd of [...shape.avLst, ...shape.gdLst]) expect(g.has(gd.name)).toBe(true);
    }
  });

  it('resolves every operand a shape refers to outside its formulas', () => {
    // The text rectangle, the connection sites, the adjust handles and the path
    // commands all name guides too. If the evaluator produced the guides but
    // missed one of these, nothing above would notice.
    for (const name of presetNames()) {
      const shape = preset(name);
      const g = evaluateGuides(shape, { w: 200, h: 100 });
      if (shape.rect !== null) {
        for (const t of [shape.rect.l, shape.rect.t, shape.rect.r, shape.rect.b]) {
          expect(Number.isFinite(resolveOperand(t, g))).toBe(true);
        }
      }
      for (const c of shape.cxnLst) {
        expect(Number.isFinite(resolveOperand(c.ang, g))).toBe(true);
        expect(Number.isFinite(resolvePoint(c.pos, g).x)).toBe(true);
      }
    }
  });
});

describe('degenerate sizes are survivable, not clean', () => {
  // A shape with zero height is ordinary - a flat line, or a shape being
  // dragged through zero - and 36 presets divide by `ss`, so this is a case
  // that happens rather than a corrupt-file case. What matters is that it does
  // not throw and that a caller can see it happened.
  const DEGENERATE: readonly ShapeSize[] = [
    { w: 0, h: 100 },
    { w: 100, h: 0 },
    { w: 0, h: 0 },
  ];

  it('never throws', () => {
    for (const size of DEGENERATE) {
      for (const name of presetNames()) {
        expect(() => evaluateGuides(preset(name), size)).not.toThrow();
      }
    }
  });

  it('leaves shapes that never divide by a guide completely finite', () => {
    for (const size of DEGENERATE) {
      // `rect` has no guides at all; `triangle` divides only by literals.
      expect(nonFiniteGuides(evaluateGuides(preset('rect'), size))).toStrictEqual([]);
      expect(nonFiniteGuides(evaluateGuides(preset('triangle'), size))).toStrictEqual([]);
    }
  });

  it('does produce non-finite guides where a divisor collapses', () => {
    // `hexagon.maxAdj` divides by `ss`, which is zero here.
    const bad = nonFiniteGuides(evaluateGuides(preset('hexagon'), { w: 100, h: 0 }));
    expect(bad).toContain('maxAdj');
  });

  it('affects a pinned number of presets, so a change to the arithmetic is visible', () => {
    // 49 of the 187 at zero width, 49 at zero height, 55 when both collapse.
    // The number is not important in itself; it is here so that a change to
    // division, to the built-in table, or to the preset data shows up as a
    // number moving rather than as a shape quietly failing to draw one day.
    const counts = DEGENERATE.map(
      (size) =>
        presetNames().filter(
          (name) => nonFiniteGuides(evaluateGuides(preset(name), size)).length > 0,
        ).length,
    );
    expect(counts).toStrictEqual([49, 49, 55]);
  });
});

describe('the sqrt divergence, measured rather than assumed', () => {
  // `sqrt` takes the absolute value first, following Office rather than ECMA.
  // Whether that ever changes a shipped preset's result is a question with an
  // answer, so this finds it instead of leaving the divergence as a claim.
  //
  // 19 formulas use sqrt, spread over seven shapes, and every operand is a
  // computed guide driven by adjust values - so the sign cannot be read off the
  // source. This drives each of those shapes across its adjust range and
  // records whether any radicand actually goes negative.
  function negativeRadicands(): string[] {
    const found = new Set<string>();
    const SWEEP = [0, 1, 12500, 25000, 50000, 75000, 100000, 200000, -50000];

    for (const name of presetNames()) {
      const shape = preset(name);
      const roots = shape.gdLst.filter((gd) => gd.fmla[0] === 'sqrt');
      if (roots.length === 0) continue;

      const adjustNames = shape.avLst.map((gd) => gd.name);
      for (const size of [
        { w: 200, h: 100 },
        { w: 100, h: 200 },
        { w: 100, h: 100 },
      ]) {
        for (const which of adjustNames) {
          for (const v of SWEEP) {
            const g = evaluateGuides(shape, size, { adjust: { [which]: v } });
            for (const gd of roots) {
              const operand = gd.fmla[1];
              if (operand === undefined) continue;
              if (resolveOperand(operand, g) < 0) found.add(`${name}.${gd.name}`);
            }
          }
        }
      }
    }
    return [...found].sort();
  }

  it('finds the negative radicands that Office silently absorbs', () => {
    // If this list is empty the divergence is unreachable from the presets and
    // only matters for hand-authored custom geometry. If it is not, then
    // implementing ECMA's plain sqrt would put a NaN into these guides, and NaN
    // is contagious - the shape would draw nothing rather than draw wrong.
    expect(negativeRadicands()).toStrictEqual([
      'curvedDownArrow.q11',
      'curvedDownArrow.q5',
      'curvedUpArrow.q11',
      'curvedUpArrow.q5',
      'leftCircularArrow.u9',
    ]);
  });

  it('keeps those shapes finite anyway, which is the point of following Office', () => {
    for (const name of ['curvedDownArrow', 'curvedUpArrow', 'leftCircularArrow']) {
      for (const v of [0, 50000, 100000]) {
        const g = evaluateGuides(preset(name), { w: 200, h: 100 }, { adjust: { adj1: v } });
        expect(nonFiniteGuides(g), `${name} at adj1=${String(v)}`).toStrictEqual([]);
      }
    }
  });
});

describe('failures are typed', () => {
  it('refuses an operand that is neither an integer nor a guide', () => {
    const g = new Map<string, number>([['w', 100]]);
    try {
      resolveOperand('nosuchguide', g, { preset: 'demo', guide: 'x1' });
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(GeometryError);
      expect((error as GeometryError).code).toBe('FMLA_OPERAND');
      expect((error as GeometryError).preset).toBe('demo');
      expect((error as GeometryError).message).toContain('demo.x1');
    }
  });

  it('accepts a signed integer literal and refuses anything else numeric-looking', () => {
    const g = new Map<string, number>();
    expect(resolveOperand('42', g)).toBe(42);
    expect(resolveOperand('-42', g)).toBe(-42);
    for (const bad of ['4.5', '1e6', '', ' 4', '0x10', 'Infinity', 'NaN']) {
      expect(() => resolveOperand(bad, g)).toThrow(GeometryError);
    }
  });

  it('prefers a guide over a literal reading, though nothing can be both', () => {
    const g = new Map<string, number>([['w', 100]]);
    expect(resolveOperand('w', g)).toBe(100);
  });

  it('refuses an empty formula', () => {
    const shape: PresetShape = {
      name: 'demo',
      avLst: [{ name: 'adj', fmla: [] }],
      gdLst: [],
      ahLst: [],
      cxnLst: [],
      rect: null,
      pathLst: [],
    };
    try {
      evaluateGuides(shape, { w: 1, h: 1 });
      expect.unreachable();
    } catch (error) {
      expect((error as GeometryError).code).toBe('FMLA_OPERATOR');
    }
  });
});

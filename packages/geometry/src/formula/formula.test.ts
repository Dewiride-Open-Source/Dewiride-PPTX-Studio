import { describe, expect, it } from 'vitest';
import { GeometryError } from '../errors.js';
import { evaluateGuides } from './evaluate.js';
import { getPreset } from '../presets/index.js';
import {
  ANGLE_UNITS_PER_DEGREE,
  FMLA_ARITY,
  FULL_CIRCLE,
  angleToRadians,
  applyOperator,
  isFmlaOperator,
  radiansToAngle,
} from './formula.js';

/**
 * The sub-phase's stated verification is "unit tests per operator", and that is
 * what the first block below is. It is also, on its own, close to circular: a
 * test asserting that the multiply-divide operator multiplies and divides is
 * checking this file against the same understanding that wrote it.
 *
 * So the interesting assertions here are the ones a wrong reading would fail:
 * operand ORDER, angle UNITS, the comparison threshold, and the two boundary
 * behaviours that the 187 presets happen not to exercise. The genuinely
 * independent checks - geometry that has to come out right for a triangle to be
 * a triangle - are in `evaluate.test.ts`, where there is a shape to check
 * against.
 */

const DEG = ANGLE_UNITS_PER_DEGREE;

describe('the operator set', () => {
  it('is exactly seventeen operators', () => {
    expect(Object.keys(FMLA_ARITY)).toHaveLength(17);
  });

  it('agrees with isFmlaOperator, in both directions', () => {
    for (const op of Object.keys(FMLA_ARITY)) expect(isFmlaOperator(op)).toBe(true);
    for (const notAnOp of ['', 'mul', 'MOD', '*', '/', '+', 'atan2', 'if']) {
      expect(isFmlaOperator(notAnOp)).toBe(false);
    }
  });

  it('refuses an operator it does not know, with a typed error', () => {
    expect(() => applyOperator('atan2', [1, 1])).toThrow(GeometryError);
    try {
      applyOperator('atan2', [1, 1]);
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(GeometryError);
      expect((error as GeometryError).code).toBe('FMLA_OPERATOR');
    }
  });

  it('names the shape and the guide in the message when it is given a site', () => {
    try {
      applyOperator('nope', [], { preset: 'blockArc', guide: 'stAng' });
      expect.unreachable();
    } catch (error) {
      expect((error as GeometryError).message).toContain('blockArc.stAng');
      expect((error as GeometryError).preset).toBe('blockArc');
    }
  });
});

describe('arithmetic', () => {
  it('multiplies then divides', () => {
    expect(applyOperator('*/', [10, 20, 5])).toBe(40);
  });

  it('adds then subtracts, third operand last', () => {
    // `blockArc.iswAng = +- 0 0 swAng` is the negation idiom and only works if
    // the third operand is the one subtracted.
    expect(applyOperator('+-', [10, 20, 5])).toBe(25);
    expect(applyOperator('+-', [0, 0, 4500])).toBe(-4500);
  });

  it('adds then divides', () => {
    expect(applyOperator('+/', [10, 20, 5])).toBe(6);
  });

  it('takes an absolute value, a square root, and a value', () => {
    expect(applyOperator('abs', [-5])).toBe(5);
    expect(applyOperator('abs', [5])).toBe(5);
    expect(applyOperator('sqrt', [16])).toBe(4);
    expect(applyOperator('val', [42])).toBe(42);
  });

  it('takes the square root of the absolute value, which is Office and not ECMA', () => {
    // MS-OI29500 records this as an Office deviation from the standard. Apache
    // POI implements the standard and returns NaN, which is contagious: every
    // guide downstream becomes NaN and the shape draws nothing at all rather
    // than drawing something wrong. Following the standard here is the worse
    // failure, so this follows Office.
    expect(applyOperator('sqrt', [-16])).toBe(4);
    expect(applyOperator('sqrt', [-16])).not.toBeNaN();
    expect(applyOperator('sqrt', [0])).toBe(0);
  });

  it('takes a maximum and a minimum', () => {
    expect(applyOperator('max', [3, 7])).toBe(7);
    expect(applyOperator('min', [3, 7])).toBe(3);
  });
});

describe('mod is the vector modulus, not a remainder', () => {
  it('is the length of a three-dimensional vector', () => {
    expect(applyOperator('mod', [3, 4, 0])).toBe(5);
    expect(applyOperator('mod', [1, 2, 2])).toBe(3);
    expect(applyOperator('mod', [2, 3, 6])).toBe(7);
  });

  it('is not a remainder, which a three-operand form could not be anyway', () => {
    // If this were `x % y` the answer would be 1. Every one of the 48 uses in
    // the corpus passes three operands, which a binary remainder cannot take -
    // that is what settles it, rather than any reading of the name.
    expect(applyOperator('mod', [7, 3, 0])).toBeCloseTo(Math.hypot(7, 3), 10);
    expect(applyOperator('mod', [7, 3, 0])).not.toBe(1);
  });
});

describe('pin clamps its middle operand', () => {
  it('passes a value already inside the range', () => {
    expect(applyOperator('pin', [0, 5, 10])).toBe(5);
  });

  it('clamps below and above', () => {
    expect(applyOperator('pin', [0, -5, 10])).toBe(0);
    expect(applyOperator('pin', [0, 15, 10])).toBe(10);
  });

  it('is inclusive at both bounds', () => {
    expect(applyOperator('pin', [0, 0, 10])).toBe(0);
    expect(applyOperator('pin', [0, 10, 10])).toBe(10);
  });

  it('is an ordered chain and not max(x, min(y, z)), which differ on inverted bounds', () => {
    // The canonical distinguishing case. The chain gives 3; the min/max
    // composition gives 10. Apache POI rewrote this operator into the min/max
    // form in 5.2.5, so the divergence exists in a shipped implementation.
    expect(applyOperator('pin', [10, 20, 3])).toBe(3);
    expect(applyOperator('pin', [10, 5, 0])).toBe(10);

    // Pinned so the behaviour is a decision rather than an accident of how the
    // comparisons happened to be ordered.
  });

  it('and a committed preset reaches the divergence at a legal size', () => {
    // Found while building 2.5, which had to know whether a handle's search
    // range could invert. `mathDivide.a3` is `pin 1000 adj3 maxAdj3`, and
    // `maxAdj3` is `min(ma3h, ma3w)` with `ma3w = 36745*w/h`. On a shape one
    // unit wide and a thousand tall that is 36.745, well under the floor of
    // 1000, so the bounds cross. This is the only one: 1568 evaluations of the
    // 196 `pin` formulas at eight aspect ratios turn up exactly one.
    const shape = getPreset('mathDivide');
    if (shape === undefined) throw new Error('no mathDivide');
    const thin = evaluateGuides(shape, { w: 1, h: 1000 });
    expect(thin.get('maxAdj3')).toBeCloseTo(36.745, 6);
    // The chain takes the ceiling, because it is applied last.
    expect(thin.get('a3')).toBeCloseTo(36.745, 6);
    // Apache POI's min/max rewrite would take the floor instead.
    expect(Math.max(1000, Math.min(11760, 36.745))).toBe(1000);
  });
});

describe('the ternary tests against zero', () => {
  it('takes the second operand when the first is positive', () => {
    expect(applyOperator('?:', [1, 2, 3])).toBe(2);
  });

  it('takes the third when the first is negative', () => {
    expect(applyOperator('?:', [-1, 2, 3])).toBe(3);
  });

  it('takes the third at exactly zero, because the comparison is strict', () => {
    // The preset corpus cannot separate `> 0` from `>= 0`: every use is an
    // angle normalisation of the form `?: a a (a + 21600000)`, and at a == 0
    // both branches name the same angle. So this boundary is pinned here
    // rather than discovered later.
    expect(applyOperator('?:', [0, 2, 3])).toBe(3);
  });
});

describe('angles are in 60000ths of a degree', () => {
  it('converts to radians and back', () => {
    expect(angleToRadians(90 * DEG)).toBeCloseTo(Math.PI / 2, 12);
    expect(radiansToAngle(Math.PI)).toBeCloseTo(180 * DEG, 6);
    expect(radiansToAngle(angleToRadians(1234567))).toBeCloseTo(1234567, 6);
  });

  it('puts a full circle at 21600000', () => {
    expect(FULL_CIRCLE).toBe(360 * DEG);
    expect(FULL_CIRCLE).toBe(21600000);
  });

  it('scales the first operand by the sine or cosine of the second', () => {
    // The multiply is real. `blockArc.wt1 = sin wd2 stAng` is half the width
    // times a sine, and reading `sin` as unary would make every arc collapse.
    expect(applyOperator('sin', [100, 90 * DEG])).toBeCloseTo(100, 9);
    expect(applyOperator('sin', [100, 0])).toBeCloseTo(0, 9);
    expect(applyOperator('cos', [100, 0])).toBeCloseTo(100, 9);
    expect(applyOperator('cos', [100, 180 * DEG])).toBeCloseTo(-100, 9);
    expect(applyOperator('tan', [100, 45 * DEG])).toBeCloseTo(100, 9);
  });
});

describe('at2 is arctan(y / x) and is not normalised', () => {
  it('takes the denominator first', () => {
    // `circularArrow.u13 = at2 1 u12` reads as arctan of u12, which requires
    // the first operand to be the denominator.
    expect(applyOperator('at2', [1, 1])).toBeCloseTo(45 * DEG, 6);
    expect(applyOperator('at2', [1, 0])).toBe(0);
    expect(applyOperator('at2', [0, 1])).toBeCloseTo(90 * DEG, 6);
  });

  it('returns a negative angle rather than wrapping it into [0, 360)', () => {
    // Three guides in `circularArrow` follow an at2 with
    // `?: a a (a + 21600000)`. That idiom is dead code if this operator
    // normalises, and the sweep direction of every circular arrow depends on
    // it staying alive.
    expect(applyOperator('at2', [1, -1])).toBeCloseTo(-45 * DEG, 6);
    expect(applyOperator('at2', [1, -1])).toBeLessThan(0);
  });

  it('is defined at the origin, including when a negative zero arrives', () => {
    expect(applyOperator('at2', [0, 0])).toBe(0);

    // Math.atan2(0, -0) is pi, not zero. A guide chain reaches a negative zero
    // easily - `*/ 0 1 -1` is enough - and without the explicit guard a
    // zero-length vector would come back as a half turn.
    expect(applyOperator('at2', [-0, 0])).toBe(0);
    expect(applyOperator('at2', [-0, -0])).toBe(0);
    expect(Math.atan2(0, -0)).toBeCloseTo(Math.PI, 12);
  });
});

describe('cat2 and sat2 are the ellipse unskew', () => {
  it('project onto the cosine and sine of arctan(z / y)', () => {
    expect(applyOperator('cat2', [10, 1, 0])).toBeCloseTo(10, 9);
    expect(applyOperator('sat2', [10, 1, 0])).toBeCloseTo(0, 9);
    expect(applyOperator('cat2', [10, 0, 1])).toBeCloseTo(0, 9);
    expect(applyOperator('sat2', [10, 0, 1])).toBeCloseTo(10, 9);
  });

  it('take the third operand as the numerator of the arctangent', () => {
    // Swapping y and z is the mistake that moves the start of a blockArc from
    // 3 o'clock to 6 o'clock. `evaluate.test.ts` checks that consequence on the
    // real shape; this checks the operator in isolation.
    expect(applyOperator('cat2', [1, 3, 4])).toBeCloseTo(Math.cos(Math.atan2(4, 3)), 12);
    expect(applyOperator('cat2', [1, 3, 4])).toBeCloseTo(0.6, 12);
    expect(applyOperator('sat2', [1, 3, 4])).toBeCloseTo(0.8, 12);
  });

  it('keep a point on the unit circle for any pair', () => {
    for (const [y, z] of [
      [1, 0],
      [3, 4],
      [-2, 5],
      [0.001, 1000],
    ] as const) {
      const c = applyOperator('cat2', [1, y, z]);
      const s = applyOperator('sat2', [1, y, z]);
      expect(c * c + s * s).toBeCloseTo(1, 12);
    }
  });
});

describe('arity', () => {
  it('refuses too few and too many operands', () => {
    for (const [op, args] of [
      ['val', [1, 2]],
      ['abs', []],
      ['*/', [1, 2]],
      ['sin', [1]],
      ['mod', [1, 2]],
      ['pin', [1, 2, 3, 4]],
    ] as const) {
      try {
        applyOperator(op, args);
        expect.unreachable(`${op} should have refused ${String(args.length)} operand(s)`);
      } catch (error) {
        expect((error as GeometryError).code).toBe('FMLA_ARITY');
      }
    }
  });

  it('accepts a fourth operand on +- only when it is exactly zero', () => {
    // The eight over-long formulas in POI's file are all of the form
    // `+- xH 0 dxB 0`. A trailing zero cannot change the answer under any
    // reading of what a fourth operand would mean, so it is accepted; a
    // non-zero one is a case where the readings genuinely diverge, and the
    // honest response there is to stop.
    expect(applyOperator('+-', [10, 0, 4, 0])).toBe(6);
    try {
      applyOperator('+-', [10, 0, 4, 1]);
      expect.unreachable();
    } catch (error) {
      expect((error as GeometryError).code).toBe('FMLA_ARITY');
    }
  });
});

describe('non-finite results are returned, not thrown', () => {
  it('divides by zero the way IEEE does', () => {
    expect(applyOperator('*/', [10, 20, 0])).toBe(Infinity);
    expect(applyOperator('*/', [-10, 20, 0])).toBe(-Infinity);
    expect(applyOperator('+/', [10, 20, 0])).toBe(Infinity);
    expect(applyOperator('*/', [0, 20, 0])).toBeNaN();
  });

  it('has no NaN path through sqrt at all, because Office takes the absolute value first', () => {
    expect(applyOperator('sqrt', [-1])).toBe(1);
  });
});

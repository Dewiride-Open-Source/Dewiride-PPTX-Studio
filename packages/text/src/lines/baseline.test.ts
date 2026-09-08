import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/text-rendering.json' with { type: 'json' };
import { TextError } from '../errors.js';

import {
  FACE_BOX_PX,
  baselineDrop,
  baselineShare,
  createFaceBoxProbe,
  type FaceBox,
} from './baseline.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface BaselineRow {
  readonly probe: string;
  readonly face: string;
  readonly size: number;
  readonly sizes: readonly number[];
  readonly lineHeight: number;
  readonly drop: number;
  readonly share: number;
}

const rows = fixture.powerpoint.baseline.rows as readonly BaselineRow[];
const faces = fixture.browser.faces as Readonly<
  Record<string, { ascent: number; descent: number; baselineShare: number }>
>;

/** One EMF logical unit is a six-hundredth of an inch, which is 0.12pt. */
const UNIT_PT = 72 / 600;

function boxOf(face: string): FaceBox {
  const box = faces[face];
  if (box === undefined) throw new Error(`the fixture has no browser metrics for ${face}`);
  return { ascent: box.ascent, descent: box.descent, ideographic: 0 };
}

/** The score of a candidate over the fixture's rows, for a test name. */
function scoreOf(predict: (row: BaselineRow) => number): string {
  const hits = rows.filter((row) => Math.abs(predict(row) - row.drop) <= UNIT_PT).length;
  return `${String(hits)}/${String(rows.length)}`;
}

describe('the baseline inside a line box', () => {
  it(`is the face's own share of the box, ${scoreOf((row) => baselineDrop(row.lineHeight, boxOf(row.face)))}`, () => {
    expect(rows.length).toBeGreaterThan(30);
    for (const row of rows) {
      expect(
        Math.abs(baselineDrop(row.lineHeight, boxOf(row.face)) - row.drop),
        `${row.probe}: ${String(row.drop)}pt measured`,
      ).toBeLessThanOrEqual(UNIT_PT);
    }
  });

  it(`is not the CSS half-leading model, ${scoreOf((row) => {
    const box = boxOf(row.face);
    return (row.lineHeight - (box.ascent + box.descent) * row.size) / 2 + box.ascent * row.size;
  })}`, () => {
    // The reading every browser implements, and the one a renderer written by
    // hand uses. It is wrong by more than a fifth of a line on Courier New.
    const misses = rows.filter((row) => {
      const box = boxOf(row.face);
      const css =
        (row.lineHeight - (box.ascent + box.descent) * row.size) / 2 + box.ascent * row.size;
      return Math.abs(css - row.drop) > UNIT_PT;
    });
    expect(misses.length).toBeGreaterThan(20);
  });

  it('is not the ascent measured down from the top of the box', () => {
    const misses = rows.filter(
      (row) => Math.abs(boxOf(row.face).ascent * row.size - row.drop) > UNIT_PT,
    );
    expect(misses.length).toBeGreaterThan(30);
  });

  it('is not a face-independent fraction of the box', () => {
    const misses = rows.filter((row) => Math.abs(row.lineHeight * 0.81 - row.drop) > UNIT_PT);
    expect(misses.length).toBeGreaterThan(10);
  });

  it('agrees with the share the fixture recorded per face', () => {
    for (const [face, box] of Object.entries(faces)) {
      expect(baselineShare(boxOf(face)), face).toBeCloseTo(box.baselineShare, 6);
    }
  });
});

describe('baselineShare', () => {
  it('refuses a box that is not one', () => {
    expect(() => baselineShare({ ascent: 0, descent: 0, ideographic: 0 })).toThrow(TextError);
    expect(() => baselineShare({ ascent: -1, descent: 1, ideographic: 0 })).toThrow(TextError);
    expect(() => baselineShare({ ascent: Number.NaN, descent: 1, ideographic: 0 })).toThrow(
      TextError,
    );
  });

  it('is between nothing and everything', () => {
    for (const face of Object.keys(faces)) {
      expect(baselineShare(boxOf(face))).toBeGreaterThan(0.7);
      expect(baselineShare(boxOf(face))).toBeLessThan(0.9);
    }
  });
});

describe('baselineDrop', () => {
  it('refuses a line box that is not a height', () => {
    const box = boxOf('Arial');
    expect(() => baselineDrop(-1, box)).toThrow(TextError);
    expect(() => baselineDrop(Number.POSITIVE_INFINITY, box)).toThrow(TextError);
  });

  it('is proportional to the line box', () => {
    const box = boxOf('Arial');
    expect(baselineDrop(24, box)).toBeCloseTo(2 * baselineDrop(12, box), 10);
  });
});

describe('createFaceBoxProbe', () => {
  it('measures at a size where the browser has stopped quantising', () => {
    // 100px reads Arial's descent share as 0.2250 where it is 0.2277, which is
    // half a point out at 54pt.
    expect(FACE_BOX_PX).toBeGreaterThanOrEqual(1000);
  });

  it('reads a face box this machine actually has', () => {
    const probe = createFaceBoxProbe();
    const box = probe.box('Arial');
    expect(box.ascent).toBeGreaterThan(0.5);
    expect(box.descent).toBeGreaterThan(0.05);
  });

  it('reads the ideographic baseline as a drop below the alphabetic one', () => {
    // Canvas reports it as a negative offset. A probe that kept that sign would
    // read every face as zero and stand each upright glyph a whole em out.
    const box = createFaceBoxProbe().box('Arial');
    expect(box.ideographic).toBeGreaterThan(0.05);
    expect(box.ideographic).toBeLessThan(0.5);
  });

  it('gives the same box twice, from the cache', () => {
    const probe = createFaceBoxProbe();
    expect(probe.box('Georgia')).toEqual(probe.box('Georgia'));
  });

  it('refuses a family that cannot be written into a CSS shorthand', () => {
    expect(() => createFaceBoxProbe().box('has "quotes"')).toThrow(TextError);
  });

  it('measures Arial close to what PowerPoint drew', () => {
    // The whole rule in one assertion: the browser's own box, put through the
    // share, lands within a pixel of where PowerPoint put a 54pt baseline.
    const measured = rows.find((row) => row.probe === 'base-Arial-54');
    expect(measured).toBeDefined();
    const probe = createFaceBoxProbe();
    const drop = baselineDrop(measured?.lineHeight ?? 0, probe.box('Arial'));
    expect(Math.abs(drop - (measured?.drop ?? 0))).toBeLessThanOrEqual(0.5);
  });
});

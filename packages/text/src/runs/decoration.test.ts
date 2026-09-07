import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/text-rendering.json' with { type: 'json' };
import { TextError } from '../errors.js';

import {
  APPROXIMATE_FACE_RULES,
  MEASURED_FACE_RULES,
  decoratedStretches,
  faceRules,
  hasFaceRules,
  strikeRules,
  underlineRules,
  type Strike,
  type Underline,
} from './decoration.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface Rule {
  readonly offset: number;
  readonly thickness: number;
  readonly left: number;
  readonly right: number;
  readonly offsetEm: number;
  readonly thicknessEm: number;
}

interface DecRow {
  readonly probe: string;
  readonly face: string;
  readonly size: number;
  readonly glyphRight: number | null;
  readonly rules: readonly Rule[];
}

interface FaceRow {
  readonly probe: string;
  readonly kind: 'u' | 's';
  readonly face: string;
  readonly size: number;
  readonly offsetEm: number;
  readonly thicknessEm: number;
  readonly chromiumOffsetEm: number;
  readonly chromiumThicknessEm: number;
}

const decoration = fixture.powerpoint.decoration as readonly DecRow[];
const faceRows = fixture.powerpoint.decorationFace as readonly FaceRow[];

/** Half a point, which at 1920 pixels over a 960pt slide is one pixel. */
const PIXEL_PT = 0.5;

function rowFor(probe: string): DecRow {
  const row = decoration.find((entry) => entry.probe === probe);
  if (row === undefined) throw new Error(`the fixture has no probe ${probe}`);
  return row;
}

describe('the measured table', () => {
  it('holds every face the fixture measured, and the same numbers', () => {
    for (const row of faceRows) {
      const rules = MEASURED_FACE_RULES[row.face];
      expect(rules, row.face).toBeDefined();
      const measured = row.kind === 'u' ? rules?.underline : rules?.strike;
      expect(measured?.offset, row.probe).toBeCloseTo(row.offsetEm, 6);
      expect(measured?.thickness, row.probe).toBeCloseTo(row.thicknessEm, 6);
    }
  });

  it('does not agree with what the browser draws for any face', () => {
    // The finding that decides the implementation: `text-decoration` is 0 of 16
    // on the offset and 0 of 16 on the thickness, so both rules are drawn.
    for (const row of faceRows) {
      const offsets = Math.abs(row.offsetEm - row.chromiumOffsetEm) * row.size;
      const thicknesses = Math.abs(row.thicknessEm - row.chromiumThicknessEm) * row.size;
      expect(Math.max(offsets, thicknesses), row.probe).toBeGreaterThan(PIXEL_PT);
    }
  });

  it('cannot be replaced by one offset for every face', () => {
    const arial = MEASURED_FACE_RULES['Arial'];
    const differ = faceRows.filter((row) => {
      const mine = row.kind === 'u' ? arial?.underline : arial?.strike;
      return Math.abs((mine?.offset ?? 0) - row.offsetEm) * row.size > PIXEL_PT;
    });
    expect(differ.length).toBeGreaterThanOrEqual(8);
  });
});

describe('faceRules', () => {
  it('answers for a measured face', () => {
    expect(faceRules('Georgia').underline.offset).toBeCloseTo(0.086167, 6);
    expect(hasFaceRules('Georgia')).toBe(true);
  });

  it('refuses a face nobody measured, rather than guessing', () => {
    expect(hasFaceRules('Wingdings')).toBe(false);
    expect(() => faceRules('Wingdings')).toThrow(TextError);
  });

  it('offers an approximation under a name that says what it is', () => {
    expect(APPROXIMATE_FACE_RULES).toEqual(MEASURED_FACE_RULES['Arial']);
  });
});

describe('underlineRules', () => {
  const rules = faceRules('Arial');

  it('draws nothing for none and for an absent value', () => {
    expect(underlineRules(undefined, 40, rules)).toEqual([]);
    expect(underlineRules('none', 40, rules)).toEqual([]);
  });

  it.each([18, 40, 66])('puts a single rule where PowerPoint did at %ipt', (size) => {
    const measured = rowFor(`dec-u-sng-${String(size)}`);
    const drawn = underlineRules('sng', measured.size, rules);
    const first = drawn[0];
    const seen = measured.rules[0];
    expect(drawn.length).toBe(1);
    expect(Math.abs((first?.top ?? 0) - (seen?.offset ?? 0))).toBeLessThanOrEqual(PIXEL_PT);
    expect(Math.abs((first?.thickness ?? 0) - (seen?.thickness ?? 0))).toBeLessThanOrEqual(
      PIXEL_PT,
    );
  });

  it.each([18, 40, 66])('puts a double rule where PowerPoint did at %ipt', (size) => {
    const measured = rowFor(`dec-u-dbl-${String(size)}`);
    const drawn = underlineRules('dbl', measured.size, rules);
    expect(drawn.length).toBe(2);
    measured.rules.forEach((seen, index) => {
      expect(Math.abs((drawn[index]?.top ?? 0) - seen.offset), measured.probe).toBeLessThanOrEqual(
        PIXEL_PT,
      );
      expect(
        Math.abs((drawn[index]?.thickness ?? 0) - seen.thickness),
        `${measured.probe} thickness`,
      ).toBeLessThanOrEqual(PIXEL_PT);
    });
  });

  it.each([18, 40, 66])('puts a heavy rule where PowerPoint did at %ipt', (size) => {
    const measured = rowFor(`dec-u-heavy-${String(size)}`);
    const drawn = underlineRules('heavy', measured.size, rules);
    const first = drawn[0];
    const seen = measured.rules[0];
    expect(drawn.length).toBe(1);
    expect(Math.abs((first?.top ?? 0) - (seen?.offset ?? 0))).toBeLessThanOrEqual(PIXEL_PT);
    expect(Math.abs((first?.thickness ?? 0) - (seen?.thickness ?? 0))).toBeLessThanOrEqual(
      PIXEL_PT,
    );
  });

  it('is not the same rule for every value', () => {
    const single = underlineRules('sng', 40, rules)[0];
    const heavy = underlineRules('heavy', 40, rules)[0];
    expect(heavy?.thickness).toBeGreaterThan(single?.thickness ?? 0);
    expect(heavy?.top).toBeLessThan(single?.top ?? 0);
  });

  it('scales with the size rather than being a fixed length', () => {
    const small = underlineRules('sng', 20, rules)[0];
    const large = underlineRules('sng', 40, rules)[0];
    expect(large?.top).toBeCloseTo(2 * (small?.top ?? 0), 10);
    expect(large?.thickness).toBeCloseTo(2 * (small?.thickness ?? 0), 10);
  });

  it('carries the pattern each value asks for', () => {
    expect(underlineRules('dotted', 40, rules)[0]?.pattern).toBe('dotted');
    expect(underlineRules('dash', 40, rules)[0]?.pattern).toBe('dashed');
    expect(underlineRules('wavy', 40, rules)[0]?.wavy).toBe(true);
    expect(underlineRules('sng', 40, rules)[0]?.pattern).toBe('solid');
  });

  it('refuses a size that is not one', () => {
    expect(() => underlineRules('sng', 0, rules)).toThrow(TextError);
    expect(() => underlineRules('sng', Number.NaN, rules)).toThrow(TextError);
  });
});

describe('strikeRules', () => {
  const rules = faceRules('Arial');

  it('draws nothing for noStrike and for an absent value', () => {
    expect(strikeRules(undefined, 40, rules)).toEqual([]);
    expect(strikeRules('noStrike', 40, rules)).toEqual([]);
  });

  it.each([18, 40, 66])('puts a single strike where PowerPoint did at %ipt', (size) => {
    const measured = rowFor(`dec-s-sngStrike-${String(size)}`);
    const drawn = strikeRules('sngStrike', measured.size, rules);
    expect(Math.abs((drawn[0]?.top ?? 0) - (measured.rules[0]?.offset ?? 0))).toBeLessThanOrEqual(
      PIXEL_PT,
    );
  });

  it.each([18, 40, 66])('puts a double strike where PowerPoint did at %ipt', (size) => {
    const measured = rowFor(`dec-s-dblStrike-${String(size)}`);
    const drawn = strikeRules('dblStrike', measured.size, rules);
    expect(drawn.length).toBe(2);
    measured.rules.forEach((seen, index) => {
      expect(Math.abs((drawn[index]?.top ?? 0) - seen.offset), measured.probe).toBeLessThanOrEqual(
        PIXEL_PT,
      );
    });
  });

  it('sits above the baseline, unlike an underline', () => {
    expect(strikeRules('sngStrike', 40, rules)[0]?.top).toBeLessThan(0);
    expect(underlineRules('sng', 40, rules)[0]?.top).toBeGreaterThan(0);
  });

  it('refuses a value that is not an ST_TextStrikeType', () => {
    expect(() => strikeRules('sngstrike' as Strike, 40, rules)).toThrow(TextError);
  });
});

describe('decoratedStretches', () => {
  it('stops at the last glyph rather than under a trailing space', () => {
    expect(decoratedStretches('Alpha ', 'sng')).toEqual([[0, 5]]);
    expect(decoratedStretches('Alpha   ', 'sng')).toEqual([[0, 5]]);
  });

  it('skips every space for words, not only the trailing ones', () => {
    expect(decoratedStretches('Alpha bravo', 'words')).toEqual([
      [0, 5],
      [6, 11],
    ]);
  });

  it('is nothing for a run that draws no rule', () => {
    expect(decoratedStretches('Alpha', undefined)).toEqual([]);
    expect(decoratedStretches('Alpha', 'none')).toEqual([]);
    expect(decoratedStretches('   ', 'sng')).toEqual([]);
  });

  it('matches the width PowerPoint drew under a trailing space', () => {
    // The rule reached the last glyph and no further: the bare probe is the
    // same string with the spaces taken out of the file, and the two agree.
    const extent = fixture.powerpoint.decorationExtent as readonly {
      id: string;
      overhang: number;
      spaceAdvance: number;
    }[];
    for (const row of extent) {
      expect(Math.abs(row.spaceAdvance), row.id).toBeLessThanOrEqual(0.6);
    }
  });
});

describe('the underline styles the fixture could not read', () => {
  it('draws dotted, dashed and wavy at the single rule position', () => {
    // PowerPoint paints these as a pattern, so the bitmap holds no solid rail
    // to measure; the position is the single rule's, which the ADR records as
    // an assumption rather than a measurement.
    const rules = faceRules('Arial');
    const single = underlineRules('sng', 40, rules)[0];
    for (const style of ['dotted', 'dash', 'wavy'] as Underline[]) {
      expect(underlineRules(style, 40, rules)[0]?.top, style).toBeCloseTo(single?.top ?? 0, 10);
    }
  });
});

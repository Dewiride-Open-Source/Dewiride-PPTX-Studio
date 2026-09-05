/**
 * The line model, against what PowerPoint actually did.
 *
 * Every assertion here recomputes the number from `lnSpc` and the font size and
 * compares it against `readings` in `corpus/ground-truth/text-metrics.json` -
 * never against a constant copied out of `line-model.ts`. A suite written the
 * other way round passes forever and proves nothing; this one goes red the
 * moment the implementation changes its mind.
 *
 * The refutations are tests too. It is not enough that the right model fits:
 * the wrong models anyone would write first - the percentage applied to the
 * font size, the exact-point spacing scaled by 1.2, the value used without
 * quantisation - are asserted to be *wrong*, with the score the analysis gave
 * them. That is what makes the fixture worth having rather than a table of
 * numbers that happen to agree with the code.
 */

import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/text-metrics.json' with { type: 'json' };
import { TextError } from './errors.js';
import {
  NATURAL_LINE_FACTOR,
  blockHeight,
  lineAdvance,
  lineTop,
  naturalLineHeight,
  quantisePercent,
  type LineSpacing,
} from './line-model.js';

/**
 * `BoundHeight` comes back as float32 points quantised to a two-hundredth of a
 * point, so the same tolerance the analysis used applies here. The smallest
 * real difference any probe can produce is more than fifteen times this.
 */
const EPS = fixture.tolerancePt;

interface Reading {
  readonly font: string;
  readonly sz: number;
  readonly lnSpc: { readonly kind: string; readonly value: number } | null;
}

/** The fixture writes `lnSpc` as JSON; this is the same thing typed. */
function spacingOf(r: Reading): LineSpacing | undefined {
  if (r.lnSpc === null) return undefined;
  return {
    kind: r.lnSpc.kind === 'pct' ? 'percent' : 'points',
    value: r.lnSpc.value,
  };
}

const advances = fixture.readings.lineAdvance;
const lastLines = fixture.readings.lastLine;

describe('the line box', () => {
  it('is a font-independent 1.2 x the font size', () => {
    expect(NATURAL_LINE_FACTOR).toBe(fixture.lineModel.naturalLineFactor);
    expect(fixture.lineModel.fontIndependent).toBe(true);

    // The measurement that earns it: every face reporting the same advance at
    // the same size and spacing. Thirteen faces whose vertical metrics disagree
    // violently; a font-derived line height gives thirteen answers.
    const single = advances.filter((r) => r.lnSpc?.kind === 'pct' && r.lnSpc.value === 100000);
    const byGroup = new Map<string, number[]>();
    for (const r of single) {
      const key = `${String(r.sz)}|${String(r.lines)}`;
      byGroup.set(key, [...(byGroup.get(key) ?? []), r.advancePt]);
    }
    expect(byGroup.size).toBeGreaterThan(5);
    for (const [key, list] of byGroup) {
      expect(Math.max(...list) - Math.min(...list), key).toBeLessThan(EPS);
      expect(list.length, key).toBeGreaterThan(1);
    }
  });

  it('computes the natural height from the size alone', () => {
    for (const sz of [800, 1050, 1200, 1800, 3200, 5400, 9600]) {
      expect(naturalLineHeight(sz)).toBeCloseTo(1.2 * (sz / 100), 10);
    }
  });
});

describe('lineAdvance, against every multi-line probe', () => {
  it.each(advances.map((r) => [`${r.font} ${String(r.sz / 100)}pt`, r] as const))(
    'reproduces PowerPoint at %s',
    (_label: string, r: (typeof advances)[number]) => {
      expect(lineAdvance(r.sz, spacingOf(r))).toBeCloseTo(r.advancePt, 2);
    },
  );

  it('fits all of them, which is the claim the fixture makes', () => {
    const fits = advances.filter(
      (r) => Math.abs(lineAdvance(r.sz, spacingOf(r)) - r.advancePt) < EPS,
    ).length;
    expect(fits).toBe(advances.length);
    expect(advances.length).toBe(463);
  });
});

describe('the models this refutes', () => {
  /**
   * Score a candidate the way the analysis did. A number, not a boolean:
   * "the child map is right" is worth little, "the child map is right 463/463
   * and applying the percentage to the font size is right 66/463" says what the
   * bug would have looked like on a real slide.
   */
  function score(fn: (sz: number, ls: LineSpacing | undefined) => number): number {
    return advances.filter((r) => Math.abs(fn(r.sz, spacingOf(r)) - r.advancePt) < EPS).length;
  }

  it('refutes applying the percentage to the font size instead of the line box', () => {
    const wrong = (sz: number, ls: LineSpacing | undefined): number =>
      ls === undefined || ls.kind === 'percent'
        ? quantisePercent(ls?.value ?? 100000) * (sz / 100)
        : ls.value / 100;
    // 150% on an 18pt run is 32.4pt, not 27pt: a fifth of the height of every
    // line of every deck that sets line spacing.
    expect(lineAdvance(1800, { kind: 'percent', value: 150000 })).toBeCloseTo(32.4, 6);
    expect(wrong(1800, { kind: 'percent', value: 150000 })).toBeCloseTo(27, 6);
    expect(score(wrong)).toBeLessThan(advances.length);
    expect(score(wrong)).toBe(66);
  });

  it('refutes scaling exact-point spacing by 1.2 as well', () => {
    const wrong = (sz: number, ls: LineSpacing | undefined): number =>
      ls === undefined || ls.kind === 'percent'
        ? quantisePercent(ls?.value ?? 100000) * 1.2 * (sz / 100)
        : 1.2 * (ls.value / 100);
    expect(lineAdvance(1800, { kind: 'points', value: 3600 })).toBeCloseTo(36, 6);
    expect(wrong(1800, { kind: 'points', value: 3600 })).toBeCloseTo(43.2, 6);
    expect(score(wrong)).toBe(397);
  });

  it('refutes ignoring a:lnSpc entirely', () => {
    expect(score((sz) => 1.2 * (sz / 100))).toBe(244);
  });
});

describe('the percentage is quantised before it is used', () => {
  it('rounds half up to a whole percent', () => {
    expect(fixture.lineModel.percentQuantisation).toBe('round half up to a whole percent');
    expect(quantisePercent(100000)).toBeCloseTo(1, 10);
    expect(quantisePercent(100400)).toBeCloseTo(1, 10);
    expect(quantisePercent(100500)).toBeCloseTo(1.01, 10);
    expect(quantisePercent(99500)).toBeCloseTo(1, 10);
    // A value no PowerPoint dialog writes but other producers do: "1.07 lines".
    expect(quantisePercent(106667)).toBeCloseTo(1.07, 10);
  });

  it('is visible in the readings: 100.4% and 100.5% lay out differently', () => {
    const at = (value: number, sz: number): number | undefined =>
      advances.find(
        (r) =>
          r.font === 'Arial' && r.sz === sz && r.lnSpc?.kind === 'pct' && r.lnSpc.value === value,
      )?.advancePt;
    for (const sz of [1200, 1800, 3200]) {
      const hundred = at(100000, sz);
      const justUnder = at(100400, sz);
      const justOver = at(100500, sz);
      const onePercentUp = at(101000, sz);
      expect(hundred, String(sz)).toBeDefined();
      expect(justUnder).toBeCloseTo(hundred ?? 0, 2);
      expect(justOver).toBeCloseTo(onePercentUp ?? 0, 2);
      expect(Math.abs((justOver ?? 0) - (hundred ?? 0))).toBeGreaterThan(EPS);
    }
  });

  it('refutes ceiling and floor with the scores the analysis gave them', () => {
    const refuted = new Map(
      fixture.lineModel.refuted.percentQuantisation.map((r) => [r.model, r.fits]),
    );
    expect(refuted.get('ceil to a whole percent')).toBe(365);
    expect(refuted.get('floor to a whole percent')).toBe(359);
    expect(refuted.get('exact: val / 100000')).toBe(353);
    for (const fits of refuted.values()) expect(fits).toBeLessThan(371);
  });

  it('rejects a percentage the schema cannot hold', () => {
    expect(() => quantisePercent(-1)).toThrow(TextError);
    expect(() => quantisePercent(Number.NaN)).toThrow(TextError);
    expect(() => lineAdvance(0)).toThrow(TextError);
    expect(() => lineAdvance(-1200)).toThrow(TextError);
  });
});

describe('how a block is composed', () => {
  it('starts on the frame top, with no half-leading above the first line', () => {
    expect(fixture.lineModel.firstLineTopOffsetPt).toBe(0);
    expect(lineTop(0, 21.6)).toBe(0);
    expect(lineTop(2, 21.6)).toBeCloseTo(43.2, 10);
  });

  it('is (n - 1) x advance + the last line, on every multi-line probe', () => {
    let checked = 0;
    for (const r of advances) {
      const one = lastLines.find(
        (q) =>
          q.font === r.font && q.sz === r.sz && JSON.stringify(q.lnSpc) === JSON.stringify(r.lnSpc),
      );
      if (one === undefined) continue;
      checked += 1;
      const predicted = blockHeight(r.lines, lineAdvance(r.sz, spacingOf(r)), one.heightPt);
      expect(predicted, `${r.font} ${String(r.sz)} x${String(r.lines)}`).toBeCloseTo(r.heightPt, 2);
    }
    expect(checked).toBeGreaterThan(100);
  });

  it('refutes n x advance, which is what every line but the last measures', () => {
    // At 100% the two agree, so a suite that only tested single spacing would
    // never notice. They part company the moment anything sets line spacing.
    const wide = advances.filter(
      (r) => r.lnSpc?.kind === 'pct' && r.lnSpc.value === 200000 && r.lines === 3,
    );
    expect(wide.length).toBeGreaterThan(0);
    for (const r of wide) {
      expect(r.lines * r.advancePt).toBeGreaterThan(r.heightPt + EPS);
    }
  });

  it('rejects a line count that is not a count', () => {
    expect(() => blockHeight(-1, 20, 20)).toThrow(TextError);
    expect(() => blockHeight(1.5, 20, 20)).toThrow(TextError);
    expect(() => lineTop(-1, 20)).toThrow(TextError);
    expect(blockHeight(0, 20, 20)).toBe(0);
  });
});

describe('the last line, which is data and not law', () => {
  it('records the best model, its score, and the face it misses', () => {
    // Committed as a measurement rather than a rule, on purpose. The suite pins
    // what is known so that 3.6 starts from the number rather than re-deriving
    // it, and so that a later claim to have solved it has something to beat.
    expect(fixture.lastLine.fits).toBe(1548);
    expect(fixture.lastLine.of).toBe(1579);
    expect(fixture.lastLine.missesOnlyOn).toEqual(['Courier New']);
    expect(fixture.lastLine.misses).toHaveLength(31);
  });

  it('beats every simpler reading of it', () => {
    const scored = new Map(fixture.lastLine.scored.map((s) => [s.model, s.fits]));
    const best = fixture.lastLine.fits;
    // Exactly one candidate reaches the best score, and every other is strictly
    // worse. Identified by its score rather than by its wording: the summary and
    // the score table describe the same model in different prose, and a test
    // that matched on the string would silently compare nothing if either
    // sentence were reworded.
    expect([...scored.values()].filter((fits) => fits === best)).toHaveLength(1);
    for (const [model, fits] of scored) {
      if (fits === best) continue;
      expect(fits, model).toBeLessThan(best);
    }
    expect(scored.get('the advance itself (what every line above the last measures)')).toBe(897);
    expect(scored.get('max(advance, 0.75 x advance + c)')).toBe(1027);
  });

  it('needs a per-face term, which is why it is not in the code yet', () => {
    const descents = fixture.lastLine.descentPt;
    expect(descents.length).toBeGreaterThan(80);
    // If it were font-independent, like the advance is, one number would do.
    const ratios = descents.filter((d) => d.sz === 1800).map((d) => d.descentPt / (d.sz / 100));
    expect(Math.max(...ratios) - Math.min(...ratios)).toBeGreaterThan(0.04);
  });
});

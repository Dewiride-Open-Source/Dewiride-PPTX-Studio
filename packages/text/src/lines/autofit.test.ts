/**
 * Autofit, re-derived from the T4 fixture.
 *
 * Every assertion below reads `corpus/ground-truth/autofit.json` and works out
 * what the answer should be from the measurement. None of them was written by
 * reading `autofit.ts`: a test that re-derives the rule fails when the code is
 * wrong, and a test written from the code passes forever and proves nothing.
 *
 * There are two replays, because the experiment used two instruments.
 *
 * - **The rungs.** Every box height PowerPoint was swept through, with the
 *   `@fontScale` and `@lnSpcReduction` it wrote for it. 2699 of them, and the
 *   implementation has to choose the same rung on every one.
 * - **The blocks.** Every text body PowerPoint laid out without autofitting,
 *   with the advance between its baselines and the height of the whole block.
 *   581 of them, and the implementation has to reproduce the height.
 *
 * Neither replay alone settles the rule - three different readings of what the
 * last line measures all fit the rungs perfectly and only one fits the blocks -
 * so both are here and both have to pass.
 */

import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/autofit.json' with { type: 'json' };

import {
  APPROXIMATE_FACE_METRICS,
  AUTOFIT_LADDER,
  MEASURED_FACE_METRICS,
  SHAPE_AUTOFIT_SLACK,
  autofitAdvance,
  effectiveFontSize,
  faceLineMetrics,
  fitAutofit,
  hasFaceLineMetrics,
  lastLineHeight,
  requiredShapeHeight,
  storedScale,
  textBodyHeight,
  type AutofitScale,
  type FaceLineMetrics,
  type ParagraphBox,
} from './autofit.js';
import { TextError } from '../errors.js';
import { NATURAL_LINE_FACTOR, type LineSpacing } from './line-model.js';

/* ------------------------------------------------------------------ shapes */

interface FixtureSpacing {
  readonly kind: string;
  readonly value: number;
}
interface FixtureInsets {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}
/** One height sweep: a configuration, and two parallel columns. */
interface FixtureSweep {
  readonly deck: string;
  readonly sz: number;
  readonly face: string;
  readonly lines: number;
  readonly lnSpc: FixtureSpacing | null;
  readonly spcBef: FixtureSpacing | null;
  readonly spcAft: FixtureSpacing | null;
  readonly spcFirstLastPara: boolean | null;
  readonly insets: FixtureInsets;
  readonly boxPt: readonly number[];
  /** An index into `fixture.ladder`. */
  readonly rung: readonly number[];
}
/** One text body PowerPoint laid out, measured rather than decided. */
interface FixtureBlock {
  readonly id: string;
  readonly deck: string;
  readonly face: string;
  readonly sz: number;
  readonly autofit: string;
  readonly fontScale?: number;
  readonly lnSpcReduction?: number;
  readonly lnSpc?: FixtureSpacing;
  readonly spcBef?: FixtureSpacing;
  readonly spcAft?: FixtureSpacing;
  readonly spcFirstLastPara?: boolean;
  readonly paragraphs: number;
  readonly lines: number;
  readonly advancePt?: number;
  readonly blockHeightPt: number;
}
interface FixtureSp {
  readonly id: string;
  readonly face: string;
  readonly sz: number;
  readonly lines: number;
  readonly tIns: number;
  readonly bIns: number;
  readonly boxPt: number;
  readonly emu: number;
}

const sweeps = fixture.sweeps as readonly FixtureSweep[];
const blocks = fixture.blocks as readonly FixtureBlock[];
const ladder = fixture.ladder as readonly AutofitScale[];
const spProbes = fixture.spAutoFit.probes as readonly FixtureSp[];
const perFace = fixture.lineBox.perFace as Readonly<Record<string, number>>;
const floorRatios = fixture.lineBox.floorRatio as Readonly<Record<string, number>>;

const EMU_PER_POINT = 12700;

function spacing(value: FixtureSpacing | null | undefined): LineSpacing | undefined {
  if (value === null || value === undefined) return undefined;
  if (value.kind !== 'percent' && value.kind !== 'points') {
    throw new Error(`the fixture states a spacing kind nobody wrote: ${value.kind}`);
  }
  return { kind: value.kind, value: value.value };
}

/**
 * The measured line advance, where the fixture carries one.
 *
 * A probe stating spcBef or spcAft has that spacing folded into every gap
 * between two line tops, so the fixture omits the number there rather than
 * publishing a sum of two rules under the name of one.
 */
function measuredAdvance(block: FixtureBlock): number {
  if (block.advancePt === undefined) {
    throw new Error(`${block.id} states paragraph spacing, so its gaps are not line advances`);
  }
  return block.advancePt;
}

function metricsFor(face: string): FaceLineMetrics {
  const metrics = MEASURED_FACE_METRICS[face];
  if (metrics === undefined)
    throw new Error(`the fixture measured ${face} and the package did not`);
  return metrics;
}

/* --------------------------------------------------------------- the ladder */

describe('the ladder', () => {
  it('is the rungs PowerPoint emitted, in the order it tried them', () => {
    expect(AUTOFIT_LADDER).toEqual(ladder);
  });

  it('has fifteen rungs and no scale between them', () => {
    expect(AUTOFIT_LADDER).toHaveLength(15);
    const scales = [...new Set(AUTOFIT_LADDER.map((r) => r.fontScale))];
    expect(scales).toEqual([
      100000, 92500, 85000, 77500, 70000, 62500, 55000, 47500, 40000, 32500, 25000,
    ]);
  });

  it('is ordered font scale descending, then reduction ascending', () => {
    for (let i = 1; i < AUTOFIT_LADDER.length; i += 1) {
      const previous = AUTOFIT_LADDER[i - 1];
      const current = AUTOFIT_LADDER[i];
      if (previous === undefined || current === undefined) throw new Error('unreachable');
      const ordered =
        previous.fontScale > current.fontScale ||
        (previous.fontScale === current.fontScale &&
          previous.lnSpcReduction < current.lnSpcReduction);
      expect(ordered, `rung ${String(i)}`).toBe(true);
    }
  });

  it('carries a reduction of 20% below 85%, and the smaller ones only above it', () => {
    // Read off the fixture rather than restated: the reductions a font scale
    // carries are exactly the ones PowerPoint was seen to emit for it.
    const byScale = new Map<number, number[]>();
    for (const rung of ladder) {
      const list = byScale.get(rung.fontScale) ?? [];
      list.push(rung.lnSpcReduction);
      byScale.set(rung.fontScale, list);
    }
    for (const [scale, reductions] of byScale) {
      if (scale < 85000) expect(reductions, String(scale)).toEqual([20000]);
      else expect(reductions.length, String(scale)).toBeGreaterThan(1);
    }
  });
});

/* ------------------------------------------------------- the two constants */

describe('the per-face line metrics', () => {
  it('are the numbers T4 fitted, for every face it fitted them for', () => {
    for (const [face, coefficient] of Object.entries(perFace)) {
      expect(metricsFor(face).lastLineCoefficient, face).toBe(coefficient);
    }
    for (const [face, ratio] of Object.entries(floorRatios)) {
      expect(metricsFor(face).floorRatio, face).toBe(ratio);
    }
    expect(Object.keys(MEASURED_FACE_METRICS).sort()).toEqual(Object.keys(perFace).sort());
  });

  it('put the floor at the 1.2 line box for five faces and not for Courier New', () => {
    const bracket = fixture.lineBox.floorBracket as Readonly<
      Record<string, { lower: number; upper: number }>
    >;
    for (const [face, { lower, upper }] of Object.entries(bracket)) {
      const ratio = metricsFor(face).floorRatio;
      expect(ratio, face).toBeGreaterThanOrEqual(lower);
      expect(ratio, face).toBeLessThan(upper);
      if (face === 'Courier New') expect(ratio).toBeGreaterThan(NATURAL_LINE_FACTOR);
      else expect(ratio).toBe(NATURAL_LINE_FACTOR);
    }
  });

  it('are not a browser metric: two faces share a descent and differ in the coefficient', () => {
    const comparison = fixture.lineBox.browserComparison as readonly {
      face: string;
      descentRatio: number;
      b: number;
    }[];
    const arial = comparison.find((r) => r.face === 'Arial');
    const verdana = comparison.find((r) => r.face === 'Verdana');
    if (arial === undefined || verdana === undefined) throw new Error('unreachable');
    expect(Math.abs(arial.descentRatio - verdana.descentRatio)).toBeLessThan(0.003);
    expect(Math.abs(arial.b - verdana.b)).toBeGreaterThan(0.01);
  });

  it('refuses a face nobody measured, rather than inventing one', () => {
    expect(hasFaceLineMetrics('Arial')).toBe(true);
    expect(hasFaceLineMetrics('Comic Sans MS')).toBe(false);
    expect(() => faceLineMetrics('Comic Sans MS')).toThrow(TextError);
    try {
      faceLineMetrics('Comic Sans MS');
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(TextError);
      expect((error as TextError).code).toBe('TEXT_FACE_METRICS');
      expect((error as TextError).token).toBe('Comic Sans MS');
    }
  });

  it('offers an approximation that has to be named to be used', () => {
    // Inside the measured spread, so a caller that reaches for it is wrong by
    // less than the difference between two real faces.
    const coefficients = Object.values(perFace);
    expect(APPROXIMATE_FACE_METRICS.lastLineCoefficient).toBeGreaterThanOrEqual(
      Math.min(...coefficients),
    );
    expect(APPROXIMATE_FACE_METRICS.lastLineCoefficient).toBeLessThanOrEqual(
      Math.max(...coefficients),
    );
    expect(APPROXIMATE_FACE_METRICS.floorRatio).toBe(NATURAL_LINE_FACTOR);
  });
});

/* ---------------------------------------------------- the effective size */

describe('the effective font size', () => {
  it('is the nominal size when the scale is 100%, fraction and all', () => {
    // The `frac` deck: 10.5pt at a stated 100% draws a 12.6pt advance, which is
    // 1.2 x 10.5 and not 1.2 x 11.
    for (const block of blocks.filter((b) => b.deck === 'frac' && b.id.startsWith('frac-'))) {
      if (block.fontScale !== undefined && block.fontScale !== 100000) continue;
      const eff = effectiveFontSize(block.sz, block.fontScale ?? 100000);
      expect(eff, block.id).toBe(block.sz / 100);
    }
  });

  it('rounds to a whole point when a scale applies, halves up', () => {
    // Both halves come from the fixture: the advance a stored scale drew, and
    // the arithmetic that has to produce it.
    const scaled = blocks.filter(
      (b) => b.autofit === 'norm' && b.fontScale !== undefined && b.fontScale !== 100000,
    );
    expect(scaled.length).toBeGreaterThan(15);
    for (const block of scaled) {
      const eff = effectiveFontSize(block.sz, block.fontScale);
      expect(Number.isInteger(eff), `${block.id} gave ${String(eff)}`).toBe(true);
    }
    // The two probes that land exactly on a half: 18pt at 25% is 4.5 and 28pt
    // at 37.5% is 10.5. Both draw the advance of the size above.
    expect(effectiveFontSize(1800, 25000)).toBe(5);
    expect(effectiveFontSize(2800, 37500)).toBe(11);
  });

  it('rejects a size or a scale the format cannot hold', () => {
    expect(() => effectiveFontSize(0)).toThrow(TextError);
    expect(() => effectiveFontSize(-1800)).toThrow(TextError);
    expect(() => effectiveFontSize(1800, -1)).toThrow(TextError);
    expect(() => effectiveFontSize(1800, Number.NaN)).toThrow(TextError);
  });
});

/* --------------------------------------------------------- the advance */

describe('the advance', () => {
  it('reproduces every advance PowerPoint drew under a stored scale', () => {
    // Only the probes whose gaps are line advances: the fixture omits the
    // number where paragraph spacing would be folded into it.
    const subset = blocks.filter((b) => b.advancePt !== undefined);
    expect(subset.length).toBeGreaterThan(500);
    for (const block of subset) {
      const eff = effectiveFontSize(block.sz, block.fontScale ?? 100000);
      const advance = autofitAdvance(eff, block.lnSpcReduction ?? 0, spacing(block.lnSpc));
      expect(advance, `${block.id} (${block.face} ${String(block.sz / 100)}pt)`).toBeCloseTo(
        measuredAdvance(block),
        2,
      );
    }
  });

  it('subtracts the reduction from the percentage and quantises afterwards', () => {
    // 149.6% reduced by 12.5% is 137.1, which rounds to 137 - not the 137.5
    // that quantising 149.6 to 150 first would give. The two are 0.066pt apart
    // at the size the fixture measured, and the fixture has the answer.
    const probe = blocks.find((b) => b.id === 'red-010');
    if (probe === undefined) throw new Error('the fixture lost red-010');
    const eff = effectiveFontSize(probe.sz, probe.fontScale ?? 100000);
    expect(autofitAdvance(eff, probe.lnSpcReduction, spacing(probe.lnSpc))).toBeCloseTo(
      measuredAdvance(probe),
      2,
    );
    const quantiseFirst = (1.5 - 0.125) * NATURAL_LINE_FACTOR * eff;
    expect(Math.abs(quantiseFirst - measuredAdvance(probe))).toBeGreaterThan(0.05);
  });

  it('ignores the reduction entirely when the spacing is an exact length', () => {
    const points: LineSpacing = { kind: 'points', value: 3000 };
    const at = effectiveFontSize(1800);
    for (const reduction of [0, 10000, 20000, 25000]) {
      expect(autofitAdvance(at, reduction, points)).toBe(30);
    }
  });

  it('floors the spacing at one per cent of the line box', () => {
    // A 15% spacing reduced by 20% is -5%. Measured at five sizes.
    for (const block of blocks.filter((b) => b.id.startsWith('zero-'))) {
      const eff = effectiveFontSize(block.sz);
      const advance = autofitAdvance(eff, block.lnSpcReduction ?? 0, spacing(block.lnSpc));
      expect(advance, block.id).toBeCloseTo(measuredAdvance(block), 2);
      expect(advance).toBeCloseTo(0.01 * NATURAL_LINE_FACTOR * eff, 3);
    }
  });

  it('rejects a spacing the format cannot hold', () => {
    expect(() => autofitAdvance(18, -1)).toThrow(TextError);
    expect(() => autofitAdvance(18, 0, { kind: 'points', value: -1 })).toThrow(TextError);
    expect(() => autofitAdvance(18, 0, { kind: 'percent', value: Number.NaN })).toThrow(TextError);
  });
});

/* ------------------------------------------------------ the block height */

describe('the height of a text body', () => {
  /** A block reading turned into the paragraphs the package takes. */
  function paragraphsOf(block: FixtureBlock): readonly ParagraphBox[] {
    // Every block probe uses one word a paragraph, so the lines divide evenly
    // and each paragraph is one line. The fixture states both counts and the
    // assertion below would fail loudly if they ever stopped agreeing.
    expect(block.lines, block.id).toBe(block.paragraphs);
    return Array.from({ length: block.paragraphs }, () => ({
      lines: 1,
      sz: block.sz,
      metrics: metricsFor(block.face),
      lnSpc: spacing(block.lnSpc),
      spcBef: spacing(block.spcBef),
      spcAft: spacing(block.spcAft),
    }));
  }

  it('reproduces every block PowerPoint laid out', () => {
    expect(blocks.length).toBeGreaterThan(500);
    let checked = 0;
    for (const block of blocks) {
      const scale = storedScale(block.fontScale, block.lnSpcReduction);
      const height = textBodyHeight(paragraphsOf(block), scale, block.spcFirstLastPara ?? false);
      expect(height, `${block.id} (${block.face} ${String(block.sz / 100)}pt)`).toBeCloseTo(
        block.blockHeightPt,
        1,
      );
      checked += 1;
    }
    expect(checked).toBe(blocks.length);
  });

  it('is not n x advance, and the fixture says by how much', () => {
    // The reading 3.2 could not tell apart from the right one, refuted here on
    // the probes where the two differ - which are exactly the ones whose line
    // spacing is not near single.
    const differing = blocks.filter(
      (b) => b.advancePt !== undefined && Math.abs(b.lines * b.advancePt - b.blockHeightPt) > 0.05,
    );
    expect(differing.length).toBeGreaterThan(100);
  });

  it('takes the whole advance for the last line at single spacing', () => {
    const metrics = metricsFor('Arial');
    const eff = effectiveFontSize(1800);
    const advance = autofitAdvance(eff);
    expect(lastLineHeight(advance, eff, metrics)).toBe(advance);
  });

  it('takes less than the advance once the spacing is looser than the floor', () => {
    const metrics = metricsFor('Arial');
    const eff = effectiveFontSize(1800);
    const advance = autofitAdvance(eff, 0, { kind: 'percent', value: 150000 });
    const last = lastLineHeight(advance, eff, metrics);
    expect(last).toBeLessThan(advance);
    expect(last).toBeCloseTo(0.75 * advance + metrics.lastLineCoefficient * eff, 6);
  });

  it('takes more than the advance once the spacing is tighter than the fitted line', () => {
    const metrics = metricsFor('Arial');
    const eff = effectiveFontSize(1800);
    const advance = autofitAdvance(eff, 0, { kind: 'percent', value: 60000 });
    expect(lastLineHeight(advance, eff, metrics)).toBeGreaterThan(advance);
  });

  it('rejects a paragraph with no lines', () => {
    const paragraph: ParagraphBox = { lines: 0, sz: 1800, metrics: metricsFor('Arial') };
    expect(() => textBodyHeight([paragraph], AUTOFIT_LADDER[0] as AutofitScale)).toThrow(TextError);
  });

  it('is zero for no paragraphs at all', () => {
    expect(textBodyHeight([], AUTOFIT_LADDER[0] as AutofitScale)).toBe(0);
  });
});

/* ------------------------------------------------------ paragraph spacing */

describe('paragraph spacing', () => {
  const spacingBlocks = blocks.filter((b) => b.spcBef !== undefined || b.spcAft !== undefined);

  it('has probes stating it, so the assertions below are about something', () => {
    expect(spacingBlocks.length).toBeGreaterThan(10);
  });

  it('drops the first spcBef and the last spcAft unless the flag says otherwise', () => {
    // Two probes differing only in the flag, whose heights differ by exactly
    // one spcBef - which is the whole of what spcFirstLastPara does.
    const off = blocks.find((b) => b.id === 'spacing-ctl037');
    const on = blocks.find((b) => b.id === 'spacing-ctl038');
    if (off === undefined || on === undefined) throw new Error('the fixture lost the flag pair');
    expect(off.spcFirstLastPara).toBe(false);
    expect(on.spcFirstLastPara).toBe(true);
    const spcBefPt = (off.spcBef?.value ?? 0) / 100;
    expect(on.blockHeightPt - off.blockHeightPt).toBeCloseTo(spcBefPt, 2);
  });

  it('treats an absent flag as false', () => {
    const absent = blocks.find((b) => b.id === 'spacing-ctl036');
    const stated = blocks.find((b) => b.id === 'spacing-ctl037');
    if (absent === undefined || stated === undefined) throw new Error('the fixture lost the pair');
    expect(absent.spcFirstLastPara).toBeUndefined();
    expect(absent.blockHeightPt).toBe(stated.blockHeightPt);
  });

  it('scales a percentage with the font and leaves an exact length alone', () => {
    const metrics = metricsFor('Arial');
    const paragraph = (spcBef: LineSpacing): ParagraphBox => ({
      lines: 1,
      sz: 1800,
      metrics,
      spcBef,
    });
    const scale: AutofitScale = { fontScale: 50000, lnSpcReduction: 0 };
    const full: AutofitScale = { fontScale: 100000, lnSpcReduction: 0 };
    const points: LineSpacing = { kind: 'points', value: 1200 };
    const percent: LineSpacing = { kind: 'percent', value: 100000 };
    // Two paragraphs, so the spacing between them is counted with the flag off.
    const pointsPair = [paragraph(points), paragraph(points)];
    const percentPair = [paragraph(percent), paragraph(percent)];
    const pointsDelta =
      textBodyHeight(pointsPair, full) - 2 * NATURAL_LINE_FACTOR * effectiveFontSize(1800);
    const pointsScaled =
      textBodyHeight(pointsPair, scale) - 2 * NATURAL_LINE_FACTOR * effectiveFontSize(1800, 50000);
    expect(pointsDelta).toBeCloseTo(12, 6);
    expect(pointsScaled).toBeCloseTo(12, 6);
    const percentDelta =
      textBodyHeight(percentPair, full) - 2 * NATURAL_LINE_FACTOR * effectiveFontSize(1800);
    const percentScaled =
      textBodyHeight(percentPair, scale) - 2 * NATURAL_LINE_FACTOR * effectiveFontSize(1800, 50000);
    expect(percentDelta).toBeCloseTo(NATURAL_LINE_FACTOR * 18, 6);
    expect(percentScaled).toBeCloseTo(NATURAL_LINE_FACTOR * 9, 6);
  });
});

/* -------------------------------------------------------- the replay */

describe('the rung PowerPoint chose', () => {
  /**
   * The whole experiment, run again.
   *
   * Every sweep is a configuration and a column of box heights; for each height
   * the implementation is asked to fit, and has to name the rung PowerPoint
   * wrote into the saved file. The line count is a constant of the probe - each
   * paragraph is one short word that cannot wrap at any scale - so the layout
   * callback returns the same paragraphs whatever rung it is handed, and the
   * fit test is the only thing under examination.
   */
  it('is what fitAutofit chooses, on every one of the swept boxes', () => {
    let probes = 0;
    const misses: string[] = [];
    for (const sweep of sweeps) {
      const metrics = metricsFor(sweep.face);
      const paragraphs: readonly ParagraphBox[] = Array.from({ length: sweep.lines }, () => ({
        lines: 1,
        sz: sweep.sz,
        metrics,
        lnSpc: spacing(sweep.lnSpc),
        spcBef: spacing(sweep.spcBef),
        spcAft: spacing(sweep.spcAft),
      }));
      for (const [index, boxPt] of sweep.boxPt.entries()) {
        const expectedIndex = sweep.rung[index];
        if (expectedIndex === undefined) throw new Error('the fixture columns disagree in length');
        const expected = ladder[expectedIndex];
        if (expected === undefined) throw new Error('the fixture names a rung it does not carry');
        const request = {
          bodyHeightPt: boxPt - sweep.insets.t - sweep.insets.b,
          spcFirstLastPara: sweep.spcFirstLastPara ?? false,
          layout: () => paragraphs,
        };
        const got = fitAutofit(request).scale;
        probes += 1;
        if (
          got.fontScale !== expected.fontScale ||
          got.lnSpcReduction !== expected.lnSpcReduction
        ) {
          if (misses.length < 8) {
            misses.push(
              `${sweep.deck} box=${String(boxPt)}pt: chose ${String(got.fontScale)}/` +
                `${String(got.lnSpcReduction)}, PowerPoint wrote ${String(expected.fontScale)}/` +
                String(expected.lnSpcReduction),
            );
          }
        }
      }
    }
    expect(misses).toEqual([]);
    expect(probes).toBe(sweeps.reduce((n, s) => n + s.boxPt.length, 0));
    expect(probes).toBeGreaterThan(2500);
  });

  it('subtracts the insets before testing, which the inset deck refutes otherwise', () => {
    const inset = sweeps.filter((s) => s.deck === 'inset' && s.insets.t + s.insets.b > 0);
    expect(inset.length).toBeGreaterThan(0);
    let wrongWithoutInsets = 0;
    for (const sweep of inset) {
      const metrics = metricsFor(sweep.face);
      const paragraphs: readonly ParagraphBox[] = Array.from({ length: sweep.lines }, () => ({
        lines: 1,
        sz: sweep.sz,
        metrics,
      }));
      for (const [index, boxPt] of sweep.boxPt.entries()) {
        const expectedIndex = sweep.rung[index];
        if (expectedIndex === undefined) throw new Error('unreachable');
        const expected = ladder[expectedIndex];
        if (expected === undefined) throw new Error('unreachable');
        const ignoring = fitAutofit({ bodyHeightPt: boxPt, layout: () => paragraphs }).scale;
        if (ignoring.fontScale !== expected.fontScale) wrongWithoutInsets += 1;
      }
    }
    expect(wrongWithoutInsets).toBeGreaterThan(0);
  });

  it('rewraps at each rung, so a wrapped body fits the rung it chose', () => {
    // The wrap deck's line counts are a consequence of the scale and cannot be
    // predicted here without the breaker, so what is asserted is the half that
    // needs no breaker: the rung PowerPoint chose does hold the text it laid
    // out. `tools/ground-truth/text/autofit/analyse.ts` proves the rung above overflows for 46 of the
    // 54, which is the other half.
    const wrap = fixture.wrap as readonly {
      id: string;
      heightPt: number;
      sz: number;
      face: string;
      lines: number;
      rung: number;
    }[];
    expect(wrap.length).toBeGreaterThan(20);
    for (const probe of wrap) {
      const rung = ladder[probe.rung];
      if (rung === undefined) throw new Error('unreachable');
      const height = textBodyHeight(
        [{ lines: probe.lines, sz: probe.sz, metrics: metricsFor(probe.face) }],
        rung,
      );
      expect(height, probe.id).toBeLessThanOrEqual(probe.heightPt + 1e-7);
    }
  });

  it('falls back to the last rung and says the text overflows', () => {
    const metrics = metricsFor('Arial');
    const paragraphs: readonly ParagraphBox[] = Array.from({ length: 40 }, () => ({
      lines: 1,
      sz: 1800,
      metrics,
    }));
    const result = fitAutofit({ bodyHeightPt: 10, layout: () => paragraphs });
    expect(result.scale).toEqual(AUTOFIT_LADDER[AUTOFIT_LADDER.length - 1]);
    expect(result.overflows).toBe(true);
    expect(result.heightPt).toBeGreaterThan(10);
  });

  it('reports the paragraphs and the height it settled on', () => {
    const metrics = metricsFor('Arial');
    const paragraphs: readonly ParagraphBox[] = [{ lines: 2, sz: 1800, metrics }];
    const result = fitAutofit({ bodyHeightPt: 500, layout: () => paragraphs });
    expect(result.scale).toEqual({ fontScale: 100000, lnSpcReduction: 0 });
    expect(result.overflows).toBe(false);
    expect(result.paragraphs).toBe(paragraphs);
    expect(result.heightPt).toBeCloseTo(2 * NATURAL_LINE_FACTOR * 18, 6);
  });

  it('rejects a body height that is not a length', () => {
    expect(() => fitAutofit({ bodyHeightPt: Number.NaN, layout: () => [] })).toThrow(TextError);
  });
});

/* -------------------------------------------------------------- view mode */

describe('a stored scale', () => {
  it('is applied as written, not snapped to a rung', () => {
    // 45%, 65% and 90% are not rungs, and the fixture has the advances
    // PowerPoint drew for them.
    const offLadder = blocks.filter(
      (b) =>
        b.fontScale !== undefined &&
        b.fontScale !== 100000 &&
        !ladder.some((r) => r.fontScale === b.fontScale),
    );
    expect(offLadder.length).toBeGreaterThan(2);
    for (const block of offLadder) {
      const scale = storedScale(block.fontScale, block.lnSpcReduction);
      expect(scale.fontScale).toBe(block.fontScale);
      const eff = effectiveFontSize(block.sz, scale.fontScale);
      expect(autofitAdvance(eff, scale.lnSpcReduction, spacing(block.lnSpc)), block.id).toBeCloseTo(
        measuredAdvance(block),
        2,
      );
    }
  });

  it('survives a scale no recomputation would produce', () => {
    // Two words in a box that holds twenty, stated at 25%. A viewer that
    // recomputed would draw 21.6pt lines; PowerPoint drew 4.8.
    const probe = blocks.find((b) => b.id === 'stored-x026');
    if (probe === undefined) throw new Error('the fixture lost stored-x026');
    const scale = storedScale(probe.fontScale, probe.lnSpcReduction);
    const eff = effectiveFontSize(probe.sz, scale.fontScale);
    expect(autofitAdvance(eff, scale.lnSpcReduction)).toBeCloseTo(measuredAdvance(probe), 2);
    expect(measuredAdvance(probe)).toBeLessThan(NATURAL_LINE_FACTOR * (probe.sz / 100));
  });

  it('leaves overflowing text overflowing', () => {
    const probe = blocks.find((b) => b.id === 'stored-x027');
    if (probe === undefined) throw new Error('the fixture lost stored-x027');
    expect(measuredAdvance(probe)).toBeCloseTo(NATURAL_LINE_FACTOR * (probe.sz / 100), 3);
    expect(probe.blockHeightPt).toBeGreaterThan(60);
  });

  it('reads absent attributes as 100% and 0%', () => {
    expect(storedScale()).toEqual({ fontScale: 100000, lnSpcReduction: 0 });
    expect(storedScale(undefined, 10000)).toEqual({ fontScale: 100000, lnSpcReduction: 10000 });
  });

  it('rejects a scale the format cannot hold', () => {
    expect(() => storedScale(-1)).toThrow(TextError);
    expect(() => storedScale(100000, -1)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------- spAutoFit */

describe('spAutoFit', () => {
  it('gives the shape the height PowerPoint wrote, on every probe it resized', () => {
    let checked = 0;
    for (const probe of spProbes) {
      // The probes PowerPoint left at their authored height are the hysteresis
      // measurement, not a height measurement; `requiredShapeHeight` answers
      // what the height should be, and the decision to apply it is a caller's.
      if (Math.abs(probe.emu / EMU_PER_POINT - probe.boxPt) < 1e-6) continue;
      const metrics = metricsFor(probe.face);
      const paragraphs: readonly ParagraphBox[] = Array.from({ length: probe.lines }, () => ({
        lines: 1,
        sz: probe.sz,
        metrics,
      }));
      const height = requiredShapeHeight(paragraphs, {
        topPt: probe.tIns,
        bottomPt: probe.bIns,
      });
      expect(height * EMU_PER_POINT, probe.id).toBeCloseTo(probe.emu, -1);
      checked += 1;
    }
    expect(checked).toBeGreaterThan(150);
  });

  it('uses a slack factor inside the bracket 181 probes measured', () => {
    const { factorMin, factorMax } = fixture.spAutoFit as { factorMin: number; factorMax: number };
    expect(SHAPE_AUTOFIT_SLACK).toBeGreaterThan(factorMin);
    expect(SHAPE_AUTOFIT_SLACK).toBeLessThan(factorMax);
  });

  it('scales the insets too, which is why the factor cannot be a font metric', () => {
    const metrics = metricsFor('Arial');
    const paragraphs: readonly ParagraphBox[] = [{ lines: 1, sz: 1800, metrics }];
    const bare = requiredShapeHeight(paragraphs, { topPt: 0, bottomPt: 0 });
    const inset = requiredShapeHeight(paragraphs, { topPt: 10, bottomPt: 10 });
    expect(inset - bare).toBeCloseTo(SHAPE_AUTOFIT_SLACK * 20, 6);
    expect(inset - bare).toBeGreaterThan(20);
  });

  it('rejects an inset that is not a length', () => {
    const paragraphs: readonly ParagraphBox[] = [
      { lines: 1, sz: 1800, metrics: metricsFor('Arial') },
    ];
    expect(() => requiredShapeHeight(paragraphs, { topPt: Number.NaN, bottomPt: 0 })).toThrow(
      TextError,
    );
  });
});

/* ------------------------------------------------- the refuted alternatives */

describe('the readings the measurement refutes', () => {
  const scores = fixture.scores as readonly { model: string; fits: number; total: number }[];

  it('has exactly one model that fits every case', () => {
    const perfect = scores.filter((s) => s.fits === s.total);
    expect(perfect.map((s) => s.model)).toEqual(['measured']);
  });

  it('records how wrong each alternative is, so a regression has a number', () => {
    const by = new Map(scores.map((s) => [s.model, s]));
    for (const model of [
      'size: no rounding',
      'size: round half to even',
      'reduction multiplies the advance',
      'block: n x advance',
      'block: floored at min(advance, the natural box)',
      'block: the floor threshold is 1.2 everywhere',
      'spcFirstLastPara defaults true',
      'the fit test ignores the insets',
      'take the tallest rung that fits',
    ]) {
      const score = by.get(model);
      expect(score, model).toBeDefined();
      expect(score?.fits, model).toBeLessThan(score?.total ?? 0);
    }
  });
});

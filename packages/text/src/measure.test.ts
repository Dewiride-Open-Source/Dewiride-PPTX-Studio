/**
 * The measurer, in the browser it ships to.
 *
 * These tests run in real Chromium rather than jsdom - `OffscreenCanvas` does
 * not exist in jsdom, and a measurer verified against a fake canvas would be
 * verified against nothing.
 *
 * Two kinds of assertion. The rules that are pure arithmetic over what the file
 * says - kerning thresholds, letter spacing, the CSS shorthand - are checked
 * against the readings in `corpus/ground-truth/text-metrics.json`. The advances
 * are checked against PowerPoint's own numbers, and that check is only as strong
 * as the fonts present: a face this machine does not have would be measured
 * through a substitute and agree with nothing. So the suite detects presence by
 * *fingerprint* rather than by asking, for the reason 3.7 will have to as well -
 * the fixture records that `document.fonts.check` returned true for a face that
 * is definitively not installed.
 */

import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/text-metrics.json' with { type: 'json' };
import { TextError } from './errors.js';
import {
  createCanvasMeasurer,
  cssFamily,
  cssFont,
  cssLetterSpacing,
  kerningEnabled,
} from './measure.js';

const measurer = createCanvasMeasurer();
const STRINGS = new Map(fixture.advances.strings.map((s) => [s.key, s.text]));

describe('the CSS the measurer drives the canvas with', () => {
  it('quotes the family, because a space is two identifiers and a generic is a generic', () => {
    expect(cssFamily('Times New Roman')).toBe('"Times New Roman"');
    expect(cssFamily('monospace')).toBe('"monospace"');
    expect(cssFont({ family: 'Arial', sz: 1800 })).toBe('18px "Arial"');
    expect(cssFont({ family: 'Arial', sz: 1050 })).toBe('10.5px "Arial"');
    expect(cssFont({ family: 'Georgia', sz: 3200, bold: true, italic: true })).toBe(
      'italic bold 32px "Georgia"',
    );
  });

  it('refuses a name that would break out of the shorthand', () => {
    expect(() => cssFamily('')).toThrow(TextError);
    expect(() => cssFamily('Ari"al')).toThrow(TextError);
    expect(() => cssFamily('Ari\\al')).toThrow(TextError);
    expect(() => cssFamily('Ari\nal')).toThrow(TextError);
    expect(() => cssFont({ family: 'Arial', sz: 0 })).toThrow(TextError);
  });

  it('writes @spc as an absolute length, not an em fraction', () => {
    expect(cssLetterSpacing(undefined)).toBe('0px');
    expect(cssLetterSpacing(0)).toBe('0px');
    expect(cssLetterSpacing(300)).toBe('3px');
    expect(cssLetterSpacing(-50)).toBe('-0.5px');
  });
});

describe('a:rPr/@kern is a minimum size, not a flag', () => {
  const readings = fixture.readings.kern;

  /**
   * Rebuild the on/off pair from the readings: at a given face and size the
   * narrowest measurement is kerned and the widest is not. Derived rather than
   * hard-coded so the test follows the fixture if the probe table changes.
   */
  function pair(font: string, sz: number): { on: number; off: number } {
    const widths = readings
      .filter((r) => r.font === font && r.sz === sz && r.kern !== null)
      .map((r) => r.widthPt);
    return { on: Math.min(...widths), off: Math.max(...widths) };
  }

  it('reproduces every kerning reading', () => {
    const withKern = readings.filter((r) => r.kern !== null);
    expect(withKern.length).toBeGreaterThan(10);
    let discriminating = 0;
    for (const r of withKern) {
      const { on, off } = pair(r.font, r.sz);
      if (Math.abs(on - off) < fixture.tolerancePt) continue; // a face with no kern pairs here
      discriminating += 1;
      const expected = kerningEnabled(r.kern ?? undefined, r.sz) ? on : off;
      expect(r.widthPt, `${r.font} ${String(r.sz)} kern=${String(r.kern)}`).toBeCloseTo(
        expected,
        2,
      );
    }
    expect(discriminating).toBeGreaterThan(8);
  });

  it('treats the threshold as inclusive, which one probe alone decides', () => {
    // A 12pt run with kern="1200" is kerned. The exclusive reading fits every
    // other probe and only fails here, which is why the probe table crosses the
    // threshold with the size rather than sitting beside it.
    expect(kerningEnabled(1200, 1200)).toBe(true);
    expect(kerningEnabled(1200, 1199)).toBe(false);
    const { on } = pair('Arial', 1200);
    const at = fixture.readings.kern.find(
      (r) => r.font === 'Arial' && r.sz === 1200 && r.kern === 1200,
    );
    expect(at?.widthPt).toBeCloseTo(on, 2);
  });

  it('reads kern="0" as never, which ECMA does not say', () => {
    expect(kerningEnabled(0, 4000)).toBe(false);
    const { off } = pair('Arial', 2400);
    const zero = fixture.readings.kern.find(
      (r) => r.font === 'Arial' && r.sz === 2400 && r.kern === 0,
    );
    expect(zero?.widthPt).toBeCloseTo(off, 2);
  });

  it('refutes the boolean readings with their scores', () => {
    const refuted = new Map(fixture.lineModel.refuted.kerning.map((r) => [r.model, r.fits]));
    expect(refuted.get('a boolean: on whenever the attribute is present')).toBe(5);
    expect(refuted.get('a boolean: on when kern is non-zero, at any size')).toBe(8);
    expect(refuted.get('on when kern > 0 and size > kern  (threshold, exclusive)')).toBe(11);
  });

  it('actually switches the canvas, and by the amount PowerPoint switches by', () => {
    // Every assertion above checks `kerningEnabled` on its own, which leaves the
    // wiring untested: a measurer that computed the flag correctly and then set
    // `fontKerning = 'normal'` unconditionally passed all of them. A mutation
    // sweep found exactly that, so this is the test that closes it.
    const text = 'AVATAR Yo To Wa';
    const arial = { family: 'Arial', sz: 1200 } as const;
    const kerned = measurer.measure(text, { ...arial, kern: 1200 }).width;
    const unkerned = measurer.measure(text, { ...arial, kern: 0 }).width;

    const { on, off } = pair('Arial', 1200);
    if (Math.abs(on - off) < fixture.tolerancePt) return; // no Arial here to measure with
    expect(kerned).toBeLessThan(unkerned);
    // The magnitude matters, not just the sign: PowerPoint's own two readings
    // differ by this much, and a measurer that toggled some other property
    // would move the number by something else.
    expect(unkerned - kerned).toBeCloseTo(off - on, 0);

    // And below the threshold the same run must stop being kerned, which is the
    // whole point of @kern being a size rather than a flag.
    const small = { family: 'Arial', sz: 800 } as const;
    expect(measurer.measure(text, { ...small, kern: 1200 }).width).toBeCloseTo(
      measurer.measure(text, { ...small, kern: 0 }).width,
      4,
    );
    expect(measurer.measure(text, { ...small, kern: 100 }).width).toBeLessThan(
      measurer.measure(text, { ...small, kern: 1200 }).width,
    );
  });

  it('records that a file stating no @kern still gets a threshold from the cascade', () => {
    // ECMA says an omitted @kern kerns at every size. Measured, an 8pt run in a
    // file that states none is *not* kerned, because 3.1's built-in p:txStyles
    // supply one. So `undefined` here means "the cascade yielded nothing", and
    // a caller that skipped the cascade will get the wrong answer.
    expect(fixture.lineModel.kernDefault.notKernedAtOrBelowHundredths).toBe(800);
    expect(fixture.lineModel.kernDefault.kernedAtOrAboveHundredths).toBe(1200);
    const { off } = pair('Arial', 800);
    const none = fixture.readings.kern.find(
      (r) => r.font === 'Arial' && r.sz === 800 && r.kern === null,
    );
    expect(none?.widthPt).toBeCloseTo(off, 2);
  });
});

describe('a:rPr/@spc', () => {
  it('is points after every character, the last included', () => {
    const readings = fixture.readings.spc;
    let checked = 0;
    for (const r of readings) {
      if (r.spc === 0) continue;
      const base = readings.find((q) => q.font === r.font && q.sz === r.sz && q.spc === 0);
      if (base === undefined) continue;
      checked += 1;
      // The reading is BoundWidth, which also covers the trailing paragraph
      // mark - and the mark is spaced too, so it is n + 1 gaps for n characters.
      const gaps = r.chars + 1;
      expect(
        r.widthPt - base.widthPt,
        `${r.font} ${String(r.sz)} spc=${String(r.spc)}`,
      ).toBeCloseTo((r.spc / 100) * gaps, 2);
    }
    expect(checked).toBeGreaterThan(10);
  });

  it('does not scale with the font size', () => {
    const readings = fixture.readings.spc;
    const delta = (font: string, sz: number, spc: number): number => {
      const at = readings.find((r) => r.font === font && r.sz === sz && r.spc === spc);
      const base = readings.find((r) => r.font === font && r.sz === sz && r.spc === 0);
      return (at?.widthPt ?? 0) - (base?.widthPt ?? 0);
    };
    // The same @spc widens the run by the same number of points at 12pt and at
    // 32pt. An em-relative reading would differ by a factor of 2.67.
    expect(delta('Arial', 1200, 300)).toBeCloseTo(delta('Arial', 3200, 300), 2);
  });

  it('is what Chromium letterSpacing does, which is why it is used directly', () => {
    const font = { family: 'Arial', sz: 1800 } as const;
    const plain = measurer.measure('0123456789', font).width;
    const spaced = measurer.measure('0123456789', { ...font, spc: 300 }).width;
    // Ten characters, ten gaps: Chromium spaces after the last one too.
    expect(spaced - plain).toBeCloseTo(30, 1);
  });
});

describe('advances against PowerPoint', () => {
  const entries = fixture.advances.entries;
  const perFont = new Map(fixture.browserAgreement.perFont.map((f) => [f.font, f]));

  /**
   * Is this the face PowerPoint measured, or a substitute?
   *
   * By fingerprint, not by asking: the fixture records that
   * `document.fonts.check` returned true for a face that is definitively not
   * installed, so the browser's own answer cannot be trusted. A face measuring
   * within 5% of PowerPoint on the reference string is the same face; the
   * recorded worst disagreement on a real face is 1.6% and the absent face was
   * out by 21%, so 5% separates them with room to spare.
   */
  function isTheSameFace(font: string): boolean {
    const at = entries.find((e) => e.font === font && e.sz === 1800);
    const text = STRINGS.get('hamburg');
    if (at === undefined || text === undefined) return false;
    const ours = measurer.measure(text, { family: font, sz: 1800 }).width;
    const theirs = at.widthsPt.hamburg;
    return Math.abs(ours - theirs) / theirs < 0.05;
  }

  const present = [...new Set(entries.map((e) => e.font))].filter(
    (f) => f !== fixture.substitution.absentFace && isTheSameFace(f),
  );

  it('resolves at least one reference face to measure against', () => {
    // Not a formality. On a machine with none of these faces this suite would
    // otherwise report success having compared nothing, and the fidelity
    // harness of 3.9 - which pins fonts by container digest for exactly this
    // reason - would be the first place anyone noticed.
    expect(
      present.length,
      'none of the reference faces resolved; see sub-phase 3.9 on pinning fonts',
    ).toBeGreaterThan(0);
  });

  it('agrees with PowerPoint within the disagreement the fixture records', () => {
    let compared = 0;
    for (const entry of entries) {
      if (!present.includes(entry.font)) continue;
      const limit = perFont.get(entry.font);
      if (limit === undefined) continue;
      for (const [key, theirs] of Object.entries(entry.widthsPt)) {
        const text = STRINGS.get(key);
        if (text === undefined) continue;
        compared += 1;
        const ours = measurer.measure(text, { family: entry.font, sz: entry.sz }).width;
        const relative = (Math.abs(ours - theirs) / theirs) * 100;
        // Per font, because the honest ceiling differs per font: Cambria at
        // 12pt is hinted by PowerPoint and not by Chromium, and no amount of
        // correctness here closes that.
        expect(relative, `${entry.font} ${String(entry.sz)} ${key}`).toBeLessThanOrEqual(
          limit.worstPct + 0.5,
        );
      }
    }
    expect(compared).toBeGreaterThan(20);
  });

  it('holds the recorded agreement to the number the plan promises to track', () => {
    expect(fixture.browserAgreement.summaryPct.median).toBeLessThan(0.25);
    expect(fixture.browserAgreement.summaryPct.p95).toBeLessThan(1);
    // Not asserted small: the maximum is Cambria at 12pt and it is a real
    // limitation, recorded so that it improves rather than being explained away.
    expect(fixture.browserAgreement.summaryPct.max).toBeGreaterThan(5);
  });

  it('scales linearly in the size, which is why zoom is a transform and not a re-layout', () => {
    const text = STRINGS.get('hamburg') ?? 'Hamburgefonstiv';
    const family = present[0];
    expect(family).toBeDefined();
    const at18 = measurer.measure(text, { family: family ?? 'Arial', sz: 1800 }).width;
    const at54 = measurer.measure(text, { family: family ?? 'Arial', sz: 5400 }).width;
    expect(at54 / at18).toBeCloseTo(3, 1);
  });
});

describe('substitution is visible in the advances and not in the browser', () => {
  it('records that the browser said a missing face was installed', () => {
    expect(fixture.substitution.browserReportedInstalled).toBe(true);
    expect(fixture.substitution.powerPointSubstitutes).toBe('Calibri');
    expect(fixture.substitution.distinctAdvanceRows).toBe(13);
    expect(fixture.substitution.facesMeasured).toBe(14);
  });

  it('and that the fingerprint catches what the browser missed', () => {
    const absent = fixture.advances.entries.find(
      (e) => e.font === fixture.substitution.absentFace && e.sz === 1800,
    );
    const calibri = fixture.advances.entries.find((e) => e.font === 'Calibri' && e.sz === 1800);
    expect(absent).toBeDefined();
    // PowerPoint measured the missing face exactly as Calibri - which is how
    // the substitution was detected at all, since nothing else reported it.
    expect(absent?.widthsPt.hamburg).toBeCloseTo(calibri?.widthsPt.hamburg ?? -1, 2);
  });
});

import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { indexFonts, systemFontDirectories, type FontLibrary } from './faces.js';
import { createFontMeasurer } from './measure.js';
import { FIXTURE, bytesOf, readFaces } from './sfnt.test.js';

/**
 * A directory of the exact fonts Chromium was measured on.
 *
 * Written to a temporary directory and indexed with `system: false`, so the
 * suite never reads a font off the machine running it - which would make the
 * result depend on what happens to be installed.
 */
let library: FontLibrary;

/** What draws a typeface no probe font is named after, and nothing stands in for. */
const LAST_RESORT = FIXTURE.fonts
  .map((font) => font.family)
  .sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : 1))[0]!;

beforeAll(() => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-t13-'));
  for (const font of FIXTURE.fonts) writeFileSync(join(dir, `${font.id}.ttf`), bytesOf(font.id));
  // Pinned to the platform T13 measured on, so the face box rows below assert
  // one rasteriser rather than whichever this suite runs on. ADR 0045.
  library = indexFonts({ extra: [dir], system: false, platform: 'win32' });
});

describe('the font library', () => {
  it('indexes one face per probe font', () => {
    expect(library.indexed).toHaveLength(FIXTURE.fonts.length);
  });

  it('resolves each probe by the family name in its own name table', () => {
    for (const font of FIXTURE.fonts) {
      const found = library.resolve(font.family, false, false);
      expect(found?.drawn, font.family).toBe(font.family);
      expect(found?.substituted, font.family).toBe(false);
    }
  });

  it('stands in for a family no directory holds, and says which face it used', () => {
    const found = library.resolve('Zzz Probe Face That Is Not Installed', false, false);
    expect(found?.asked).toBe('Zzz Probe Face That Is Not Installed');
    expect(found?.substituted).toBe(true);
    expect(found?.drawn).toBe(LAST_RESORT);
  });

  it('names this platform font directories without reading them', () => {
    for (const platform of ['win32', 'darwin', 'linux']) {
      expect(systemFontDirectories(platform).length, platform).toBeGreaterThan(0);
    }
    expect(systemFontDirectories('win32').some((d) => /Fonts$/i.test(d))).toBe(true);
  });

  it('takes the first --font-dir that has the family, so a caller can override', () => {
    const first = mkdtempSync(join(tmpdir(), 'pptx-studio-first-'));
    const second = mkdtempSync(join(tmpdir(), 'pptx-studio-second-'));
    // The same family in both directories, which is the only arrangement where
    // "first wins" and "last wins" give different answers.
    writeFileSync(join(first, 'a.ttf'), bytesOf('split'));
    writeFileSync(join(second, 'b.ttf'), bytesOf('split'));
    expect(
      indexFonts({ extra: [first, second], system: false }).resolve(
        'PptxStudio Split',
        false,
        false,
      )?.file,
    ).toContain('pptx-studio-first-');
    expect(
      indexFonts({ extra: [second, first], system: false }).resolve(
        'PptxStudio Split',
        false,
        false,
      )?.file,
    ).toContain('pptx-studio-second-');
  });

  it('refuses a --font-dir that is not there', () => {
    expect(() =>
      indexFonts({ extra: [join(tmpdir(), 'no-such-font-dir-8d21')], system: false }),
    ).toThrowError(/no such directory/);
  });
});

describe('measuring against what Chromium reported', () => {
  /**
   * The whole point of the experiment: 432 widths, and the measurer has to
   * reproduce every one of them from the font tables alone.
   */
  it('reproduces all 432 widths exactly', () => {
    const measurer = createFontMeasurer(library).measurer;
    let compared = 0;
    for (const row of FIXTURE.browser.rows) {
      const family = FIXTURE.fonts.find((f) => f.id === row.id)!.family;
      for (const px of FIXTURE.browser.sizes) {
        for (const text of FIXTURE.browser.strings) {
          const expected = row.widths[String(px)]?.[text];
          expect(expected, `${row.id} ${String(px)}px ${text}`).toBeTypeOf('number');
          const got = measurer.measure(text, { family, sz: px * 100 }).width;
          expect(got, `${row.id} ${String(px)}px ${JSON.stringify(text)}`).toBe(expected);
          compared += 1;
        }
      }
    }
    expect(compared).toBe(432);
  });

  it('quantises only non-negative advances, so trunc and floor are one rule', () => {
    let checked = 0;
    for (const font of FIXTURE.fonts) {
      const face = readFaces(bytesOf(font.id), font.id)[0]!;
      for (const px of FIXTURE.browser.sizes) {
        for (const text of FIXTURE.browser.strings) {
          for (const ch of [...text]) {
            const units = face.advanceOf(ch.codePointAt(0)!) ?? 0;
            expect(units, `${font.id} ${ch}`).toBeGreaterThanOrEqual(0);
            const x = (units * px) / face.metrics.unitsPerEm;
            expect(Math.floor(x), `${font.id} ${ch} ${String(px)}px`).toBe(Math.trunc(x));
            checked += 1;
          }
        }
      }
    }
    const perFont = FIXTURE.browser.strings.reduce((n, s) => n + [...s].length, 0);
    expect(checked).toBe(FIXTURE.fonts.length * FIXTURE.browser.sizes.length * perFont);
  });

  it('would not reproduce them with exact float arithmetic', () => {
    // The plausible implementation - sum advance x size / upem in doubles - and
    // the fixture says it fits 150 of 432. If this ever stops disagreeing, the
    // test above has stopped being evidence of anything.
    const font = FIXTURE.fonts.find((f) => f.id === 'plain')!;
    const naive = (text: string, px: number): number =>
      [...text].reduce(
        (sum, ch) =>
          sum +
          ((ch === ' ' ? font.spec.spaceAdvance : font.spec.advance) * px) / font.spec.unitsPerEm,
        0,
      );
    const row = FIXTURE.browser.rows.find((r) => r.id === 'plain')!;
    expect(naive('ABCDEFGH', 8)).not.toBe(row.widths['8']?.['ABCDEFGH']);
  });

  it('applies the kern the fixture recorded, and only where the pair has one', () => {
    const measurer = createFontMeasurer(library).measurer;
    const both = FIXTURE.fonts.find((f) => f.id === 'kern-and-gpos')!;
    const kerned = measurer.measure('AB', { family: both.family, sz: 100 * 100 }).width;
    const control = measurer.measure('AC', { family: both.family, sz: 100 * 100 }).width;
    // GPOS says -100 units on a 1000 unit em at 100px, so ten pixels.
    expect(control - kerned).toBeCloseTo(10, 9);
  });

  it('drops kerning below a:rPr/@kern, which is a minimum size', () => {
    const measurer = createFontMeasurer(library).measurer;
    const family = FIXTURE.fonts.find((f) => f.id === 'gpos-only')!.family;
    const on = measurer.measure('AB', { family, sz: 1800, kern: 1200 }).width;
    const off = measurer.measure('AB', { family, sz: 800, kern: 1200 }).width;
    const plain = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    expect(on).toBeLessThan(measurer.measure('AC', { family, sz: 1800 }).width);
    expect(off).toBe(measurer.measure('AB', { family: plain, sz: 800 }).width);
  });

  it('adds a:rPr/@spc after every character, the last included', () => {
    const measurer = createFontMeasurer(library).measurer;
    const family = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    const bare = measurer.measure('ABC', { family, sz: 1800 }).width;
    const spaced = measurer.measure('ABC', { family, sz: 1800, spc: 200 }).width;
    expect(spaced - bare).toBeCloseTo(3 * 2, 9);
  });
});

describe('the face box probe', () => {
  it('reports the browser fractions for every probe', () => {
    const { faceBox } = createFontMeasurer(library);
    const scale = FIXTURE.browser.boxPx;
    for (const row of FIXTURE.browser.rows) {
      const family = FIXTURE.fonts.find((f) => f.id === row.id)!.family;
      const box = faceBox.box(family);
      expect(box.ascent * scale, `${row.id} ascent`).toBeCloseTo(row.box.ascent, 9);
      expect(box.descent * scale, `${row.id} descent`).toBeCloseTo(row.box.descent, 9);
    }
  });

  /** What `packages/text/src/lines/baseline.ts` makes of one browser reading. */
  const offsetOf = (ideographicBaseline: number, px: number): number => {
    const offset = -ideographicBaseline / px;
    return offset > 0 && offset < 1 ? offset : 0;
  };

  it('reproduces the ideographic baseline Chromium reported for every probe', () => {
    const { faceBox } = createFontMeasurer(library);
    let compared = 0;
    for (const row of FIXTURE.browser.rows) {
      const family = FIXTURE.fonts.find((f) => f.id === row.id)!.family;
      for (const px of FIXTURE.browser.baselineSizes) {
        const want = offsetOf(row.baselines[String(px)]!.ideographic, px);
        const got = faceBox.box(family).ideographic;
        expect(got, `${row.id} ideographic at ${String(px)}px`).toBeCloseTo(want, 9);
        compared += 1;
      }
    }
    expect(compared).toBe(FIXTURE.browser.rows.length * FIXTURE.browser.baselineSizes.length);
  });

  it('answers a Han run exactly as it answers a Latin one, so caching by family is sound', () => {
    const { faceBox } = createFontMeasurer(library);
    let compared = 0;
    for (const row of FIXTURE.browser.rows) {
      if (row.hanBaselines === null) continue;
      const family = FIXTURE.fonts.find((f) => f.id === row.id)!.family;
      for (const px of FIXTURE.browser.baselineSizes) {
        const han = row.hanBaselines[String(px)]!.ideographic;
        expect(han, `${row.id} at ${String(px)}px`).toBe(row.baselines[String(px)]!.ideographic);
        expect(faceBox.box(family).ideographic, row.id).toBeCloseTo(offsetOf(han, px), 9);
        compared += 1;
      }
    }
    expect(compared).toBe(8);
  });

  it('is not the descent, which is what a BASE table moves it off', () => {
    const font = FIXTURE.fonts.find((f) => f.id === 'base-table')!;
    const box = createFontMeasurer(library).faceBox.box(font.family);
    expect(box.ideographic).not.toBe(box.descent);
    expect(box.ideographic * FIXTURE.browser.boxPx).toBeCloseTo(
      -(font.spec.base?.['DFLT']?.['ideo'] ?? 0),
      9,
    );
  });

  it('falls back to the descent for a face whose BASE says nothing about ideo', () => {
    const { faceBox } = createFontMeasurer(library);
    for (const id of ['split', 'base-no-dflt', 'base-no-ideo']) {
      const box = faceBox.box(FIXTURE.fonts.find((f) => f.id === id)!.family);
      expect(box.ideographic, id).toBe(box.descent);
    }
  });

  it('answers zero for a coordinate outside the em, where uprightPen would throw', () => {
    const font = FIXTURE.fonts.find((f) => f.id === 'base-outsize')!;
    const row = FIXTURE.browser.rows.find((r) => r.id === 'base-outsize')!;
    // Chromium hands a 1.2 em coordinate over as written rather than clamping
    // it, so both engines see the same unusable number and both answer zero.
    expect(-row.baselines[String(FIXTURE.browser.boxPx)]!.ideographic).toBe(1200);
    expect(1200 / font.spec.unitsPerEm).toBeGreaterThan(1);
    const box = createFontMeasurer(library).faceBox.box(font.family);
    expect(box.ideographic).toBe(0);
    expect(box.ideographic).not.toBe(box.descent);
  });
});

describe('what it says it did', () => {
  it('names the typefaces it drew and the files they came from', () => {
    const fonts = createFontMeasurer(library);
    const family = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    fonts.measurer.measure('A', { family, sz: 1800 });
    const used = fonts.used();
    expect(used).toHaveLength(1);
    expect(used[0]?.asked).toBe(family);
    expect(used[0]?.substituted).toBe(false);
    expect(used[0]?.file).toMatch(/plain\.ttf$/);
  });

  it('records a code point no indexed face can draw', () => {
    const fonts = createFontMeasurer(library);
    const family = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    fonts.measurer.measure('Z', { family, sz: 1800 });
    expect(fonts.missing()).toEqual(['Z'.codePointAt(0)]);
  });

  it('borrows a face that has the glyph before giving up on it', () => {
    // Every probe font has A..H, so a family that lacks nothing cannot show
    // this; `plain` measuring `A` through its own tables is the control.
    const fonts = createFontMeasurer(library);
    const family = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    expect(fonts.measurer.measure('A', { family, sz: 10000 }).width).toBeGreaterThan(0);
    expect(fonts.missing()).toEqual([]);
  });

  it('draws a typeface nothing stands in for, and reports the substitution', () => {
    const fonts = createFontMeasurer(library);
    const width = fonts.measurer.measure('A', { family: 'Nothing At All', sz: 1800 }).width;
    expect(width).toBeGreaterThan(0);
    const used = fonts.used();
    expect(used).toHaveLength(1);
    expect(used[0]?.asked).toBe('Nothing At All');
    expect(used[0]?.substituted).toBe(true);
    expect(used[0]?.drawn).toBe(LAST_RESORT);
  });

  it('measures and boxes a substituted typeface with the one face it chose', () => {
    // The measurer and the box probe are asked separately and only the family
    // name reaches both, so a choice that depended on the text would split them.
    const fonts = createFontMeasurer(library);
    expect(fonts.measurer.measure('ABC', { family: 'Nothing At All', sz: 1800 }).width).toBe(
      fonts.measurer.measure('ABC', { family: LAST_RESORT, sz: 1800 }).width,
    );
    expect(fonts.faceBox.box('Nothing At All')).toEqual(fonts.faceBox.box(LAST_RESORT));
  });

  it('refuses only when the library is empty, which is nothing to draw with', () => {
    const empty = indexFonts({
      extra: [mkdtempSync(join(tmpdir(), 'pptx-studio-empty-'))],
      system: false,
    });
    const fonts = createFontMeasurer(empty);
    expect(() => fonts.measurer.measure('A', { family: 'Nothing At All', sz: 1800 })).toThrowError(
      /no face on this machine .* the font library is empty/s,
    );
  });

  it('refuses a size the cascade could not have produced', () => {
    const fonts = createFontMeasurer(library);
    const family = FIXTURE.fonts.find((f) => f.id === 'plain')!.family;
    expect(() => fonts.measurer.measure('A', { family, sz: 0 })).toThrowError(
      /not a positive size/,
    );
  });
});

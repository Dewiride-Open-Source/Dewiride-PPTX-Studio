import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { indexFonts, systemFontDirectories, type FontLibrary } from './faces.js';
import { createFontMeasurer } from './measure.js';
import { FIXTURE, bytesOf } from './sfnt.test.js';

/**
 * A directory of the exact fonts Chromium was measured on.
 *
 * Written to a temporary directory and indexed with `system: false`, so the
 * suite never reads a font off the machine running it - which would make the
 * result depend on what happens to be installed.
 */
let library: FontLibrary;

beforeAll(() => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-studio-t13-'));
  for (const font of FIXTURE.fonts) writeFileSync(join(dir, `${font.id}.ttf`), bytesOf(font.id));
  library = indexFonts({ extra: [dir], system: false });
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

  it('reports nothing for a family no directory holds and nothing stands in for', () => {
    expect(library.resolve('Zzz Probe Face That Is Not Installed', false, false)).toBeUndefined();
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
   * The whole point of the experiment: 252 widths, and the measurer has to
   * reproduce every one of them from the font tables alone.
   */
  it('reproduces all 252 widths exactly', () => {
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
    expect(compared).toBe(252);
  });

  it('would not reproduce them with exact float arithmetic', () => {
    // The plausible implementation - sum advance x size / upem in doubles - and
    // the fixture says it fits 88 of 252. If this ever stops disagreeing, the
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

  it('puts the ideographic baseline on the descent, as Chromium does with no BASE table', () => {
    const { faceBox } = createFontMeasurer(library);
    // `split` has an ascent of 0.9 against a descent of 0.3, so the two cannot
    // be mistaken for one another the way they can in a face where they are close.
    const box = faceBox.box(FIXTURE.fonts.find((f) => f.id === 'split')!.family);
    expect(box.ideographic).toBe(box.descent);
    expect(box.ideographic).not.toBe(box.ascent);
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

  it('refuses a typeface nothing in the library can stand in for', () => {
    const fonts = createFontMeasurer(library);
    expect(() => fonts.measurer.measure('A', { family: 'Nothing At All', sz: 1800 })).toThrowError(
      /no face on this machine/,
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

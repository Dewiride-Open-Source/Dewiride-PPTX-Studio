import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import { repoPath } from '../../../../tools/repo/root.ts';
import { RenderError, isRenderError } from './errors.js';
import { facesIn, fontsIn, faceOf } from './sfnt.js';

/**
 * The measurement, not a copy of it.
 *
 * Every number asserted below comes out of `corpus/ground-truth/font-metrics.json`,
 * which carries the exact font bytes Chromium was shown and the widths it
 * reported. A test that re-derived the rule from `sfnt.ts` would pass whatever
 * `sfnt.ts` said. T13, ADR 0042.
 */
interface Baselines {
  readonly alphabetic: number;
  readonly hanging: number;
  readonly ideographic: number;
}

interface Fixture {
  readonly chromium: string;
  readonly faceBox: { readonly answer: string };
  readonly kerning: { readonly answer: string };
  readonly ideographic: { readonly answer: string };
  readonly fonts: readonly {
    readonly id: string;
    readonly family: string;
    readonly base64: string;
    readonly spec: {
      readonly unitsPerEm: number;
      readonly hheaAscender: number;
      readonly hheaDescender: number;
      readonly winAscent: number;
      readonly winDescent: number;
      readonly typoAscender: number;
      readonly typoDescender: number;
      readonly useTypoMetrics?: boolean;
      readonly advance: number;
      readonly spaceAdvance: number;
      readonly ideograph?: number;
      readonly base?: Record<string, Record<string, number>>;
    };
  }[];
  readonly browser: {
    readonly boxPx: number;
    readonly sizes: readonly number[];
    readonly baselineSizes: readonly number[];
    readonly strings: readonly string[];
    readonly rows: readonly {
      readonly id: string;
      readonly box: { readonly ascent: number; readonly descent: number };
      readonly widths: Record<string, Record<string, number>>;
      readonly baselines: Record<string, Baselines>;
      readonly hanBaselines: Record<string, Baselines & { width: number }> | null;
    }[];
  };
}

export const FIXTURE = JSON.parse(
  readFileSync(repoPath('corpus/ground-truth/font-metrics.json'), 'utf8'),
) as Fixture;

export function bytesOf(id: string): Uint8Array {
  const font = FIXTURE.fonts.find((f) => f.id === id);
  if (font === undefined) throw new Error(`no probe font ${id} in the fixture`);
  return new Uint8Array(Buffer.from(font.base64, 'base64'));
}

const A = 'A'.codePointAt(0)!;
const B = 'B'.codePointAt(0)!;
const C = 'C'.codePointAt(0)!;

describe('the font directory', () => {
  it('reads every probe font the fixture recorded', () => {
    for (const font of FIXTURE.fonts) {
      const faces = facesIn(bytesOf(font.id), font.id);
      expect(faces, font.id).toHaveLength(1);
      expect(faces[0]?.family, font.id).toBe(font.family);
    }
  });

  it('refuses bytes that carry no SFNT signature', () => {
    const notAFont = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0, 0, 0, 0, 0]);
    expect(() => facesIn(notAFont, 'a.zip')).toThrowError(RenderError);
    try {
      facesIn(notAFont, 'a.zip');
    } catch (error) {
      expect(isRenderError(error) && error.code).toBe('CLI_FONT_UNREADABLE');
      expect(isRenderError(error) && error.subject).toBe('a.zip');
    }
  });

  it('refuses a file too short to hold a directory', () => {
    expect(() => facesIn(new Uint8Array(4), 'stub.ttf')).toThrowError(/too short/);
  });

  it('exposes the same face through fontsIn and faceOf', () => {
    const tables = fontsIn(bytesOf('plain'), 'plain');
    expect(tables).toHaveLength(1);
    expect(faceOf(tables[0]!, 'plain').family).toBe('PptxStudio Plain');
  });
});

describe('cmap and hmtx', () => {
  it('maps the characters the font declares and no others', () => {
    const face = facesIn(bytesOf('plain'), 'plain')[0]!;
    const spec = FIXTURE.fonts.find((f) => f.id === 'plain')!.spec;
    expect(face.advanceOf(A)).toBe(spec.advance);
    expect(face.advanceOf(' '.codePointAt(0)!)).toBe(spec.spaceAdvance);
    // The probe fonts carry A..H and space; anything else has no glyph, and
    // `undefined` rather than zero is what lets a caller fall back.
    expect(face.advanceOf('Z'.codePointAt(0)!)).toBeUndefined();
    expect(face.advanceOf(0x4e00)).toBeUndefined();
  });

  it('prefers the format 12 subtable, which is the only one that reaches past the BMP', () => {
    // U+20000 is in the font's format 12 subtable and cannot be in its format 4
    // one, so a reader that takes the first subtable it recognises answers "no
    // glyph" here. The format 4 table is written first on purpose.
    const face = facesIn(bytesOf('astral'), 'astral')[0]!;
    expect(face.advanceOf(0x20000)).toBeTypeOf('number');
    expect(face.advanceOf(A)).toBeTypeOf('number');
    expect(facesIn(bytesOf('plain'), 'plain')[0]!.advanceOf(0x20000)).toBeUndefined();
  });

  it('maps the ideograph in the fonts that declare one, and no others', () => {
    for (const font of FIXTURE.fonts) {
      const face = facesIn(bytesOf(font.id), font.id)[0]!;
      const drawn = face.advanceOf(font.spec.ideograph ?? 0x4e00);
      expect(drawn, font.id).toBe(
        font.spec.ideograph === undefined ? undefined : font.spec.advance,
      );
    }
  });

  it('reads a maximal advance as unsigned, which is why flooring cannot differ', () => {
    // `hmtx.advanceWidth` is a UFWORD, so 0xFFFF is 65535 and never -1. No
    // probe font has an advance over 32767, so nothing else here can say so.
    const bytes = bytesOf('plain');
    const hmtx = fontsIn(bytes, 'plain')[0]!.get('hmtx')!;
    hmtx[8] = 0xff;
    hmtx[9] = 0xff;
    expect(faceOf(fontsIn(bytes, 'plain')[0]!, 'plain').advanceOf(A)).toBe(65535);
  });

  it('reports unitsPerEm from head', () => {
    for (const font of FIXTURE.fonts) {
      const face = facesIn(bytesOf(font.id), font.id)[0]!;
      expect(face.metrics.unitsPerEm, font.id).toBe(font.spec.unitsPerEm);
    }
  });
});

describe('the face box', () => {
  /** The reading the fixture settled: usWin, unless fsSelection bit 7 is set. */
  it('agrees with what Chromium reported for every probe', () => {
    const { boxPx } = FIXTURE.browser;
    for (const row of FIXTURE.browser.rows) {
      const font = FIXTURE.fonts.find((f) => f.id === row.id)!;
      const face = facesIn(bytesOf(row.id), row.id)[0]!;
      const scale = boxPx / face.metrics.unitsPerEm;
      expect(face.metrics.ascent * scale, `${row.id} ascent`).toBeCloseTo(row.box.ascent, 9);
      expect(face.metrics.descent * scale, `${row.id} descent`).toBeCloseTo(row.box.descent, 9);
      expect(face.metrics.source, row.id).toBe(
        font.spec.useTypoMetrics === true ? 'sTypo' : 'usWin',
      );
    }
  });

  it('is not hhea, which is the reading everybody writes first', () => {
    const split = FIXTURE.fonts.find((f) => f.id === 'split')!;
    const face = facesIn(bytesOf('split'), 'split')[0]!;
    expect(face.metrics.ascent).not.toBe(split.spec.hheaAscender);
    expect(face.metrics.ascent).toBe(split.spec.winAscent);
  });

  it('follows fsSelection bit 7 onto the typographic metrics', () => {
    const font = FIXTURE.fonts.find((f) => f.id === 'split-usetypo')!;
    const face = facesIn(bytesOf('split-usetypo'), 'split-usetypo')[0]!;
    expect(font.spec.useTypoMetrics).toBe(true);
    expect(face.metrics.ascent).toBe(font.spec.typoAscender);
    expect(face.metrics.descent).toBe(-font.spec.typoDescender);
    // The same bytes with the bit clear read the other pair, so the bit is what
    // moved the answer rather than the two fonts differing some other way.
    const clear = facesIn(bytesOf('split'), 'split')[0]!;
    expect(clear.metrics.ascent).toBe(font.spec.winAscent);
  });
});

describe('kerning', () => {
  it('reads a legacy kern table', () => {
    const face = facesIn(bytesOf('kern-only'), 'kern-only')[0]!;
    expect(face.kernBetween(A, B)).toBe(-200);
    expect(face.kernBetween(A, C)).toBe(0);
  });

  it('reads a GPOS PairPos', () => {
    const face = facesIn(bytesOf('gpos-only'), 'gpos-only')[0]!;
    expect(face.kernBetween(A, B)).toBe(-200);
    expect(face.kernBetween(A, C)).toBe(0);
  });

  it('takes GPOS over the legacy table when they disagree, as Chromium does', () => {
    const face = facesIn(bytesOf('kern-and-gpos'), 'kern-and-gpos')[0]!;
    expect(FIXTURE.kerning.answer).toBe('GPOS when present, else the kern table');
    // The font says -200 in `kern` and -100 in GPOS. Reading the legacy table
    // first is the cheaper implementation and the wrong one.
    expect(face.kernBetween(A, B)).toBe(-100);
    expect(face.kernBetween(A, B)).not.toBe(-200);
  });

  it('is zero for a font that declares none', () => {
    const face = facesIn(bytesOf('plain'), 'plain')[0]!;
    expect(face.kernBetween(A, B)).toBe(0);
  });

  it('is zero for a pair the font has no glyph for', () => {
    const face = facesIn(bytesOf('gpos-only'), 'gpos-only')[0]!;
    expect(face.kernBetween(A, 'Z'.codePointAt(0)!)).toBe(0);
  });
});

describe('the ideographic baseline', () => {
  const ideoOf = (id: string): number | undefined =>
    FIXTURE.fonts.find((f) => f.id === id)?.spec.base?.['DFLT']?.['ideo'];

  it('is the BASE ideo coordinate of the DFLT script for every probe that has one', () => {
    expect(FIXTURE.ideographic.answer).toBe(
      'the BASE ideo coordinate of the DFLT script, else the face box descent',
    );
    for (const font of FIXTURE.fonts) {
      const face = facesIn(bytesOf(font.id), font.id)[0]!;
      const coordinate = ideoOf(font.id);
      expect(face.metrics.ideographic, font.id).toBe(
        coordinate === undefined ? undefined : -coordinate,
      );
    }
  });

  it('is not the run script, the icfb tag, or anything a metric could supply', () => {
    const spec = FIXTURE.fonts.find((f) => f.id === 'base-table')!.spec;
    const face = facesIn(bytesOf('base-table'), 'base-table')[0]!;
    // The font says -450 under DFLT, -410 under latn and -380 under hani; only
    // one of the three is what Chromium read.
    expect(face.metrics.ideographic).toBe(-(spec.base?.['DFLT']?.['ideo'] ?? 0));
    expect(face.metrics.ideographic).not.toBe(-(spec.base?.['latn']?.['ideo'] ?? 0));
    expect(face.metrics.ideographic).not.toBe(-(spec.base?.['hani']?.['ideo'] ?? 0));
    expect(face.metrics.ideographic).not.toBe(-(spec.base?.['DFLT']?.['icfb'] ?? 0));
    expect(face.metrics.ideographic).not.toBe(face.metrics.descent);
  });

  it('is nothing at all for a BASE table that names every script but DFLT', () => {
    // latn, hani and kana each carry a coordinate and Chromium reads none of them.
    expect(
      facesIn(bytesOf('base-no-dflt'), 'base-no-dflt')[0]!.metrics.ideographic,
    ).toBeUndefined();
    expect(ideoOf('base-no-dflt')).toBeUndefined();
  });

  it('is nothing at all for a BASE table carrying no ideo tag', () => {
    expect(
      facesIn(bytesOf('base-no-ideo'), 'base-no-ideo')[0]!.metrics.ideographic,
    ).toBeUndefined();
    expect(FIXTURE.fonts.find((f) => f.id === 'base-no-ideo')?.spec.base?.['DFLT']?.['hang']).toBe(
      610,
    );
  });

  it('is nothing at all for a font with no BASE table', () => {
    for (const id of ['plain', 'split', 'split-usetypo', 'astral']) {
      expect(facesIn(bytesOf(id), id)[0]!.metrics.ideographic, id).toBeUndefined();
    }
  });
});

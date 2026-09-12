import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

import { repoPath } from '../../../../tools/repo/root.ts';
import { RenderError, isRenderError } from './errors.js';
import { backendFor, facesIn, fontsIn, faceOf } from './sfnt.js';

/** T13 was measured on Windows, so every fixture row below is the DirectWrite reading. */
export const readFaces = (bytes: Uint8Array, subject: string) =>
  facesIn(bytes, subject, 'directwrite');

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
  readonly faceBox: {
    readonly answer: string;
    readonly rounding: string;
    readonly scores: readonly { readonly model: string; readonly missed: readonly string[] }[];
  };
  readonly kerning: {
    readonly answer: string;
    readonly scores: readonly {
      readonly model: string;
      readonly fits: number;
      readonly of: number;
    }[];
  };
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
      const faces = readFaces(bytesOf(font.id), font.id);
      expect(faces, font.id).toHaveLength(1);
      expect(faces[0]?.family, font.id).toBe(font.family);
    }
  });

  it('refuses bytes that carry no SFNT signature', () => {
    const notAFont = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0, 0, 0, 0, 0]);
    expect(() => readFaces(notAFont, 'a.zip')).toThrowError(RenderError);
    try {
      readFaces(notAFont, 'a.zip');
    } catch (error) {
      expect(isRenderError(error) && error.code).toBe('CLI_FONT_UNREADABLE');
      expect(isRenderError(error) && error.subject).toBe('a.zip');
    }
  });

  it('refuses a file too short to hold a directory', () => {
    expect(() => readFaces(new Uint8Array(4), 'stub.ttf')).toThrowError(/too short/);
  });

  it('exposes the same face through fontsIn and faceOf', () => {
    const tables = fontsIn(bytesOf('plain'), 'plain');
    expect(tables).toHaveLength(1);
    expect(faceOf(tables[0]!, 'plain', 'directwrite').family).toBe('PptxStudio Plain');
  });
});

describe('cmap and hmtx', () => {
  it('maps the characters the font declares and no others', () => {
    const face = readFaces(bytesOf('plain'), 'plain')[0]!;
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
    const face = readFaces(bytesOf('astral'), 'astral')[0]!;
    expect(face.advanceOf(0x20000)).toBeTypeOf('number');
    expect(face.advanceOf(A)).toBeTypeOf('number');
    expect(readFaces(bytesOf('plain'), 'plain')[0]!.advanceOf(0x20000)).toBeUndefined();
  });

  it('maps the ideograph in the fonts that declare one, and no others', () => {
    for (const font of FIXTURE.fonts) {
      const face = readFaces(bytesOf(font.id), font.id)[0]!;
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
    expect(faceOf(fontsIn(bytes, 'plain')[0]!, 'plain', 'directwrite').advanceOf(A)).toBe(65535);
  });

  it('reports unitsPerEm from head', () => {
    for (const font of FIXTURE.fonts) {
      const face = readFaces(bytesOf(font.id), font.id)[0]!;
      expect(face.metrics.unitsPerEm, font.id).toBe(font.spec.unitsPerEm);
    }
  });
});

describe('the face box', () => {
  /** The reading the fixture settled: usWin, unless fsSelection bit 7 is set. */
  it('reads the pair the fixture named, on every probe', () => {
    expect(FIXTURE.faceBox.answer).toBe('usWin, or sTypo when fsSelection bit 7 is set');
    for (const font of FIXTURE.fonts) {
      const { spec } = font;
      const face = readFaces(bytesOf(font.id), font.id)[0]!;
      const typo = spec.useTypoMetrics === true;
      expect(face.metrics.ascent, `${font.id} ascent`).toBe(
        typo ? spec.typoAscender : spec.winAscent,
      );
      expect(face.metrics.descent, `${font.id} descent`).toBe(
        typo ? -spec.typoDescender : spec.winDescent,
      );
      expect(face.metrics.source, font.id).toBe(typo ? 'sTypo' : 'usWin');
    }
  });

  it('answers in font units, unrounded, so the measurer can round at the size the browser reads', () => {
    // 1902 on a 2048 em is 928.7109375 px at the probe size, and the browser
    // said 929; the rounding belongs to the probe, not to the table reader.
    const { spec } = FIXTURE.fonts.find((f) => f.id === 'split-2048')!;
    const row = FIXTURE.browser.rows.find((r) => r.id === 'split-2048')!;
    const face = readFaces(bytesOf('split-2048'), 'split-2048')[0]!;
    const scale = FIXTURE.browser.boxPx / spec.unitsPerEm;
    expect(face.metrics.ascent).toBe(spec.winAscent);
    expect(Number.isInteger(face.metrics.ascent * scale)).toBe(false);
    expect(face.metrics.ascent * scale).not.toBe(row.box.ascent);
    expect(Number.isInteger(row.box.ascent)).toBe(true);
  });

  it('is not hhea through DirectWrite, which is the reading everybody writes first', () => {
    const split = FIXTURE.fonts.find((f) => f.id === 'split')!;
    const face = readFaces(bytesOf('split'), 'split')[0]!;
    expect(face.metrics.ascent).not.toBe(split.spec.hheaAscender);
    expect(face.metrics.ascent).toBe(split.spec.winAscent);
  });

  it('reads hhea through FreeType and usWin through DirectWrite, from one file', () => {
    const { spec } = FIXTURE.fonts.find((f) => f.id === 'split')!;
    const freetype = facesIn(bytesOf('split'), 'split', 'freetype')[0]!;
    const directwrite = facesIn(bytesOf('split'), 'split', 'directwrite')[0]!;
    expect(spec.hheaAscender).not.toBe(spec.winAscent);
    expect(freetype.metrics.source).toBe('hhea');
    expect(freetype.metrics.ascent).toBe(spec.hheaAscender);
    expect(directwrite.metrics.source).toBe('usWin');
    expect(directwrite.metrics.ascent).toBe(spec.winAscent);
  });

  it('follows fsSelection bit 7 onto sTypo through DirectWrite and FreeType', () => {
    const { spec } = FIXTURE.fonts.find((f) => f.id === 'split-usetypo')!;
    for (const backend of ['freetype', 'directwrite'] as const) {
      const face = facesIn(bytesOf('split-usetypo'), 'split-usetypo', backend)[0]!;
      expect(face.metrics.source, backend).toBe('sTypo');
      expect(face.metrics.backend, backend).toBe(backend);
      expect(face.metrics.ascent, backend).toBe(spec.typoAscender);
    }
  });

  it('reads hhea through CoreText whatever bit 7 says, which only this probe can show', () => {
    // On macos-latest split-usetypo was the one face of 248 with bit 7 set and
    // hhea apart from sTypo; CoreText read 800/200. ADR 0053.
    const { spec } = FIXTURE.fonts.find((f) => f.id === 'split-usetypo')!;
    const face = facesIn(bytesOf('split-usetypo'), 'split-usetypo', 'coretext')[0]!;
    expect(spec.useTypoMetrics).toBe(true);
    expect(spec.hheaAscender).not.toBe(spec.typoAscender);
    expect(face.metrics.source).toBe('hhea');
    expect(face.metrics.backend).toBe('coretext');
    expect(face.metrics.ascent).toBe(spec.hheaAscender);
    expect(face.metrics.descent).toBe(-spec.hheaDescender);
    expect(face.metrics.candidates.useTypoMetrics).toBe(true);
  });

  it('names FreeType on Linux, CoreText on macOS and DirectWrite everywhere else', () => {
    expect(backendFor('linux')).toBe('freetype');
    expect(backendFor('darwin')).toBe('coretext');
    expect(backendFor('win32')).toBe('directwrite');
    expect(backendFor('freebsd')).toBe('directwrite');
  });
  it('carries the shape the IPA Gothic faces have, so nothing new has to be built', () => {
    // Three distinct descents, and a usWin box taller than the em: 900 + 300 on
    // a 1000 em, where ipag.ttf is 1802 + 401 on a 2048 one. ADR 0044.
    const spec = FIXTURE.fonts.find((f) => f.id === 'split')!.spec;
    const descents = [-spec.hheaDescender, spec.winDescent, -spec.typoDescender];
    expect(new Set(descents).size).toBe(3);
    expect(spec.winAscent + spec.winDescent).toBeGreaterThan(spec.unitsPerEm);
    expect(FIXTURE.browser.rows.find((row) => row.id === 'split')?.box).toEqual({
      ascent: spec.winAscent,
      descent: spec.winDescent,
    });
  });

  it('follows fsSelection bit 7 onto the typographic metrics', () => {
    const font = FIXTURE.fonts.find((f) => f.id === 'split-usetypo')!;
    const face = readFaces(bytesOf('split-usetypo'), 'split-usetypo')[0]!;
    expect(font.spec.useTypoMetrics).toBe(true);
    expect(face.metrics.ascent).toBe(font.spec.typoAscender);
    expect(face.metrics.descent).toBe(-font.spec.typoDescender);
    // The same bytes with the bit clear read the other pair, so the bit is what
    // moved the answer rather than the two fonts differing some other way.
    const clear = readFaces(bytesOf('split'), 'split')[0]!;
    expect(clear.metrics.ascent).toBe(font.spec.winAscent);
  });
});

describe('the candidate pairs', () => {
  it('hands back what each of the three tables said, chosen or not', () => {
    for (const font of FIXTURE.fonts) {
      const { spec } = font;
      const { candidates } = readFaces(bytesOf(font.id), font.id)[0]!.metrics;
      expect(candidates.hhea, font.id).toEqual({
        ascent: spec.hheaAscender,
        descent: -spec.hheaDescender,
      });
      expect(candidates.usWin, font.id).toEqual({
        ascent: spec.winAscent,
        descent: spec.winDescent,
      });
      expect(candidates.sTypo, font.id).toEqual({
        ascent: spec.typoAscender,
        descent: -spec.typoDescender,
      });
      expect(candidates.useTypoMetrics, font.id).toBe(spec.useTypoMetrics === true);
    }
  });

  it('keeps the three apart on the probe built to disagree with itself', () => {
    const { candidates } = readFaces(bytesOf('split'), 'split')[0]!.metrics;
    const descents = [candidates.hhea, candidates.usWin, candidates.sTypo].map(
      (pair) => pair?.descent,
    );
    expect(descents).toEqual([200, 300, 100]);
  });

  it('is where the chosen box came from, on every probe', () => {
    for (const font of FIXTURE.fonts) {
      const { metrics } = readFaces(bytesOf(font.id), font.id)[0]!;
      const chosen =
        metrics.source === 'sTypo'
          ? metrics.candidates.sTypo
          : metrics.source === 'hhea'
            ? metrics.candidates.hhea
            : metrics.candidates.usWin;
      expect(chosen, font.id).toEqual({ ascent: metrics.ascent, descent: metrics.descent });
    }
  });

  it('names hhea alone for a face carrying no OS/2 table', () => {
    const tables = fontsIn(bytesOf('split'), 'split')[0]!;
    const bare = new Map([...tables].filter(([tag]) => tag !== 'OS/2'));
    const { metrics } = faceOf(bare, 'split', 'directwrite');
    expect(metrics.candidates.usWin).toBeUndefined();
    expect(metrics.candidates.sTypo).toBeUndefined();
    expect(metrics.candidates.useTypoMetrics).toBe(false);
    expect(metrics.ascent).toBe(metrics.candidates.hhea.ascent);
    expect(metrics.descent).toBe(metrics.candidates.hhea.descent);
    // The diagnostics name the table the numbers came out of, and hhea is the
    // only one this face has, through either rasteriser.
    expect(metrics.source).toBe('hhea');
    expect(faceOf(bare, 'split', 'freetype').metrics.source).toBe('hhea');
  });

  it('reads no OS/2 shorter than the 78 bytes of version 0', () => {
    const tables = fontsIn(bytesOf('split'), 'split')[0]!;
    const short = new Map(tables).set('OS/2', tables.get('OS/2')!.subarray(0, 77));
    expect(faceOf(short, 'split', 'directwrite').metrics.candidates.usWin).toBeUndefined();
    expect(faceOf(new Map(tables), 'split', 'directwrite').metrics.candidates.usWin).toEqual({
      ascent: 900,
      descent: 300,
    });
  });
});

describe('kerning', () => {
  it('reads a legacy kern table', () => {
    const face = readFaces(bytesOf('kern-only'), 'kern-only')[0]!;
    expect(face.kernBetween(A, B)).toBe(-200);
    expect(face.kernBetween(A, C)).toBe(0);
  });

  it('reads a GPOS PairPos', () => {
    const face = readFaces(bytesOf('gpos-only'), 'gpos-only')[0]!;
    expect(face.kernBetween(A, B)).toBe(-200);
    expect(face.kernBetween(A, C)).toBe(0);
  });

  it('takes GPOS over the legacy table when they disagree, as Chromium does', () => {
    const face = readFaces(bytesOf('kern-and-gpos'), 'kern-and-gpos')[0]!;
    expect(FIXTURE.kerning.answer).toBe(
      'GPOS when it yields a pair, else the kern table, following one type 9 Extension lookup',
    );
    // The font says -200 in `kern` and -100 in GPOS. Reading the legacy table
    // first is the cheaper implementation and the wrong one.
    expect(face.kernBetween(A, B)).toBe(-100);
    expect(face.kernBetween(A, B)).not.toBe(-200);
  });

  it('follows a type 9 Extension lookup to the PairPos behind it', () => {
    const face = readFaces(bytesOf('gpos-extension'), 'gpos-extension')[0]!;
    expect(face.kernBetween(A, B)).toBe(-200);
    expect(face.kernBetween(A, C)).toBe(0);
  });

  it('reads the Extension probe exactly as Chromium measured the plain one', () => {
    // Two fonts alike but for the wrapper, and Chromium drew them the same
    // width at every size, so the pair behind the extension is honoured.
    const wrapped = FIXTURE.browser.rows.find((r) => r.id === 'gpos-extension')!;
    const plain = FIXTURE.browser.rows.find((r) => r.id === 'gpos-only')!;
    for (const px of FIXTURE.browser.sizes) {
      for (const text of FIXTURE.browser.strings) {
        expect(wrapped.widths[String(px)]?.[text], `${text} at ${String(px)}px`).toBe(
          plain.widths[String(px)]?.[text],
        );
      }
    }
    expect(readFaces(bytesOf('gpos-extension'), 'gpos-extension')[0]!.kernBetween(A, B)).toBe(
      readFaces(bytesOf('gpos-only'), 'gpos-only')[0]!.kernBetween(A, B),
    );
  });

  it('scores the reading that skips type 9 below the one that follows it', () => {
    const fitsOf = (model: string): number =>
      FIXTURE.kerning.scores.find((s) => s.model === model)?.fits ?? -1;
    const answer = FIXTURE.kerning.scores.find((s) => s.model === FIXTURE.kerning.answer)!;
    expect(answer.fits).toBe(answer.of);
    expect(
      fitsOf('GPOS when it yields a pair, else the kern table, type 2 lookups only'),
    ).toBeLessThan(answer.of);
  });

  it('does not follow an Extension that names another Extension', () => {
    // ExtensionPosFormat1.extensionLookupType, 2 patched to 9. The specification
    // forbids that nesting, so the pair is unreadable rather than a level down.
    const bytes = bytesOf('gpos-extension');
    const gpos = fontsIn(bytes, 'gpos-extension')[0]!.get('GPOS')!;
    const view = new DataView(gpos.buffer, gpos.byteOffset, gpos.byteLength);
    const lookupList = view.getUint16(8);
    const lookup = lookupList + view.getUint16(lookupList + 2);
    const subtable = lookup + view.getUint16(lookup + 6);
    expect(view.getUint16(lookup)).toBe(9);
    expect(view.getUint16(subtable + 2)).toBe(2);
    view.setUint16(subtable + 2, 9);
    const face = faceOf(fontsIn(bytes, 'gpos-extension')[0]!, 'gpos-extension', 'directwrite');
    expect(face.kernBetween(A, B)).toBe(0);
  });

  it('is zero for a font that declares none', () => {
    const face = readFaces(bytesOf('plain'), 'plain')[0]!;
    expect(face.kernBetween(A, B)).toBe(0);
  });

  it('is zero for a pair the font has no glyph for', () => {
    const face = readFaces(bytesOf('gpos-only'), 'gpos-only')[0]!;
    expect(face.kernBetween(A, 'Z'.codePointAt(0)!)).toBe(0);
  });
});

describe('the ideographic baseline', () => {
  const ideoOf = (id: string): number | undefined =>
    FIXTURE.fonts.find((f) => f.id === id)?.spec.base?.['DFLT']?.['ideo'];

  it('is the BASE ideo coordinate of the DFLT script for every probe that has one', () => {
    expect(FIXTURE.ideographic.answer).toBe(
      'the BASE ideo coordinate of the DFLT script, else the face box descent, rounded as the box is',
    );
    for (const font of FIXTURE.fonts) {
      const face = readFaces(bytesOf(font.id), font.id)[0]!;
      const coordinate = ideoOf(font.id);
      expect(face.metrics.ideographic, font.id).toBe(
        coordinate === undefined ? undefined : -coordinate,
      );
    }
  });

  it('is not the run script, the icfb tag, or anything a metric could supply', () => {
    const spec = FIXTURE.fonts.find((f) => f.id === 'base-table')!.spec;
    const face = readFaces(bytesOf('base-table'), 'base-table')[0]!;
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
      readFaces(bytesOf('base-no-dflt'), 'base-no-dflt')[0]!.metrics.ideographic,
    ).toBeUndefined();
    expect(ideoOf('base-no-dflt')).toBeUndefined();
  });

  it('is nothing at all for a BASE table carrying no ideo tag', () => {
    expect(
      readFaces(bytesOf('base-no-ideo'), 'base-no-ideo')[0]!.metrics.ideographic,
    ).toBeUndefined();
    expect(FIXTURE.fonts.find((f) => f.id === 'base-no-ideo')?.spec.base?.['DFLT']?.['hang']).toBe(
      610,
    );
  });

  it('is nothing at all for a font with no BASE table', () => {
    for (const id of ['plain', 'split', 'split-usetypo', 'astral']) {
      expect(readFaces(bytesOf(id), id)[0]!.metrics.ideographic, id).toBeUndefined();
    }
  });
});

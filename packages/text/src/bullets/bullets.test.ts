/**
 * Bullets, fields and script runs, re-derived from the T5 fixture.
 *
 * Every assertion reads `corpus/ground-truth/bullets.json` and works out what
 * the answer should be from the measurement. None was written by reading the
 * source: a test that re-derives the rule fails when the code is wrong, and a
 * test written from the code passes forever and proves nothing.
 *
 * The fixture holds three kinds of row, because the experiment used three
 * instruments and they fail differently:
 *
 * - **`rendered`** is what PowerPoint's own `ExtTextOutW` call was given, so a
 *   bullet's characters come back as characters. 5097 of them.
 * - **`indents`, `blips`, `wraps`** are positions in points from
 *   `Paragraphs(i).BoundLeft`, which is the only instrument for a picture
 *   bullet - it draws no text at all.
 * - **`fields`** are `TextRange.Text`, because four of the sixteen locales are
 *   shaped into glyph ids before drawing and the EMF cannot read those back.
 *
 * The adversarial half matters as much: `scores` records what every rival model
 * managed, and the tests below assert the rivals are *wrong* rather than only
 * that ours is right.
 */

import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/bullets.json' with { type: 'json' };

import {
  ALPHABETS,
  AUTONUMBER_SCHEMES,
  START_AT_MAX,
  UNMEASURED_SCHEMES,
  autonumberRule,
  autonumberTypeface,
  formatAutonumber,
  isAutonumberScheme,
} from './autonumber.js';
import {
  BLIP_BULLET_HEIGHT,
  blipBulletWidth,
  bulletLayout,
  drawableBullet,
  isStartAtInRange,
  isSymbolTypeface,
  numberParagraphs,
  symbolBulletChar,
  type NumberedParagraph,
  type ResolvedBullet,
} from './bullets.js';
import { TextError } from '../errors.js';
import {
  RESERVED_FIELD_TYPES,
  datePatternFor,
  isReservedFieldType,
  renderDatePattern,
  renderField,
} from '../fields/fields.js';
import { scriptSlotOf, splitScriptRuns, typefaceFor } from '../runs/script-runs.js';

/* -------------------------------------------------------------------------- */
/* the fixture's own encoding                                                 */
/* -------------------------------------------------------------------------- */

/**
 * Undo the run-length packing the fixture applies to long readings.
 *
 * `char*count` for four or more identical characters, which is what keeps the
 * count sweep - thirty `a`s and a full stop, over and over - inside the corpus
 * size cap. Mechanical and exactly reversible, and unpacked here rather than
 * trusted: the test would silently compare nothing if the encoding were wrong.
 */
function unpack(packed: string): string {
  let out = '';
  for (let at = 0; at < packed.length;) {
    const char = packed[at] ?? '';
    const star = packed[at + 1];
    if (star === '*') {
      let digits = '';
      let scan = at + 2;
      while (scan < packed.length && /[0-9]/.test(packed[scan] ?? '')) digits += packed[scan++];
      out += char.repeat(Number(digits));
      at = scan;
    } else {
      out += char;
      at++;
    }
  }
  return out;
}

interface RenderedRow {
  readonly scheme: string;
  readonly n: number;
  readonly drawn: string;
}

/** Every scheme reading, flattened out of the fixture's scheme/source/value nesting. */
const renderedRows: readonly RenderedRow[] = Object.entries(fixture.rendered).flatMap(
  ([scheme, sources]) =>
    Object.values(sources as Record<string, Record<string, string>>).flatMap((values) =>
      Object.entries(values).map(([n, drawn]) => ({ scheme, n: Number(n), drawn: unpack(drawn) })),
    ),
);

/** What a section's winning model scored, so a test can quote the measurement. */
function scoreOf(section: string, model: string): number {
  const scores = (fixture.scores as Record<string, { name: string; fits: number }[]>)[section];
  const found = scores?.find((entry) => entry.name === model);
  if (found === undefined) throw new Error(`no ${model} score in section ${section}`);
  return found.fits;
}

/* -------------------------------------------------------------------------- */
/* the vocabulary                                                             */
/* -------------------------------------------------------------------------- */

describe('the 41 autonumber schemes', () => {
  it('is the list PowerPoint itself wrote, in its own order', () => {
    expect([...AUTONUMBER_SCHEMES]).toEqual(fixture.schemes);
    expect(AUTONUMBER_SCHEMES).toHaveLength(41);
  });

  it('splits every scheme name a reading exists for', () => {
    for (const scheme of fixture.schemes) {
      if (UNMEASURED_SCHEMES.includes(scheme)) continue;
      const rule = autonumberRule(scheme);
      expect(rule.prefix + rule.numeral + rule.suffix).not.toBe('');
      expect(scheme.startsWith(rule.numeral)).toBe(true);
    }
  });

  it('refuses the two schemes PowerPoint shapes into glyph ids', () => {
    // The fixture names them: their EMF records hold indices into a font, so
    // the letters were never read and guessing them is not on offer.
    expect([...UNMEASURED_SCHEMES].sort()).toEqual([...fixture.glyphRunSchemes].sort());
    for (const scheme of UNMEASURED_SCHEMES) {
      expect(isAutonumberScheme(scheme)).toBe(true);
      expect(() => formatAutonumber(scheme, 1)).toThrow(TextError);
    }
  });

  it('recognises exactly those 41 and nothing else', () => {
    for (const scheme of fixture.schemes) expect(isAutonumberScheme(scheme)).toBe(true);
    for (const wrong of ['arabicPeriodd', 'alphaLc', 'ArabicPeriod', '']) {
      expect(isAutonumberScheme(wrong)).toBe(false);
    }
    expect(() => autonumberRule('notAScheme')).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* what each scheme renders                                                   */
/* -------------------------------------------------------------------------- */

describe('what a scheme renders', () => {
  it('reproduces every one of PowerPoint own readings', () => {
    expect(renderedRows.length).toBe(scoreOf('B. what a scheme renders', 'measured'));
    const wrong: string[] = [];
    for (const row of renderedRows) {
      const got = formatAutonumber(row.scheme, row.n);
      if (got !== row.drawn) wrong.push(`${row.scheme}@${String(row.n)}: ${got} != ${row.drawn}`);
    }
    expect(wrong.slice(0, 8)).toEqual([]);
    expect(renderedRows.length).toBeGreaterThan(5000);
  });

  // The adversarial half. Each of these is a reading somebody writes first, and
  // each is refuted by a row the fixture holds - so the assertion is that our
  // answer is the measured one and *not* the plausible one.
  it('is not bijective base-26: after z comes aa, then aaa', () => {
    // The bands are 26 wide and the letter is (n-1) mod 26, so 53 is the first
    // of the third band and its letter is `a` again. Bijective base-26 would
    // give "ba" at 53 and "aa" only at 27.
    expect(formatAutonumber('alphaLcPeriod', 26)).toBe('z.');
    expect(formatAutonumber('alphaLcPeriod', 27)).toBe('aa.');
    expect(formatAutonumber('alphaLcPeriod', 53)).toBe('aaa.');
    expect(formatAutonumber('alphaLcPeriod', 53)).not.toBe('ba.');
    expect(formatAutonumber('alphaLcPeriod', 54)).toBe('bbb.');
  });

  it('uses subtractive roman numerals, and just adds Ms past 3999', () => {
    expect(formatAutonumber('romanLcParenBoth', 4)).toBe('(iv)');
    expect(formatAutonumber('romanLcParenBoth', 4)).not.toBe('(iiii)');
    expect(formatAutonumber('romanUcPeriod', 3999)).toBe('MMMCMXCIX.');
    expect(formatAutonumber('romanUcPeriod', 4000)).toBe('MMMM.');
  });

  it('writes the East Asian zero as U+25CB and not U+3007', () => {
    const hundred = formatAutonumber('ea1ChsPlain', 100);
    expect(hundred).toBe('一○○');
    expect(hundred).not.toContain('〇');
  });

  it('gives the three East Asian families three different rules at 26', () => {
    expect(formatAutonumber('ea1ChsPlain', 26)).toBe('二十六');
    expect(formatAutonumber('ea1ChtPlain', 26)).toBe('二六');
    expect(formatAutonumber('ea1JpnKorPlain', 10)).toBe('一○');
    expect(formatAutonumber('ea1ChsPlain', 10)).toBe('十');
  });

  it('takes the fullwidth stop for the two double-byte schemes', () => {
    expect(formatAutonumber('arabicDbPeriod', 1)).toBe('１．');
    expect(formatAutonumber('ea1JpnChsDbPeriod', 1)).toBe('一．');
    expect(formatAutonumber('arabicPeriod', 1)).toBe('1.');
  });

  it('wraps the Wingdings circled numbers at ten and the Unicode ones at twenty', () => {
    expect(formatAutonumber('circleNumWdWhitePlain', 11)).toBe(
      formatAutonumber('circleNumWdWhitePlain', 1),
    );
    expect(formatAutonumber('circleNumDbPlain', 20)).toBe('⑳');
    expect(formatAutonumber('circleNumDbPlain', 21)).toBe('21');
  });

  it('wraps the alphabets at their measured points and not at a shared one', () => {
    expect(ALPHABETS['alphaLc']?.wrap).toBe(780);
    expect(ALPHABETS['hebrew2']?.wrap).toBe(392);
    // 781 is where the Latin counter restarts, measured one value at a time.
    expect(formatAutonumber('alphaLcPeriod', 781)).toBe('a.');
    expect(formatAutonumber('alphaLcPeriod', 780)).toBe(`${'z'.repeat(30)}.`);
  });

  it('gives Thai its own off-by-one and nobody else', () => {
    expect(ALPHABETS['thaiAlpha']?.bias).toBe(1);
    expect(ALPHABETS['alphaLc']?.bias).toBe(0);
    expect(ALPHABETS['hindiAlpha']?.bias).toBe(0);
    // Thai's bands begin at 42 and then at 41k; every other alphabet's at Ak+1.
    const thai = ALPHABETS['thaiAlpha']?.letters ?? '';
    expect(formatAutonumber('thaiAlphaPeriod', 41)).toBe(`${thai[40] ?? ''}.`);
    expect(formatAutonumber('thaiAlphaPeriod', 82)).toBe(`${(thai[40] ?? '').repeat(3)}.`);
  });

  it('fills Hebrew with the last letter rather than repeating the current one', () => {
    const hebrew = ALPHABETS['hebrew2']?.letters ?? '';
    const tav = hebrew[21] ?? '';
    const dalet = hebrew[3] ?? '';
    expect(formatAutonumber('hebrew2Minus', 26)).toBe(`${tav}${dalet}-`);
    expect(formatAutonumber('hebrew2Minus', 26)).not.toBe(`${dalet}${dalet}-`);
  });

  it('skips the three Thai consonants PowerPoint leaves out', () => {
    const thai = ALPHABETS['thaiAlpha']?.letters ?? '';
    expect(thai).toHaveLength(41);
    for (const skipped of ['ฃ', 'ฅ', 'ฆ']) expect(thai).not.toContain(skipped);
  });

  it('refuses a value outside ST_TextBulletStartAtNum, which PowerPoint repairs', () => {
    // The bound is read off the hostile packages rather than off the constant:
    // one probe per package, and PowerPoint repaired every one it refused. Using
    // `START_AT_MAX + 1` here would move with the constant under test and prove
    // nothing about where the boundary is.
    const repaired = new Set(fixture.repaired);
    expect(repaired.has('hostile-start-32768')).toBe(true);
    expect(repaired.has('hostile-start-0')).toBe(true);
    expect(repaired.has('hostile-start-neg1')).toBe(true);
    expect(repaired.has('hostile-start-32767')).toBe(false);

    expect(START_AT_MAX).toBe(32767);
    expect(isStartAtInRange(1)).toBe(true);
    expect(isStartAtInRange(32767)).toBe(true);
    for (const bad of [0, -1, 32768, 65536, 1.5]) {
      expect(isStartAtInRange(bad)).toBe(false);
      expect(() => formatAutonumber('arabicPeriod', bad)).toThrow(TextError);
    }
  });

  it('asks for Wingdings only where the glyphs exist nowhere else', () => {
    expect(autonumberTypeface('circleNumWdWhitePlain')).toBe('Wingdings');
    expect(autonumberTypeface('circleNumWdBlackPlain')).toBe('Wingdings');
    expect(autonumberTypeface('arabicPeriod')).toBeNull();
    expect(autonumberTypeface('circleNumDbPlain')).toBeNull();
  });
});

/* -------------------------------------------------------------------------- */
/* the numbering pass                                                         */
/* -------------------------------------------------------------------------- */

interface SeqRow {
  readonly id: string;
  readonly paras: readonly NumberedParagraph[];
  readonly numbers: readonly (number | null)[];
}

describe('the numbering pass', () => {
  const rows = fixture.sequences as unknown as readonly SeqRow[];

  it('reproduces every sequence PowerPoint laid out', () => {
    expect(rows.length).toBe(scoreOf('C. the numbering pass', 'measured'));
    for (const row of rows) {
      expect({ id: row.id, numbers: numberParagraphs(row.paras) }).toEqual({
        id: row.id,
        numbers: row.numbers,
      });
    }
  });

  it('lets a deeper level pass through and a shallower one break the run', () => {
    const outAndBack = rows.find((row) => row.id === 'seq-level-out-and-back');
    const twice = rows.find((row) => row.id === 'seq-level-twice');
    // The outer list resumes at three after a nested one; the inner restarts.
    expect(outAndBack?.numbers).toEqual([1, 2, 1, 2, 3, 4]);
    expect(twice?.numbers).toEqual([1, 1, 2, 1]);
  });

  it('starts a new run where startAt changes, so 7 is followed by 1', () => {
    expect(rows.find((row) => row.id === 'seq-startat-first')?.numbers).toEqual([7, 1, 2, 3]);
    expect(rows.find((row) => row.id === 'seq-startat-every')?.numbers).toEqual([7, 8, 9, 10]);
  });

  it('treats an absent startAt as exactly 1 rather than as its own value', () => {
    expect(rows.find((row) => row.id === 'seq-startat-zero-then')?.numbers).toEqual([1, 2, 3, 4]);
  });

  it('skips an unnumbered paragraph without consuming a number', () => {
    for (const id of ['seq-buNone-middle', 'seq-buChar-middle']) {
      expect(rows.find((row) => row.id === id)?.numbers).toEqual([1, 2, null, 1, 2]);
    }
  });

  it('counts each text body on its own', () => {
    expect(rows.find((row) => row.id === 'seq-second-body-a')?.numbers).toEqual([1, 2]);
    expect(rows.find((row) => row.id === 'seq-second-body-b')?.numbers).toEqual([1, 2]);
  });

  it('refuses to number past the end of ST_TextBulletStartAtNum', () => {
    const paras = [
      { level: 0, scheme: 'arabicPeriod', startAt: START_AT_MAX },
      { level: 0, scheme: 'arabicPeriod', startAt: START_AT_MAX },
    ];
    expect(() => numberParagraphs(paras)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* character bullets and the private-use mapping                              */
/* -------------------------------------------------------------------------- */

interface CharRow {
  readonly id: string;
  readonly char: string;
  readonly font: string | null;
  readonly drawn: string;
  readonly face: string | null;
}

describe('the private-use mapping for buChar', () => {
  const rows = fixture.characters as unknown as readonly CharRow[];

  it('reproduces every character PowerPoint drew', () => {
    expect(rows.length).toBe(scoreOf('D. the private-use mapping for buChar', 'measured'));
    for (const row of rows) {
      expect({ id: row.id, drawn: symbolBulletChar(row.char, row.font ?? undefined) }).toEqual({
        id: row.id,
        drawn: row.drawn,
      });
    }
  });

  it('goes through the ANSI codepage, so a bullet character is U+F095', () => {
    // The one that matters: U+2022 is byte 0x95 in Windows-1252 and nothing at
    // all in Latin-1, and it is the commonest bullet character there is.
    expect(symbolBulletChar('•', 'Wingdings')).toBe('');
    expect(symbolBulletChar('•', 'Wingdings')).not.toBe('');
  });

  it('leaves a text face alone', () => {
    expect(symbolBulletChar('§', 'Arial')).toBe('§');
    expect(symbolBulletChar('•', 'Arial')).toBe('•');
    expect(symbolBulletChar('§', undefined)).toBe('§');
  });

  it('knows the symbol faces the experiment measured, and claims no others', () => {
    for (const face of fixture.symbolFaces) expect(isSymbolTypeface(face)).toBe(true);
    for (const face of ['Arial', 'Calibri', 'Wingdings 4', 'wingdings']) {
      expect(isSymbolTypeface(face)).toBe(false);
    }
  });
});

/* -------------------------------------------------------------------------- */
/* face, size and colour                                                      */
/* -------------------------------------------------------------------------- */

interface FontRow {
  readonly id: string;
  readonly kind: 'autonum' | 'char';
  readonly buFont: string | null;
  readonly runFace: string;
  readonly face: string | null;
}

describe('which face draws the bullet', () => {
  const rows = fixture.fonts as unknown as readonly FontRow[];

  it('reproduces every face PowerPoint asked GDI for', () => {
    expect(rows.length).toBe(scoreOf('E. which face draws the bullet', 'measured'));
    const themes: Record<string, string> = { '+mj-lt': 'Georgia', '+mn-lt': 'Tahoma' };
    for (const row of rows) {
      const bullet: ResolvedBullet<string> =
        row.kind === 'autonum'
          ? { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8 }
          : {
              kind: 'char',
              char: '•',
              ...(row.buFont === null
                ? {}
                : row.buFont === 'tx'
                  ? { font: { kind: 'text' as const } }
                  : {
                      font: {
                        kind: 'typeface' as const,
                        value: themes[row.buFont] ?? row.buFont,
                      },
                    }),
            };
      const drawn = drawableBullet(bullet, { typeface: row.runFace, sz: 2400, color: 'k' }, 8);
      expect({ id: row.id, face: drawn.kind === 'text' ? drawn.typeface : null }).toEqual({
        id: row.id,
        face: row.face,
      });
    }
  });

  it('ignores buFont on an autonumber, which is what PowerPoint own UI writes', () => {
    const drawn = drawableBullet(
      {
        kind: 'autonum',
        scheme: 'arabicPeriod',
        startAt: 8,
        font: { kind: 'typeface', value: 'Wingdings' },
      },
      { typeface: 'Courier New', sz: 2400, color: 'k' },
      8,
    );
    expect(drawn.kind === 'text' && drawn.typeface).toBe('Courier New');
  });

  it('honours buFont on a character bullet', () => {
    const drawn = drawableBullet(
      { kind: 'char', char: '•', font: { kind: 'typeface', value: 'Wingdings' } },
      { typeface: 'Arial', sz: 2400, color: 'k' },
      undefined,
    );
    expect(drawn.kind === 'text' && drawn.typeface).toBe('Wingdings');
    expect(drawn.kind === 'text' && drawn.text).toBe('');
  });
});

interface SizeRow {
  readonly id: string;
  readonly runSz: number;
  readonly szPct: number | null;
  readonly szPts: number | null;
  readonly szTx: boolean;
  readonly pointsDrawn: number;
}

describe('what size the bullet is drawn at', () => {
  const rows = fixture.sizes as unknown as readonly SizeRow[];

  it('reproduces every size PowerPoint drew', () => {
    expect(rows.length).toBe(scoreOf('F. what size the bullet is drawn at', 'measured'));
    for (const row of rows) {
      const bullet: ResolvedBullet<string> = {
        kind: 'char',
        char: '•',
        ...(row.szTx
          ? { size: { kind: 'text' as const } }
          : row.szPct !== null
            ? { size: { kind: 'percent' as const, value: row.szPct } }
            : row.szPts !== null
              ? { size: { kind: 'points' as const, value: row.szPts } }
              : {}),
      };
      const drawn = drawableBullet(bullet, { typeface: 'Arial', sz: row.runSz, color: 'k' });
      // `pointsDrawn` is `lfHeight` divided by a calibration constant, so it
      // carries the integer quantisation of the logical unit - about an eighth
      // of a point. The analysis scored the same rows the same way.
      const pt = drawn.kind === 'text' ? drawn.sz / 100 : 0;
      expect(`${row.id}: ${String(pt)}`).toBe(
        Math.abs(pt - row.pointsDrawn) < 0.07 ? `${row.id}: ${String(pt)}` : `${row.id}: mismatch`,
      );
    }
  });

  it('takes a percentage of the first run size, not the largest', () => {
    const small = rows.find((row) => row.id === 'size-two-runs-small-first');
    const large = rows.find((row) => row.id === 'size-two-runs-large-first');
    expect(small?.pointsDrawn).toBeCloseTo(8, 1);
    expect(large?.pointsDrawn).toBeCloseTo(40, 1);
  });
});

interface ColourRow {
  readonly id: string;
  readonly buClr: string | null;
  readonly clrTx: boolean;
  readonly runClr: string;
  readonly drawnRgb: string;
}

describe('what colour the bullet is drawn in', () => {
  const rows = fixture.colours as unknown as readonly ColourRow[];

  it('reproduces every colour PowerPoint set', () => {
    expect(rows.length).toBe(scoreOf('G. what colour the bullet is drawn in', 'measured'));
    for (const row of rows) {
      const bullet: ResolvedBullet<string> = {
        kind: row.id.startsWith('clr-num') ? 'autonum' : 'char',
        ...(row.id.startsWith('clr-num') ? { scheme: 'arabicPeriod', startAt: 8 } : { char: '•' }),
        ...(row.clrTx
          ? { color: { kind: 'text' as const } }
          : row.buClr !== null
            ? { color: { kind: 'color' as const, value: row.buClr } }
            : {}),
      };
      const drawn = drawableBullet(bullet, { typeface: 'Arial', sz: 2400, color: row.runClr }, 8);
      expect({ id: row.id, rgb: drawn.kind === 'text' ? drawn.color : '' }).toEqual({
        id: row.id,
        rgb: row.drawnRgb,
      });
    }
  });

  it('lets buClr reach an autonumber, unlike buFont', () => {
    const numbered = rows.find((row) => row.id === 'clr-num-srgb');
    expect(numbered?.drawnRgb).toBe(numbered?.buClr);
    expect(numbered?.drawnRgb).not.toBe(numbered?.runClr);
  });
});

/* -------------------------------------------------------------------------- */
/* picture bullets and geometry                                               */
/* -------------------------------------------------------------------------- */

interface BlipRow {
  readonly id: string;
  readonly aspect: number;
  readonly sizePt: number;
  readonly szPct: number;
  readonly advancePt: number;
}

describe('what box a picture bullet occupies', () => {
  const rows = fixture.blips as unknown as readonly BlipRow[];

  it('reproduces every width PowerPoint reserved', () => {
    expect(rows.length).toBe(scoreOf('H. what box a picture bullet occupies', 'measured'));
    for (const row of rows) {
      const drawn = drawableBullet(
        {
          kind: 'blip',
          blip: 'rId2',
          ...(row.szPct === 100000 ? {} : { size: { kind: 'percent' as const, value: row.szPct } }),
        },
        { typeface: 'Arial', sz: row.sizePt * 100, color: 'k' },
      );
      const width = drawn.kind === 'picture' ? blipBulletWidth(drawn.sz, row.aspect) : -1;
      expect(Math.abs(width - row.advancePt)).toBeLessThan(0.05);
    }
  });

  it('scales to seven tenths of the font size and follows the aspect ratio', () => {
    expect(BLIP_BULLET_HEIGHT).toBe(fixture.blipHeightFactor);
    expect(blipBulletWidth(2400, 1)).toBeCloseTo(16.8, 6);
    expect(blipBulletWidth(2400, 4)).toBeCloseTo(67.2, 6);
    expect(blipBulletWidth(2400, 0.25)).toBeCloseTo(4.2, 6);
  });

  it('refuses an aspect ratio that is not one', () => {
    for (const bad of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => blipBulletWidth(2400, bad)).toThrow(TextError);
    }
  });
});

interface IndentRow {
  readonly id: string;
  readonly marL: number;
  readonly indent: number;
  readonly bulletPt: number;
  readonly textLeftPt: number;
}

describe('where the bullet and the text go', () => {
  const rows = fixture.indents as unknown as readonly IndentRow[];
  const wraps = fixture.wraps as unknown as readonly {
    id: string;
    marL: number;
    indent: number;
    secondLinePt: number;
  }[];

  it('reproduces every first-line position PowerPoint laid out', () => {
    expect(rows.length).toBe(
      scoreOf('I. where the text starts on a bulleted first line', 'measured'),
    );
    for (const row of rows) {
      const layout = bulletLayout(row.marL, row.indent, row.bulletPt);
      expect(Math.abs(layout.firstLineLeftPt - row.textLeftPt)).toBeLessThan(0.05);
    }
  });

  it('reproduces every wrapped line, which is marL and nothing else', () => {
    expect(wraps.length).toBe(scoreOf('J. where a wrapped line starts', 'measured'));
    for (const row of wraps) {
      const layout = bulletLayout(row.marL, row.indent, 28.75);
      expect(layout.wrappedLeftPt).toBeCloseTo(row.secondLinePt, 6);
    }
  });

  it('does not move the bullet for a positive indent', () => {
    // Eighteen probes at indent=+12 put the text exactly where indent=0 does.
    expect(bulletLayout(24, 12, 28.75).firstLineLeftPt).toBeCloseTo(
      bulletLayout(24, 0, 28.75).firstLineLeftPt,
      6,
    );
    expect(bulletLayout(24, 12, 28.75).bulletLeftPt).toBe(24);
  });

  it('clamps the bullet to the frame when the hanging indent is deeper than marL', () => {
    expect(bulletLayout(0, -36, 28.75).bulletLeftPt).toBe(0);
    // ...and the first line is held out to the hanging width, not to the bullet.
    expect(bulletLayout(0, -36, 28.75).firstLineLeftPt).toBe(36);
    expect(bulletLayout(0, -36, 57.5).firstLineLeftPt).toBeCloseTo(57.5, 6);
  });

  it('refuses lengths that are not lengths', () => {
    expect(() => bulletLayout(Number.NaN, 0, 10)).toThrow(TextError);
    expect(() => bulletLayout(0, 0, -1)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* script runs                                                                */
/* -------------------------------------------------------------------------- */

interface ScriptRow {
  readonly sample: string;
  readonly text: string;
  readonly expectSlot: string;
}

describe('which script slot draws a character', () => {
  const rows = fixture.scripts as unknown as readonly ScriptRow[];

  it('reproduces every slot PowerPoint chose', () => {
    expect(rows.length).toBe(scoreOf('K. which script slot draws a character', 'measured'));
    for (const row of rows) {
      for (const char of row.text) {
        expect({ sample: row.sample, slot: scriptSlotOf(char.codePointAt(0) ?? 0) }).toEqual({
          sample: row.sample,
          slot: row.expectSlot,
        });
      }
    }
  });

  it('sends Greek and Cyrillic to latin, not to cs', () => {
    expect(scriptSlotOf(0x0391)).toBe('latin');
    expect(scriptSlotOf(0x0414)).toBe('latin');
    expect(scriptSlotOf(0x05d0)).toBe('cs');
  });

  it('uses a:sym for the private-use area and nowhere else', () => {
    expect(scriptSlotOf(0xf0a7)).toBe('sym');
    expect(scriptSlotOf(0x2022)).toBe('latin');
  });

  it('falls back to latin for a slot the run does not state', () => {
    // Replayed from the decks that drop one slot at a time. Only the rows where
    // PowerPoint drew in a face the probe actually named are scored: elsewhere
    // Windows substituted, and a substitute is a property of the face pair
    // rather than of the rule. The decisive row is the private-use character
    // with `a:sym` dropped, which is drawn in the latin face directly.
    const declared: Record<string, string> = fixture.scriptFaces;
    const named = new Set(Object.values(declared));
    const rows = (
      fixture.scriptFallbacks as unknown as readonly {
        sample: string;
        text: string;
        dropped: string;
        slot: string;
        faces: readonly string[];
      }[]
    ).filter((row) => row.faces.every((face) => named.has(face)));
    expect(rows.length).toBeGreaterThan(40);
    for (const row of rows) {
      const faces: Record<string, string | undefined> = { ...declared };
      faces[row.dropped] = undefined;
      for (const char of row.text) {
        const used = typefaceFor(char.codePointAt(0) ?? 0, {
          latin: declared['latin'] ?? '',
          ea: faces['ea'],
          cs: faces['cs'],
          sym: faces['sym'],
        });
        expect(`${row.sample}/${row.dropped}: ${used}`).toBe(
          `${row.sample}/${row.dropped}: ${row.faces[0] ?? ''}`,
        );
      }
    }
  });

  it('does not try the other non-latin slots on the way', () => {
    // With ea absent, a Japanese character goes to latin and not to cs.
    expect(typefaceFor(0x3042, { latin: 'Georgia', cs: 'Courier New' })).toBe('Georgia');
    expect(typefaceFor(0x3042, { latin: 'Georgia', ea: 'MS Gothic' })).toBe('MS Gothic');
  });

  it('splits a mixed run into the stretches PowerPoint draws one at a time', () => {
    const faces = { latin: 'Georgia', ea: 'MS Gothic', cs: 'Courier New', sym: 'Wingdings' };
    const runs = splitScriptRuns('Azあいא', faces);
    expect(runs.map((run) => [run.text, run.typeface])).toEqual([
      ['Az', 'Georgia'],
      ['あい', 'MS Gothic'],
      ['א', 'Courier New'],
    ]);
  });

  it('keeps a surrogate pair whole', () => {
    const runs = splitScriptRuns('a\u{1f600}b', { latin: 'Georgia' });
    expect(runs.map((run) => run.text).join('')).toBe('a\u{1f600}b');
    expect(() => scriptSlotOf(0x110000)).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* fields                                                                     */
/* -------------------------------------------------------------------------- */

interface FieldRow {
  readonly id: string;
  readonly type: string;
  readonly lang: string;
  readonly cached: string;
  readonly rendered: string;
}

interface PatternRow {
  readonly lang: string;
  readonly type: string;
  readonly pattern: string;
  readonly rendered: string;
}

describe('fields', () => {
  const rows = fixture.fields as unknown as readonly FieldRow[];
  const patterns = fixture.datePatterns as unknown as readonly PatternRow[];

  it('knows the eighteen types PowerPoint recomputes', () => {
    expect(RESERVED_FIELD_TYPES).toHaveLength(18);
    for (const row of rows) expect(isReservedFieldType(row.type)).toBe(true);
    expect(isReservedFieldType('notAReservedType')).toBe(false);
  });

  it('never shows the cached text for a reserved type', () => {
    expect(rows.length).toBe(scoreOf('L. whether a field shows its cached text', 'measured'));
    for (const row of rows) expect(row.rendered).not.toBe(row.cached);
  });

  it('shows the cached text for a type nothing reserves', () => {
    expect(fixture.unknownFieldType).toBe('CACHED');
    const shown = renderField({
      type: 'notAReservedType',
      lang: 'en-US',
      cachedText: 'CACHED',
      now: new Date(2026, 8, 5, 12, 0, 0),
    });
    expect(shown).toEqual({
      text: 'CACHED',
      source: 'cached',
      reason: 'notAReservedType is not a reserved field type',
    });
  });

  it('renders slidenum from the slide position, not from the cache', () => {
    for (const row of fixture.slideNumbers) {
      expect(
        renderField({
          type: 'slidenum',
          lang: 'en-US',
          cachedText: '#stale#',
          slideNumber: row.slide,
          now: new Date(2026, 8, 5),
        }).text,
      ).toBe(row.rendered);
    }
    expect(() =>
      renderField({ type: 'slidenum', lang: 'en-US', cachedText: '', now: new Date() }),
    ).toThrow(TextError);
  });

  it('reproduces every measured pattern from the string it was derived from', () => {
    expect(patterns.length).toBeGreaterThan(150);
    for (const row of patterns) {
      expect(datePatternFor(row.type, row.lang)).toBe(row.pattern);
    }
  });

  it('keys the format on @lang, which the plan was right about', () => {
    // Eight distinct forms of the same type across sixteen locales.
    const forms = new Set(rows.filter((row) => row.type === 'datetime1').map((r) => r.rendered));
    expect(forms.size).toBeGreaterThan(4);
    expect(datePatternFor('datetime1', 'en-US')).toBe('M/d/yyyy');
    expect(datePatternFor('datetime1', 'de-DE')).toBe('dd.MM.yyyy');
  });

  it('renders a pattern the way PowerPoint rendered it', () => {
    const at = new Date(2026, 8, 5, 12, 56, 50);
    expect(renderDatePattern('M/d/yyyy', at, 'en-US')).toBe('9/5/2026');
    expect(renderDatePattern('dd.MM.yyyy', at, 'de-DE')).toBe('05.09.2026');
    expect(renderDatePattern('dddd, MMMM d, yyyy', at, 'en-US')).toBe(
      'Saturday, September 5, 2026',
    );
  });

  it('shows the cached text where the locale was never measured', () => {
    const shown = renderField({
      type: 'datetime2',
      lang: 'xx-XX',
      cachedText: 'last saved',
      now: new Date(2026, 8, 5),
    });
    expect(shown.source).toBe('cached');
    expect(shown.reason).toContain('xx-XX');
  });

  it('falls back from a region to its language and no further', () => {
    expect(datePatternFor('datetime1', 'en-US')).not.toBeNull();
    // en-US is measured; a bare "en" is not, and must not borrow the American form.
    expect(datePatternFor('datetime1', 'zz')).toBeNull();
  });
});

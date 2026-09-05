/**
 * The 41 `a:buAutoNum` schemes, measured against PowerPoint in experiment T5.
 *
 * ## Why this is measured and not derived
 *
 * ECMA-376 gives each value of `ST_TextAutonumberScheme` a one-line gloss -
 * "Lowercase alphabetical characters with a period" - and no sequence. What a
 * reader has to reproduce is the exact string, and the standard says nothing
 * about which letter follows `z`, which glyph stands for zero in the East Asian
 * schemes, or what happens at 900.
 *
 * PowerPoint answers all of it, because `Slide.Export(path, "EMF")` records the
 * `ExtTextOutW` call it makes to draw the bullet. So the fixture holds the
 * string itself for 5097 scheme-and-value pairs, and this file is the rule that
 * reproduces every one of them. `corpus/ground-truth/bullets.json`, and
 * `docs/adr/phase-3-text/0031-bullets-fields-and-script-runs.md`.
 *
 * ## Four things a plausible implementation gets wrong
 *
 * - **Alphabetic numbering is not bijective base-26.** After `z` comes `aa`,
 *   then `bb`, then `cc` - the letter is `(n-1) mod 26` and it is repeated
 *   `floor((n-1)/26)+1` times. Scored 1396 of 5097 for the bijective reading.
 * - **The East Asian zero is U+25CB WHITE CIRCLE**, not U+3007 IDEOGRAPHIC
 *   NUMBER ZERO. They are indistinguishable on screen and a table of Chinese
 *   numerals gives the wrong one.
 * - **The three East Asian families have three different rules.** Simplified
 *   keeps the ten-form to 99, Traditional only to 19, and Japanese/Korean not at
 *   all. So 26 is three characters in one, two in the next, and two different
 *   ones in the third.
 * - **The Wingdings circled numbers wrap at ten**; the Unicode circled digits
 *   run to twenty and then fall back to ASCII.
 */

import { TextError } from '../errors.js';

/**
 * `ST_TextAutonumberScheme`, in `PpNumberedBulletStyle` order.
 *
 * Not transcribed from the standard: PowerPoint's own object model was walked
 * from 0 to 47, styles 0 to 40 were accepted, and each one wrote its own
 * `a:buAutoNum/@type` into the saved package. `tools/ground-truth/text/bullets/analyse.ts` re-reads that
 * package and throws if this list disagrees.
 */
export const AUTONUMBER_SCHEMES = [
  'alphaLcPeriod',
  'alphaUcPeriod',
  'arabicParenR',
  'arabicPeriod',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'arabicParenBoth',
  'arabicPlain',
  'romanUcParenBoth',
  'romanUcParenR',
  'ea1ChsPlain',
  'ea1ChsPeriod',
  'circleNumDbPlain',
  'circleNumWdWhitePlain',
  'circleNumWdBlackPlain',
  'ea1ChtPlain',
  'ea1ChtPeriod',
  'arabic2Minus',
  'arabic1Minus',
  'hebrew2Minus',
  'ea1JpnKorPlain',
  'ea1JpnKorPeriod',
  'arabicDbPlain',
  'arabicDbPeriod',
  'thaiAlphaPeriod',
  'thaiAlphaParenR',
  'thaiAlphaParenBoth',
  'thaiNumPeriod',
  'thaiNumParenR',
  'thaiNumParenBoth',
  'hindiAlphaPeriod',
  'hindiNumPeriod',
  'ea1JpnChsDbPeriod',
  'hindiNumParenR',
  'hindiAlpha1Period',
] as const;

export type AutonumberScheme = (typeof AUTONUMBER_SCHEMES)[number];

const SCHEME_SET: ReadonlySet<string> = new Set(AUTONUMBER_SCHEMES);

/** Whether a string is one of the 41 values PowerPoint's own UI can produce. */
export function isAutonumberScheme(value: string): value is AutonumberScheme {
  return SCHEME_SET.has(value);
}

/**
 * `ST_TextBulletStartAtNum`, which the schema bounds at 1 to 32767.
 *
 * Both ends are load-bearing rather than decorative: a package written with
 * `startAt="0"`, `startAt="-1"`, `startAt="32768"` or `startAt="65536"` is one
 * PowerPoint **repairs** rather than reads, measured on four single-probe
 * packages, and `32767` opens clean.
 */
export const START_AT_MIN = 1;
export const START_AT_MAX = 32767;

/* -------------------------------------------------------------------------- */
/* alphabetic                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * An alphabet, and the three numbers measured with it.
 *
 * - **`wrap`** is where the counter restarts. 780 for five of the six, which is
 *   26 x 30 - a Latin number the Thai and Devanagari alphabets inherited, since
 *   their own maximum counts come out at 20, 49 and 23 rather than 30. Hebrew
 *   wraps at 392, which no arithmetic on 22 produces.
 * - **`bias`** is an off-by-one only Thai has: its bands begin at 42, 82, 123
 *   and 164 - that is 41k for every band after the second - while Latin's begin
 *   at 26k+1 and Devanagari's at 16k+1. Nothing in the shape of the rule
 *   predicts it, so it is a measured parameter rather than a formula.
 * - **`fillWithLast`** is Hebrew's alone: it writes the *last* letter of the
 *   alphabet for each repeat and then the current one, so 26 is tav-dalet where
 *   Latin would give dalet-dalet.
 *
 * All three were scored against 499 readings apiece; each alphabet has exactly
 * one combination that fits and the runner-up misses by ten to a hundred.
 */
export interface Alphabet {
  readonly letters: string;
  readonly wrap: number;
  readonly bias: number;
  readonly fillWithLast: boolean;
}

/**
 * The six alphabets, read off the sweep one value at a time.
 *
 * The Thai list is 41 letters, not the 44 consonants the script has: PowerPoint
 * skips U+0E03, U+0E05 and U+0E06. Nothing but walking n upwards shows that.
 */
export const ALPHABETS: Readonly<Record<string, Alphabet>> = {
  alphaLc: { letters: 'abcdefghijklmnopqrstuvwxyz', wrap: 780, bias: 0, fillWithLast: false },
  alphaUc: { letters: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', wrap: 780, bias: 0, fillWithLast: false },
  thaiAlpha: {
    letters: 'กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ',
    wrap: 780,
    bias: 1,
    fillWithLast: false,
  },
  hindiAlpha: {
    letters: 'अआइईउऊऋऌऍऎएऐऑऒओऔ',
    wrap: 780,
    bias: 0,
    fillWithLast: false,
  },
  hindiAlpha1: {
    letters: 'कखगघङचछजझञटठडढणतथदधनपफबभमयरलळवशषसह',
    wrap: 780,
    bias: 0,
    fillWithLast: false,
  },
  hebrew2: {
    letters: 'אבגדהוזחטיכלמנסעפצקרשת',
    wrap: 392,
    bias: 0,
    fillWithLast: true,
  },
};

function alphabetic(alphabet: Alphabet, n: number): string {
  const size = alphabet.letters.length;
  const reduced = ((n - 1) % alphabet.wrap) + 1;
  // The `<= size` guard is load-bearing only where the bias is one: without it
  // Thai's forty-first item would be two letters, and it is measured as one.
  const count = reduced <= size ? 1 : Math.floor((reduced - 1 + alphabet.bias) / size) + 1;
  const letter = alphabet.letters[(reduced - 1) % size] ?? '';
  const fill = alphabet.fillWithLast ? (alphabet.letters[size - 1] ?? '') : letter;
  return fill.repeat(count - 1) + letter;
}

/* -------------------------------------------------------------------------- */
/* roman                                                                      */
/* -------------------------------------------------------------------------- */

const ROMAN: readonly (readonly [number, string])[] = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

/**
 * Subtractive roman numerals, and nothing special above 3999.
 *
 * 4 is `iv` and not `iiii` - the additive reading scores 5037 of 5097 - and
 * 4000 is simply `mmmm`, with no overline and no switch to another notation.
 */
function roman(n: number): string {
  let left = n;
  let out = '';
  for (const [value, glyph] of ROMAN) {
    while (left >= value) {
      out += glyph;
      left -= value;
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* East Asian                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * The East Asian digits, and the one that is not the obvious character.
 *
 * Zero is U+25CB WHITE CIRCLE. U+3007 IDEOGRAPHIC NUMBER ZERO is what a table
 * of Chinese numerals gives, it is indistinguishable on screen in every font
 * that has both, and it is not what PowerPoint draws: `ea1ChsPlain` at 100 is
 * U+4E00 U+25CB U+25CB. Scored 5001 of 5097 for the ideographic reading.
 */
const CJK_DIGITS = '○一二三四五六七八九';
const CJK_TEN = '十';

const positional = (table: string, n: number): string =>
  String(n)
    .split('')
    .map((digit) => table[Number(digit)] ?? '')
    .join('');

/** Simplified Chinese keeps the ten-form all the way to 99: 26 is two-ten-six. */
function cjkSimplified(n: number): string {
  if (n < 10) return CJK_DIGITS[n] ?? '';
  if (n < 100) {
    const tens = Math.floor(n / 10);
    const ones = n % 10;
    return (
      (tens === 1 ? '' : (CJK_DIGITS[tens] ?? '')) +
      CJK_TEN +
      (ones === 0 ? '' : (CJK_DIGITS[ones] ?? ''))
    );
  }
  return positional(CJK_DIGITS, n);
}

/** Traditional Chinese uses the ten-form only for 10 to 19; 20 is two-circle. */
function cjkTraditional(n: number): string {
  if (n < 10) return CJK_DIGITS[n] ?? '';
  if (n < 20) return CJK_TEN + (n % 10 === 0 ? '' : (CJK_DIGITS[n % 10] ?? ''));
  return positional(CJK_DIGITS, n);
}

/* -------------------------------------------------------------------------- */
/* circled                                                                    */
/* -------------------------------------------------------------------------- */

/** Circled digits U+2460..U+2473 cover 1 to 20; above that PowerPoint gives up. */
const circledUnicode = (n: number): string =>
  n >= 1 && n <= 20 ? String.fromCodePoint(0x245f + n) : String(n);

/**
 * Wingdings circled numbers: white at U+F081, black at U+F08C, ten of each.
 *
 * Above ten they **wrap** rather than falling back to digits, so 11 draws the
 * same glyph as 1 and 4000 the same as 10. That is the opposite of what the
 * Unicode circled digits do, and both were probed to 4000 to find out.
 *
 * The code points are in the private-use area because these are symbol-font
 * glyphs: a renderer must ask for them in Wingdings, and nothing else.
 */
const circledWingdings = (n: number, black: boolean): string =>
  String.fromCodePoint((black ? 0xf08c : 0xf081) + ((n - 1) % 10));

/* -------------------------------------------------------------------------- */
/* the table                                                                  */
/* -------------------------------------------------------------------------- */

const FULLWIDTH_DIGITS = '０１２３４５６７８９';
const THAI_DIGITS = '๐๑๒๓๔๕๖๗๘๙';
const HINDI_DIGITS = '०१२३४५६७८९';

const NUMERALS: Readonly<Record<string, (n: number) => string>> = {
  arabic: (n) => String(n),
  arabicDb: (n) => positional(FULLWIDTH_DIGITS, n),
  alphaLc: (n) => alphabetic(ALPHABETS['alphaLc'] as Alphabet, n),
  alphaUc: (n) => alphabetic(ALPHABETS['alphaUc'] as Alphabet, n),
  romanLc: roman,
  romanUc: (n) => roman(n).toUpperCase(),
  ea1Chs: cjkSimplified,
  ea1Cht: cjkTraditional,
  // Japanese and Korean are positional throughout: ten is one-circle, not the
  // ten glyph. `ea1JpnChsDb` follows them rather than the Chs in its own name.
  ea1JpnKor: (n) => positional(CJK_DIGITS, n),
  ea1JpnChsDb: (n) => positional(CJK_DIGITS, n),
  circleNumDb: circledUnicode,
  circleNumWdWhite: (n) => circledWingdings(n, false),
  circleNumWdBlack: (n) => circledWingdings(n, true),
  thaiAlpha: (n) => alphabetic(ALPHABETS['thaiAlpha'] as Alphabet, n),
  thaiNum: (n) => positional(THAI_DIGITS, n),
  hindiAlpha: (n) => alphabetic(ALPHABETS['hindiAlpha'] as Alphabet, n),
  hindiAlpha1: (n) => alphabetic(ALPHABETS['hindiAlpha1'] as Alphabet, n),
  hindiNum: (n) => positional(HINDI_DIGITS, n),
  hebrew2: (n) => alphabetic(ALPHABETS['hebrew2'] as Alphabet, n),
};

/**
 * The full stop the double-byte schemes take, which is not the ASCII one.
 *
 * U+FF0E FULLWIDTH FULL STOP. `arabicDbPeriod` and `ea1JpnChsDbPeriod` are the
 * only two, and concatenating `"."` instead is wrong on every Japanese and
 * Chinese numbered list: scored 5025 of 5097.
 */
const FULLWIDTH_PERIOD = '．';

/** The decoration each scheme name ends in, longest suffix first. */
const DECORATIONS: readonly (readonly [string, string, string])[] = [
  ['ParenBoth', '(', ')'],
  ['ParenR', '', ')'],
  ['Period', '', '.'],
  ['Plain', '', ''],
  ['Minus', '', '-'],
];

/** A scheme split into its numeral system and the characters around it. */
export interface AutonumberRule {
  readonly numeral: string;
  readonly prefix: string;
  readonly suffix: string;
}

/**
 * The two of the 41 whose rendering T5 could not read.
 *
 * `arabic1Minus` and `arabic2Minus` are the Arabic abjad and alphabet, and
 * PowerPoint **shapes them before drawing**: it hands GDI a run of glyph ids
 * rather than characters, so the EMF - which is what makes every other scheme a
 * reading rather than a fit - holds indices into the face's `glyf` table and no
 * text at all. The repeat structure is still visible in them, and the fixture
 * keeps them under `glyphRuns` for that, but the letters are not.
 *
 * Guessing them from the Unicode block would produce a sequence that renders
 * plausibly and has never been checked against anything, which is the one thing
 * this project's rules forbid outright. So they throw, and closing the gap needs
 * the `cmap` reversal that 8.1's SFNT reader will make possible.
 */
export const UNMEASURED_SCHEMES: readonly string[] = ['arabic1Minus', 'arabic2Minus'];

/**
 * Split a scheme name into a numeral system and a decoration.
 *
 * The scheme names are compositional and the split is exact - 41 names, five
 * decorations, nineteen numeral systems - which is why this is a function rather
 * than a 41-row table nobody could check.
 */
export function autonumberRule(scheme: string): AutonumberRule {
  for (const [tail, prefix, suffix] of DECORATIONS) {
    if (!scheme.endsWith(tail)) continue;
    const numeral = scheme.slice(0, scheme.length - tail.length);
    if (NUMERALS[numeral] === undefined) continue;
    const wide = numeral === 'arabicDb' || numeral === 'ea1JpnChsDb';
    return { numeral, prefix, suffix: wide && suffix === '.' ? FULLWIDTH_PERIOD : suffix };
  }
  if (UNMEASURED_SCHEMES.includes(scheme)) {
    throw new TextError(
      'TEXT_AUTONUMBER_UNMEASURED',
      `${scheme} is a real scheme whose rendering T5 could not read: PowerPoint shapes it into ` +
        'glyph ids before drawing, so the measurement holds indices rather than characters',
      scheme,
    );
  }
  throw new TextError('TEXT_AUTONUMBER_SCHEME', `${scheme} is not an autonumber scheme`, scheme);
}

/**
 * The string PowerPoint draws for `scheme` at `n`.
 *
 * `n` is the bullet's number, which is `startAt` plus its position in the run -
 * see `numberParagraphs` in `bullets.ts`. It is not clamped here: a caller with
 * a value outside `ST_TextBulletStartAtNum` has read a package PowerPoint would
 * have repaired, and that is worth an error rather than a plausible string.
 */
export function formatAutonumber(scheme: string, n: number): string {
  if (!Number.isInteger(n) || n < START_AT_MIN || n > START_AT_MAX) {
    throw new TextError(
      'TEXT_AUTONUMBER_VALUE',
      `bullet number ${String(n)} is outside ST_TextBulletStartAtNum (${String(START_AT_MIN)}..${String(START_AT_MAX)})`,
      String(n),
    );
  }
  const rule = autonumberRule(scheme);
  const numeral = NUMERALS[rule.numeral];
  if (numeral === undefined) {
    throw new TextError('TEXT_AUTONUMBER_SCHEME', `no numeral system for ${rule.numeral}`, scheme);
  }
  return rule.prefix + numeral(n) + rule.suffix;
}

/**
 * The typeface a scheme needs, where it needs one.
 *
 * Two of the 41 draw private-use glyphs that exist only in Wingdings, so a
 * renderer that hands them to the run's face draws two `.notdef` boxes. Every
 * other scheme is drawn in the run's own face - `a:buFont` is measured **inert**
 * for an autonumber - so this returns `null` for them and the caller does not
 * have to know which is which.
 */
export function autonumberTypeface(scheme: string): string | null {
  const { numeral } = autonumberRule(scheme);
  return numeral === 'circleNumWdWhite' || numeral === 'circleNumWdBlack' ? 'Wingdings' : null;
}

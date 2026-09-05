/**
 * Experiment T3 - where does PowerPoint break a line?
 *
 * The probe table. `build-break-deck.ts` turns it into packages,
 * `read-breaks.ps1` reads what PowerPoint laid out, and `analyse-breaks.ts`
 * scores the candidate models against the result.
 *
 * ## The method
 *
 * `TextRange2.Lines(i, 1).Text` returns the *text* of line i. So unlike T2,
 * which had to infer a rule from a float, this experiment reads the answer
 * directly: the string on line 1 **is** the break position.
 *
 * That turns one string into a complete measurement, by way of a width sweep.
 * A greedy line breaker sets line 1 to the longest prefix that both fits the
 * box and ends at a break opportunity. So if the box is made exactly wide
 * enough for `i` code points and no more, line 1 comes back as
 *
 *     the largest break opportunity at or before i
 *
 * and sweeping `i` from 1 to the end of the string enumerates the whole
 * opportunity set, with one shape per `i` and nothing left to infer. The
 * observed lengths are a monotone non-decreasing step function of `i`, and the
 * set of distinct values it takes is the answer.
 *
 * Two things make this falsifiable rather than merely convenient:
 *
 * 1. **The steps are the data.** A model that permits a break the sweep never
 *    produced is refuted by that absence, not just unsupported by it. There is
 *    no "the corpus happened not to exercise it" escape.
 * 2. **A too-long line 1 is a finding, not an error.** If line 1 comes back
 *    *longer* than `i` code points, PowerPoint's fit test is measuring
 *    something narrower than the prefix - which is exactly the question the
 *    trailing-space and `hangingPunct` probes exist to ask. The analysis
 *    separates that from a mis-sized box using the recorded margin.
 *
 * ## Where the box widths come from
 *
 * `measure-break-widths.ts` measures every prefix of every probe string in
 * Chromium, and each cut probe's box is placed at the **midpoint** between the
 * prefix of length `i` and the prefix of length `i + 1`. T2 put Chromium and
 * PowerPoint a median of 0.15% apart (p95 0.70%), and half a character advance
 * at 24pt is several points of margin, so the midpoint is safe by two orders of
 * magnitude on every probe but the pathological ones - and the margin is
 * recorded per probe so the analysis can say which those were instead of
 * guessing.
 *
 * Every string also gets a `wrap="none"` control shape. Its `BoundWidth` is
 * PowerPoint's own answer for the whole string, so the Chromium estimate is
 * checked against the authority rather than trusted.
 *
 * ## Why the strings are written as escapes
 *
 * `\u3002` and `\uff61` are both "a full stop" and are different characters
 * with different line-break classes. A reader cannot tell them apart in a
 * source file, and neither can a diff. Every non-ASCII probe character is an
 * escape with the character named beside it.
 */

/** Which face a probe string is set in. Resolved to a typeface in the builder. */
export type FaceKey = 'latin' | 'mono' | 'ea' | 'thai' | 'arabic';

/**
 * The typeface each face key names.
 *
 * All three of `a:latin`, `a:ea` and `a:cs` are set to the same name on every
 * probe, so that no probe can accidentally measure PowerPoint's script splitter
 * instead of its line breaker. That is 3.5's question, and giving it a second
 * answer here would make both harder to read.
 */
export const FACES: Readonly<Record<FaceKey, string>> = {
  latin: 'Arial',
  mono: 'Courier New',
  // Ships with Windows 10 and 11 and covers Latin as well as kana and kanji, so
  // the mixed-script probes stay in one face and one measurement.
  ea: 'Yu Gothic',
  thai: 'Leelawadee UI',
  arabic: 'Arial',
};

/**
 * What a probe string is asking about. Only used to group the report - the
 * scoring never reads it, so a mislabelled string cannot flatter a model.
 */
export type Category =
  | 'space'
  | 'hyphen'
  | 'number'
  | 'punct'
  | 'invisible'
  | 'ea'
  | 'script'
  | 'uri'
  | 'unbreakable'
  | 'kinsoku';

/**
 * The three `a:pPr` attributes that tailor line breaking, as the file states
 * them. `undefined` means the attribute is absent, which is a distinct case
 * from either boolean: ECMA-376 gives `latinLnBrk` a default of `true` and the
 * plan claims PowerPoint behaves as `false`, so "absent" is the case that
 * settles it.
 */
export interface BreakFlags {
  readonly latinLnBrk?: boolean | undefined;
  readonly eaLnBrk?: boolean | undefined;
  readonly hangingPunct?: boolean | undefined;
}

export interface BreakString {
  readonly id: string;
  readonly category: Category;
  readonly face: FaceKey;
  /** The probe text. Written with `\u` escapes for everything outside ASCII. */
  readonly text: string;
  /** What this string discriminates, and what the wrong answer would look like. */
  readonly note: string;
  /**
   * Paragraph tailoring flags. A string listing more than one variant is
   * emitted once per variant, with the variant index in the probe id.
   */
  readonly flags?: readonly BreakFlags[] | undefined;
}

/** 24pt: big enough that half a character advance is several points of margin. */
export const PROBE_SIZE = 2400;

// --------------------------------------------------------------- named chars
// Every non-ASCII probe character, named once, so the strings below read as
// prose and the escapes cannot drift from what they are meant to be.

const NBSP = '\u00a0'; // NO-BREAK SPACE, class GL
const SHY = '\u00ad'; // SOFT HYPHEN, class BA
const NB_HYPHEN = '\u2011'; // NON-BREAKING HYPHEN, class GL
const EN_DASH = '\u2013'; // EN DASH, class BA
const EM_DASH = '\u2014'; // EM DASH, class B2
const LSQUO = '\u2018'; // LEFT SINGLE QUOTATION MARK, class QU
const RSQUO = '\u2019'; // RIGHT SINGLE QUOTATION MARK, class QU
const LDQUO = '\u201c'; // LEFT DOUBLE QUOTATION MARK, class QU
const RDQUO = '\u201d'; // RIGHT DOUBLE QUOTATION MARK, class QU
const ZWSP = '\u200b'; // ZERO WIDTH SPACE, class ZW
const WJ = '\u2060'; // WORD JOINER, class WJ

// Japanese. The kinsoku sets in `p:kinsoku` name these by code point, so the
// probes have to use exactly the ones the sets contain.
const IDEO_COMMA = '\u3001'; // IDEOGRAPHIC COMMA, invalStChars
const IDEO_STOP = '\u3002'; // IDEOGRAPHIC FULL STOP, invalStChars
const L_CORNER = '\u300c'; // LEFT CORNER BRACKET, invalEndChars
const R_CORNER = '\u300d'; // RIGHT CORNER BRACKET, invalStChars
const PROLONG = '\u30fc'; // KATAKANA-HIRAGANA PROLONGED SOUND MARK, invalStChars
const SMALL_YA = '\u3083'; // HIRAGANA LETTER SMALL YA, invalStChars
const SMALL_TSU = '\u3063'; // HIRAGANA LETTER SMALL TU, invalStChars

/** `\u65e5\u672c\u8a9e` - "Japanese". Three kanji, all class ID. */
const NIHONGO = '\u65e5\u672c\u8a9e';
/** `\u30c6\u30b9\u30c8` - "test". Three katakana. */
const TESUTO = '\u30c6\u30b9\u30c8';
/** `\u3067\u3059` - the copula. Two hiragana. */
const DESU = '\u3067\u3059';
/** `\u3053\u3093\u306b\u3061\u306f` - "hello". Five hiragana. */
const KONNICHIWA = '\u3053\u3093\u306b\u3061\u306f';

/** `\u0e2a\u0e27\u0e31\u0e2a\u0e14\u0e35\u0e0a\u0e32\u0e27\u0e42\u0e25\u0e01` - "hello world" in Thai, with no spaces. */
const THAI_HELLO = '\u0e2a\u0e27\u0e31\u0e2a\u0e14\u0e35\u0e0a\u0e32\u0e27\u0e42\u0e25\u0e01';
/** `\u0645\u0631\u062d\u0628\u0627 \u0628\u0627\u0644\u0639\u0627\u0644\u0645` - "hello world" in Arabic. */
const ARABIC_HELLO = '\u0645\u0631\u062d\u0628\u0627 \u0628\u0627\u0644\u0639\u0627\u0644\u0645';

// ------------------------------------------------------------- the strings

/**
 * The probe strings.
 *
 * Each is short - at most 28 code points - for two reasons. The sweep costs one
 * shape per code point, and the box is placed from a Chromium estimate whose
 * absolute error grows with the length of the prefix being estimated. A short
 * string keeps both small.
 *
 * ## Why nearly every string starts with a short word and a space
 *
 * The first pass of T3 did not, and it made a whole class of probe say nothing.
 * When a box is too narrow for any opportunity, PowerPoint breaks at exactly as
 * many code points as fit - and a string whose *only* candidate opportunities
 * are the ones under test then reads `line 1 = fitMax` at every width, which is
 * equally consistent with "every position breaks" and "no position breaks".
 * `alpha/beta/gamma` measured as an unbroken staircase and proved nothing about
 * the solidus.
 *
 * A guaranteed opportunity near the start fixes it. Once a real break exists at
 * position 3, every later reading of `line 1 = 3` proves that *nothing* in
 * between is an opportunity - the greedy rule would have preferred it. The
 * question turns from "is this consistent with" into "is this refuted by", which
 * is the only form worth measuring.
 *
 * `aa`/`bbb`/`cccc` fillers rather than real words: they have no kern pairs in
 * any probed face, so a break position cannot be confused with a kerning
 * difference, and a filler cannot accidentally contain the opportunity the probe
 * is asking about.
 */
export const STRINGS: readonly BreakString[] = [
  // ---------------------------------------------------------------- spaces
  {
    id: 'sp-words',
    category: 'space',
    face: 'latin',
    text: 'alpha beta gamma delta',
    note: 'The base case. If the opportunity set is not exactly the three positions after the three spaces, nothing below is worth reading.',
  },
  {
    id: 'sp-trail',
    category: 'space',
    face: 'latin',
    text: 'aaaa bbbb',
    note: 'The trailing-space discriminator. If the space is excluded from the fit test, line 1 comes back as five code points at a box that only fits four.',
  },
  {
    id: 'sp-double',
    category: 'space',
    face: 'latin',
    text: 'aaaa  bbbb  cccc',
    note: 'Two spaces. A break between them would split the run; the opportunity should be after the second, with both hanging.',
  },
  {
    id: 'sp-triple',
    category: 'space',
    face: 'latin',
    text: 'aa   bb   cc',
    note: 'Three spaces, to show that whatever happens with two is a rule about runs and not a special case for pairs.',
  },
  {
    id: 'sp-lead',
    category: 'space',
    face: 'latin',
    text: ' aaaa bbbb',
    note: 'A leading space. Whether it is measured at all decides where every line of a centred paragraph sits.',
  },
  {
    id: 'sp-mono',
    category: 'space',
    face: 'mono',
    text: 'alpha beta gamma',
    note: 'The same question in a monospaced face, where every prefix width is exact by construction and a mis-sized box cannot hide inside proportional advances.',
  },

  // ------------------------------------------------------- hyphens and dashes
  {
    id: 'hy-hyphen',
    category: 'hyphen',
    face: 'latin',
    text: 'aa well-known-thing',
    note: 'HY: break after the hyphen, never before it.',
  },
  {
    id: 'hy-nobreak',
    category: 'hyphen',
    face: 'latin',
    text: `aa well${NB_HYPHEN}known`,
    note: 'U+2011 NON-BREAKING HYPHEN. Looks identical to U+002D on screen and must produce no opportunity at all - and must go on suppressing one when latinLnBrk turns character breaking on.',
    flags: [{}, { latinLnBrk: true }],
  },
  {
    id: 'hy-endash',
    category: 'hyphen',
    face: 'latin',
    text: `aa alpha${EN_DASH}beta`,
    note: 'EN DASH: break after, not before.',
  },
  {
    id: 'hy-emdash',
    category: 'hyphen',
    face: 'latin',
    text: `aa alpha${EM_DASH}beta`,
    note: 'EM DASH is the one class UAX#14 lets a line break on *both* sides of, so the position before it and the position after it are two separate questions and this string asks both.',
  },
  {
    id: 'hy-slash',
    category: 'hyphen',
    face: 'latin',
    text: 'aa alpha/beta',
    note: 'SOLIDUS: no break before it is uncontroversial; whether one is permitted after it is where implementations differ.',
  },
  {
    id: 'hy-datenum',
    category: 'hyphen',
    face: 'latin',
    text: 'aa 2026-09-04',
    note: 'A hyphen between digits. The number rules fold some punctuation into the number and forbid the break that hy-hyphen permits - the same character, two answers, decided by its neighbours.',
  },

  // ---------------------------------------------------------------- numbers
  {
    id: 'nu-thousands',
    category: 'number',
    face: 'latin',
    text: 'aa 1,234,567',
    note: 'A comma between digits. A breaker that treats it as ordinary punctuation splits every large number in every financial deck.',
  },
  {
    id: 'nu-decimal',
    category: 'number',
    face: 'latin',
    text: 'aa 3.14159',
    note: 'The same for a decimal point.',
  },
  {
    id: 'nu-currency',
    category: 'number',
    face: 'latin',
    text: 'aa $1,000.00',
    note: 'A currency sign before a number: no break after it.',
  },
  {
    id: 'nu-percent',
    category: 'number',
    face: 'latin',
    text: 'aa 1,000%',
    note: 'A percent sign after a number: no break before it.',
  },
  {
    id: 'nu-spaced',
    category: 'number',
    face: 'latin',
    text: 'aa 1 000 000',
    note: 'The control for the three above: spaces inside a number are ordinary opportunities, so the number rule is about the characters and not about digits generally.',
  },

  // ------------------------------------------------------------ punctuation
  {
    id: 'pu-paren',
    category: 'punct',
    face: 'latin',
    text: 'aa (bbb) cc',
    note: 'No break after an opening bracket, none before a closing one.',
  },
  {
    id: 'pu-bracket',
    category: 'punct',
    face: 'latin',
    text: 'aa [bbb] cc',
    note: 'The same for square brackets, which some tables class differently.',
  },
  {
    id: 'pu-comma',
    category: 'punct',
    face: 'latin',
    text: 'aa bbb, ccc',
    note: 'No break before the comma; the opportunity is after the space that follows it. Repeated across hangingPunct so a Latin comma can be compared with an ideographic one.',
    flags: [{}, { hangingPunct: true }, { hangingPunct: false }],
  },
  {
    id: 'pu-stop',
    category: 'punct',
    face: 'latin',
    text: 'aa bbb. ccc',
    note: 'The same for a full stop, which is one character with two jobs - sentence and decimal point.',
  },
  {
    id: 'pu-excl',
    category: 'punct',
    face: 'latin',
    text: 'aa bbb! ccc',
    note: 'No break before an exclamation mark.',
  },
  {
    id: 'pu-colon',
    category: 'punct',
    face: 'latin',
    text: 'aa bbb: ccc',
    note: 'The two characters most often mis-classed as ordinary punctuation.',
  },
  {
    id: 'pu-quote',
    category: 'punct',
    face: 'latin',
    text: `aa ${LSQUO}bbb${RSQUO} cc`,
    note: 'Quotation marks are ambiguous by design - the same character opens and closes - so a table that resolves them by position and one that forbids both sides disagree here.',
  },
  {
    id: 'pu-dquote',
    category: 'punct',
    face: 'latin',
    text: `aa ${LDQUO}bbb${RDQUO} cc`,
    note: 'The double-quote form of the same question.',
  },

  // -------------------------------------------------- invisible characters
  {
    id: 'iv-zwsp',
    category: 'invisible',
    face: 'latin',
    text: `aa bbbb${ZWSP}cccc`,
    note: 'ZERO WIDTH SPACE: the canonical "break here, invisibly". An opportunity with no advance, so the break position and the width are independent - and the space at 3 is what makes its absence provable.',
  },
  {
    id: 'iv-wj',
    category: 'invisible',
    face: 'latin',
    text: `aa bbbb${WJ}cccc`,
    note: 'WORD JOINER: must suppress an opportunity where there is one to suppress, and must go on suppressing the character-level break when latinLnBrk turns it on.',
    flags: [{}, { latinLnBrk: true }],
  },
  {
    id: 'iv-nbsp',
    category: 'invisible',
    face: 'latin',
    text: `aa bbbb${NBSP}cccc`,
    note: 'NO-BREAK SPACE: indistinguishable from a space on screen, and must produce no opportunity. The classic silently-wrong wrap. The latinLnBrk variant asks the harder half - whether it also suppresses the character-level break, the way the non-breaking hyphen does - which the first pass left as an unmeasured generalisation and a surviving mutant.',
    flags: [{}, { latinLnBrk: true }],
  },
  {
    id: 'iv-shy',
    category: 'invisible',
    face: 'latin',
    text: `aa bbbb${SHY}cccc`,
    note: 'SOFT HYPHEN: an opportunity that costs a visible hyphen when taken and nothing when not. Whether PowerPoint takes it at all is measured, not assumed.',
  },

  // ------------------------------------------------------------- East Asian
  {
    id: 'ea-plain',
    category: 'ea',
    face: 'ea',
    text: `${NIHONGO}\u306e${TESUTO}${DESU}`,
    note: 'Ideographs break between any two. Repeated across eaLnBrk, which is supposed to be the attribute that stops them.',
    flags: [{}, { eaLnBrk: true }, { eaLnBrk: false }],
  },
  {
    id: 'ea-kinsoku-start',
    category: 'ea',
    face: 'ea',
    text: `${NIHONGO}${IDEO_COMMA}${TESUTO}${IDEO_STOP}${DESU}`,
    note: 'The comma and the full stop may not start a line, so the opportunity before each must be absent while every other position remains. Repeated across hangingPunct, which decides whether the forbidden character hangs outside the measure or forces a fallback.',
    flags: [{}, { hangingPunct: true }, { hangingPunct: false }],
  },
  {
    id: 'ea-kinsoku-end',
    category: 'ea',
    face: 'ea',
    text: `${NIHONGO}${L_CORNER}${TESUTO}${R_CORNER}${DESU}`,
    note: 'The opening bracket may not end a line and the closing one may not start one - the two directions of the filter, in one string, across hangingPunct.',
    flags: [{}, { hangingPunct: true }, { hangingPunct: false }],
  },
  {
    id: 'ea-small-kana',
    category: 'ea',
    face: 'ea',
    text: `${KONNICHIWA}${SMALL_TSU}\u3068${SMALL_YA}\u3093`,
    note: 'Small kana are in the strict start-of-line set and in no Unicode class that would exclude them, so they separate a kinsoku filter from UAX#14 alone.',
  },
  {
    id: 'ea-prolong',
    category: 'ea',
    face: 'ea',
    text: `\u30c7\u30fc\u30bf${PROLONG}\u30d9\u30fc\u30b9`,
    note: 'The prolonged sound mark is class CJ in Unicode and in the strict start set in Japanese practice - the two disagree, and this says which PowerPoint follows.',
  },

  // ------------------------------------------------------------ mixed script
  {
    id: 'sc-mixed',
    category: 'script',
    face: 'ea',
    text: `${NIHONGO}abcdef${DESU}`,
    note: 'A Latin word inside Japanese: an opportunity at each transition and none inside the Latin word. Also the probe that proves an ideograph boundary is a real opportunity rather than a forced break, because positions after it fall back to it.',
  },
  {
    id: 'sc-cjk-punct',
    category: 'script',
    face: 'ea',
    text: `aa ${IDEO_COMMA}bb${IDEO_STOP}cc`,
    note: 'An ideographic comma and full stop with no ideograph anywhere near them. Every other East Asian probe has one as a neighbour, so the classifier has only ever been asked about these characters in company - and a mutation that stopped classifying CJK punctuation as East Asian survived the whole suite.',
  },
  {
    id: 'sc-fullwidth',
    category: 'script',
    face: 'ea',
    text: 'aa ＡＢＣbb',
    note: 'FULLWIDTH LATIN CAPITAL A, B, C. Unicode calls them Script=Latin, so nothing but an explicit range makes them East Asian - and whether they break between themselves is the question that range is answering.',
  },
  {
    id: 'sc-halfwidth-kana',
    category: 'script',
    face: 'ea',
    text: 'aa ｱｲｳbb',
    note: 'HALFWIDTH KATAKANA A, I, U. Script=Katakana, so the script test alone should reach them; this is the control that says whether the explicit range is doing the work or the script property is.',
  },
  {
    id: 'sc-thai',
    category: 'script',
    face: 'thai',
    text: THAI_HELLO,
    note: 'Thai has no spaces and needs dictionary segmentation. Whatever PowerPoint does here is the honest scope boundary, and the plan already says we may over-run.',
  },
  {
    id: 'sc-arabic',
    category: 'script',
    face: 'arabic',
    text: ARABIC_HELLO,
    note: 'Right-to-left with a space. Reading line text back in visual order is its own hazard, so this is recorded and reported rather than scored blind.',
  },

  // -------------------------------------------------------------------- URIs
  {
    id: 'uri-http',
    category: 'uri',
    face: 'latin',
    text: 'aa http://a.example/b',
    note: 'The shape every deck contains. Breaking a URL in the wrong place is visible to any reader who tries to type it back in.',
  },
  {
    id: 'uri-path',
    category: 'uri',
    face: 'latin',
    text: 'aa C:\\dir\\file.txt',
    note: 'A Windows path: backslash, colon and dot in one string, none of them in the class a reader would guess.',
  },

  // ------------------------------------------------------------- unbreakable
  {
    id: 'un-long',
    category: 'unbreakable',
    face: 'latin',
    text: 'antidisestablishmentarian',
    note: 'One word with no opportunity anywhere and nothing before it. This is the control for what PowerPoint does when the line would otherwise be empty - and the only string here that is deliberately left without an early space.',
    flags: [{}, { latinLnBrk: true }, { latinLnBrk: false }],
  },
  {
    id: 'un-longpair',
    category: 'unbreakable',
    face: 'latin',
    text: 'ab antidisestablishmentarian',
    note: 'The same word with a real opportunity in front of it, so a forced break inside the word is distinguishable from the break after the space. The variant with latinLnBrk absent is what settles ECMA against Office.',
    flags: [{}, { latinLnBrk: true }, { latinLnBrk: false }],
  },
];

// ------------------------------------------------------ the kinsoku sweep

/**
 * Every character asked whether it may begin or end a line.
 *
 * `ea-custom` and `ea-empty` established that PowerPoint carries a built-in
 * kinsoku list and uses it unless the file both states its own sets *and* turns
 * `strictFirstAndLastChars` off. That list is the single most load-bearing
 * table in this sub-phase, and copying it out of a comment in `a10-rtl-cjk`
 * would make it the one rule here that is neither cited nor measured.
 *
 * So each candidate is measured. `日本語` + X + `テスト`
 * puts X at position 4 between two runs of ideographs that break freely, and
 * one string answers both questions at once:
 *
 * - the opportunity at 3 is absent iff X may not **begin** a line;
 * - the opportunity at 4 is absent iff X may not **end** a line.
 *
 * The controls at the end are ordinary characters that no kinsoku list
 * contains. If any of them comes back restricted, the probe is measuring
 * something other than the list and the whole sweep is void - which is the
 * point of including them.
 */
const KINSOKU_START_CANDIDATES =
  // ASCII closers and terminators, which the Japanese sets do contain.
  '!%),.:;?]}' +
  // Latin-1 and General Punctuation.
  '¢’”‰′″℃' +
  // CJK punctuation: comma, full stop, iteration mark, the closing brackets.
  '、。々〉》」』】〕' +
  // Small hiragana, the voiced marks, the hiragana iteration marks.
  'ぁぃぅぇぉっゃゅょゎ' +
  '゛゜ゝゞ' +
  // Small katakana, the middle dot, the prolonged sound mark, iteration marks.
  'ァィゥェォッャュョヮ' +
  'ヵヶ・ーヽヾ' +
  // Fullwidth and halfwidth forms of the same closers.
  '！％），．：；？］｝' +
  '｡｣､･ﾞﾟ';

const KINSOKU_END_CANDIDATES = '$([\\{' + '£¥‘“' + '〈《「『【〔' + '＃＄（［｛￡￥';

/**
 * Characters that must come back unrestricted.
 *
 * Two kanji, two kana, a Latin letter and a digit: nothing any kinsoku list
 * names. A sweep that restricts one of these is not measuring kinsoku.
 */
const KINSOKU_CONTROLS = 'あア漢字A1';

/** Both sides of the sweep, plus the controls, de-duplicated and in a stable order. */
export function kinsokuCandidates(): readonly string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const ch of [
    ...codePoints(KINSOKU_START_CANDIDATES),
    ...codePoints(KINSOKU_END_CANDIDATES),
    ...codePoints(KINSOKU_CONTROLS),
  ]) {
    if (seen.has(ch)) continue;
    seen.add(ch);
    out.push(ch);
  }
  return out;
}

/**
 * One probe string per candidate.
 *
 * The id carries the code point rather than the character, so a failing row in
 * the report names something a reader can look up instead of a glyph they may
 * have no font for.
 */
export function kinsokuStrings(): readonly BreakString[] {
  return kinsokuCandidates().map((ch) => {
    const cp = (ch.codePointAt(0) ?? 0).toString(16).padStart(4, '0');
    return {
      id: `kn-${cp}`,
      category: 'kinsoku' as const,
      face: 'ea' as const,
      text: `${NIHONGO}${ch}${TESUTO}`,
      note: `Whether U+${cp.toUpperCase()} may begin a line (opportunity 3) or end one (opportunity 4).`,
    };
  });
}

/** Every probe string: the hand-written table above, then the kinsoku sweep. */
export function allStrings(): readonly BreakString[] {
  return [...STRINGS, ...kinsokuStrings()];
}

// ------------------------------------------------------------------- probes

export type ProbeKind =
  /** `wrap="none"`: PowerPoint's own width for the whole string, as the control. */
  | 'natural'
  /** `wrap="1"` at a box width that holds exactly `fitMax` code points and no more. */
  | 'cut';

export interface Probe {
  readonly id: string;
  readonly kind: ProbeKind;
  readonly stringId: string;
  readonly face: FaceKey;
  readonly text: string;
  readonly flags: BreakFlags;
  /** Which flag variant of the string this is; 0 when the string states none. */
  readonly variant: number;
  /**
   * For `cut`: how many code points the box was sized to hold.
   *
   * The measurement is what line 1 comes back as given that only this many fit.
   * A greedy breaker answers with the largest break opportunity at or before
   * `fitMax`, so the distinct answers across the sweep are the opportunity set.
   */
  readonly fitMax?: number | undefined;
  /** For `cut`: the box width, in points. */
  readonly widthPt?: number | undefined;
  /**
   * For `cut`: how far the box width sits from the nearest width that would
   * change `fitMax`.
   *
   * This is the probe's own error budget. T2 put Chromium within 0.15% of
   * PowerPoint at the median, so a margin of several points is safe by two
   * orders of magnitude - but the number is carried per probe rather than
   * assumed, because the analysis has to be able to say *which* probes were
   * thin instead of discovering it as a contradiction later.
   */
  readonly marginPt?: number | undefined;
}

/** Code points, not UTF-16 units. Every probe string is BMP; this keeps it honest anyway. */
export function codePoints(text: string): readonly string[] {
  return [...text];
}

/**
 * A width that two prefixes must differ by before a box can be placed between
 * them, in points.
 *
 * Below this they are the same width for measurement purposes and no box can
 * separate them. One point at 24pt is 0.04 em, and every visible glyph in every
 * probed face is wider than that - so what this merges is only ever a combining
 * mark or a zero-advance control: a ZWSP, a word joiner, a soft hyphen that was
 * not taken, a Thai vowel sign. Those cuts are merged rather than dropped, so
 * the sweep still asks about the position; it just asks about it as a range.
 */
const WIDTH_EPS = 1;

/**
 * The grid a box width is snapped to, in points.
 *
 * A box width has to be a whole number of EMU, because that is what the file can
 * hold, and it has to survive the float round trip through `points * 12700`.
 * Only a *dyadic* point value does both: 12700 is 4 x 3175, so a multiple of a
 * quarter point is exactly representable in binary and lands on a whole EMU,
 * while an arbitrary n/12700 is neither and comes back from the multiplication
 * a few ulps off an integer.
 *
 * The cost is that the box sits up to an eighth of a point away from the
 * midpoint. The narrowest advance any probe has at 24pt is a comma at about
 * 6.7pt, so the margin this eats into is at least 3.3pt - a factor of 26.
 * `marginPt` is recomputed after the snap rather than before, so what a probe
 * reports is the margin of the width the deck actually states.
 */
const QUANTUM = 0.25;

/** Prefix widths for one string: `w[k]` is the width of its first `k` code points, `w[0] = 0`. */
export type PrefixWidths = readonly number[];

/**
 * Every probe, with its box already sized.
 *
 * Generating this needs the measured widths, which is why it lives behind a
 * function taking them rather than being a constant: a box width that came from
 * nowhere is exactly the kind of number this repository does not commit.
 *
 * Cuts that resolve to the same `fitMax` are emitted once. A string containing
 * a zero-advance character has fewer distinct box widths than it has code
 * points, and running the same shape twice would inflate the probe count
 * without adding a measurement.
 */
export function makeProbes(widthsOf: (stringId: string) => PrefixWidths): readonly Probe[] {
  const probes: Probe[] = [];

  for (const s of allStrings()) {
    const variants = s.flags ?? [{}];
    const points = codePoints(s.text);
    const w = widthsOf(s.id);
    if (w.length !== points.length + 1) {
      throw new Error(
        `${s.id}: expected ${String(points.length + 1)} prefix widths, got ${String(w.length)}`,
      );
    }

    // For each cut, the smallest prefix strictly wider than it. Everything
    // between the two is the same width, so one box asks about all of them.
    const cuts: { fitMax: number; widthPt: number; marginPt: number }[] = [];
    for (let cut = 1; cut < points.length; cut += 1) {
      const here = w[cut];
      if (here === undefined) throw new Error(`${s.id}: no width for prefix ${String(cut)}`);
      let next = cut + 1;
      while (next <= points.length) {
        const there = w[next];
        if (there === undefined) throw new Error(`${s.id}: no width for prefix ${String(next)}`);
        if (there > here + WIDTH_EPS) break;
        next += 1;
      }
      // Every remaining code point is zero-advance: no box can hold `cut` and
      // exclude the rest, so there is nothing left to ask.
      if (next > points.length) continue;
      const upper = w[next];
      if (upper === undefined) throw new Error(`${s.id}: no width for prefix ${String(next)}`);
      const fitMax = next - 1;
      if (cuts.some((c) => c.fitMax === fitMax)) continue;
      const widthPt = Math.round((here + upper) / 2 / QUANTUM) * QUANTUM;
      // The distance to whichever boundary is nearer, not half the gap: after
      // the snap the box is no longer exactly in the middle, and the smaller
      // side is the one that decides whether the probe is sound.
      const marginPt = Math.min(widthPt - here, upper - widthPt);
      if (marginPt <= 0) {
        throw new Error(
          `${s.id}: box for ${String(fitMax)} code points snapped outside ` +
            `[${String(here)}, ${String(upper)}]`,
        );
      }
      cuts.push({ fitMax, widthPt, marginPt });
    }

    variants.forEach((flags, variant) => {
      const suffix = s.flags === undefined ? '' : `-v${String(variant)}`;
      probes.push({
        id: `${s.id}${suffix}-nat`,
        kind: 'natural',
        stringId: s.id,
        face: s.face,
        text: s.text,
        flags,
        variant,
      });
      for (const c of cuts) {
        probes.push({
          id: `${s.id}${suffix}-c${String(c.fitMax).padStart(2, '0')}`,
          kind: 'cut',
          stringId: s.id,
          face: s.face,
          text: s.text,
          flags,
          variant,
          fitMax: c.fitMax,
          widthPt: c.widthPt,
          marginPt: c.marginPt,
        });
      }
    });
  }

  return probes;
}

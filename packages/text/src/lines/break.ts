/**
 * Where a line may end, and where it does.
 *
 * Every rule here was measured against Microsoft PowerPoint in experiment T3 -
 * 2221 probes across twelve packages - by sizing a text box to hold exactly `n`
 * code points and reading back the *text* of line 1 through the object model.
 * Line 1 is the break position, so nothing had to be inferred from a width.
 * The fixture is `corpus/ground-truth/line-breaks.json`; the reasoning and the
 * refuted alternatives are in `docs/adr/0029-line-breaking.md`.
 *
 * ## This is not UAX#14, and that is the finding
 *
 * The plan called for UAX#14 through the `linebreak` package, then DrawingML
 * tailoring on top. The measurement refutes the shape of that design: what
 * PowerPoint does is not a *tailoring* of UAX#14 but a much smaller rule that
 * UAX#14 is a superset of. It does not break after a solidus, anywhere in a
 * URL, at a ZERO WIDTH SPACE, or at any of the number and symbol classes UAX#14
 * spends most of its rules on. Scored over the same probes, a faithful UAX#14
 * reading fits 1878 of 2212 where the rule below fits 2212.
 *
 * So there is no Unicode line-break table in this package and no dependency
 * that carries one. The classes that matter are: the space, four characters
 * that break after themselves, two that glue, East Asian script, and the
 * kinsoku sets. Everything else is a letter.
 *
 * ## The rule
 *
 * A line may end at position `k` - between code points `k-1` and `k` - when:
 *
 * - `k-1` is a space and `k` is not, so a run of spaces yields one opportunity
 *   after the last of them;
 * - `k-1` is a hyphen-minus, en dash, em dash or soft hyphen. Only *after*: the
 *   position before an em dash is not an opportunity, which is where this
 *   departs from UAX#14's two-sided B2 class;
 * - either neighbour is East Asian; or
 * - `a:pPr/@latinLnBrk` is `1`, which makes every remaining position an
 *   opportunity.
 *
 * Except that a NO-BREAK SPACE or a NON-BREAKING HYPHEN on either side
 * suppresses it, and - for East Asian positions only - the kinsoku sets do.
 *
 * ## The fit test
 *
 * Trailing spaces are never charged to the width, whatever `hangingPunct` says.
 * With `hangingPunct` on, which is its measured default, one trailing comma or
 * full stop is not charged either. And a line that ends on a soft hyphen pays
 * for the hyphen it draws, which is the one place a break costs more than the
 * text it contains.
 *
 * If nothing fits, the line takes exactly as many code points as do - PowerPoint
 * force-breaks an unbreakable word rather than overflowing, and `latinLnBrk="0"`
 * does not stop it.
 */

import { TextError } from './errors.js';
import { BUILT_IN_KINSOKU, HANGING_PUNCTUATION, type KinsokuSets } from './kinsoku.js';

/**
 * The three `a:pPr` attributes that tailor line breaking.
 *
 * `undefined` means the attribute is absent, which is not the same as either
 * boolean - see the defaults below. Restated here rather than imported from
 * `@pptx-studio/model` because this package is a layer beneath it.
 */
export interface BreakTailoring {
  /**
   * `@latinLnBrk`. **Absent behaves as `false`.**
   *
   * ECMA-376 gives the attribute a default of `true`. PowerPoint does not:
   * measured over 24 box widths on a 28-code-point string, a paragraph stating
   * nothing lays out identically to one stating `0` and differently from one
   * stating `1` at every width. Following the standard here breaks every long
   * word in every deck.
   */
  readonly latinLnBrk?: boolean | undefined;
  /**
   * `@eaLnBrk`.
   *
   * Measured **inert**: all three of absent, `1` and `0` lay out identically on
   * a nine-ideograph string at every box width. It is accepted and ignored
   * rather than absent from this interface, so that a caller passing it is not
   * quietly getting a different rule than they think, and so the next person to
   * wonder finds the answer here.
   */
  readonly eaLnBrk?: boolean | undefined;
  /**
   * `@hangingPunct`. **Absent behaves as `true`.**
   *
   * Only affects the six characters in `HANGING_PUNCTUATION`, and only the fit
   * test - it never adds or removes an opportunity. Measured inert on Latin
   * punctuation: a Latin comma hangs by way of the space after it, not by way
   * of this attribute.
   */
  readonly hangingPunct?: boolean | undefined;
}

export interface BreakOptions {
  readonly tailoring?: BreakTailoring | undefined;
  /**
   * The kinsoku sets in force. Defaults to PowerPoint's built-in list, which is
   * what applies unless a package both states `p:kinsoku` and turns
   * `strictFirstAndLastChars` off - see `kinsokuInForce`.
   */
  readonly kinsoku?: KinsokuSets | undefined;
}

const SPACE = ' ';

/**
 * The four characters a line may end on that are not a space.
 *
 * Measured one at a time, each in a string with a real opportunity in front of
 * it so that a refused break falls back somewhere visible: HYPHEN-MINUS, EN
 * DASH, EM DASH, SOFT HYPHEN. Every other candidate - solidus, reverse solidus,
 * colon, full stop, comma, the closing brackets - was refuted.
 */
const BREAK_AFTER: ReadonlySet<string> = new Set([
  '-', // HYPHEN-MINUS
  '–', // EN DASH
  '—', // EM DASH
  '­', // SOFT HYPHEN
]);

/**
 * Characters that suppress an opportunity on either side of themselves.
 *
 * Both are measured the same way, and the harder half of the question is the
 * one that matters: with `latinLnBrk="1"` turning every other position in the
 * string into an opportunity, the two positions around each of these stayed
 * closed. So they do not merely fail to create a break, they suppress one that
 * would otherwise exist.
 *
 * The NO-BREAK SPACE half of that was a generalisation from the hyphen until a
 * surviving mutant said so, and then it was measured rather than argued.
 */
const GLUE: ReadonlySet<string> = new Set([
  '‑', // NON-BREAKING HYPHEN
  ' ', // NO-BREAK SPACE
]);

/** SOFT HYPHEN: invisible until a line ends on it, and then it draws a hyphen. */
const SOFT_HYPHEN = '­';

/**
 * East Asian, for the purpose of "may a line break here".
 *
 * Han, kana and Hangul come from Unicode script properties, which the engine
 * already carries - so this needs no data table and no dependency. CJK
 * punctuation and the halfwidth and fullwidth forms are `Script=Common` and
 * have to be named by range; every character T3's sweep found restricted falls
 * in one of the two.
 */
const EA_SCRIPT = /^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]$/u;

/** CJK Symbols and Punctuation: the ideographic comma and stop, the brackets. */
const CJK_PUNCT_START = 0x3000;
const CJK_PUNCT_END = 0x303f;
/** Halfwidth and Fullwidth Forms: the fullwidth ASCII and the halfwidth kana. */
const WIDTH_FORMS_START = 0xff00;
const WIDTH_FORMS_END = 0xffef;

export function isEastAsian(ch: string): boolean {
  const cp = ch.codePointAt(0);
  if (cp === undefined) return false;
  if (cp >= CJK_PUNCT_START && cp <= CJK_PUNCT_END) return true;
  if (cp >= WIDTH_FORMS_START && cp <= WIDTH_FORMS_END) return true;
  return EA_SCRIPT.test(ch);
}

/** Code points, not UTF-16 units. Every index in this module counts these. */
export function toCodePoints(text: string): readonly string[] {
  return [...text];
}

/**
 * Every position at which a line may end, ascending.
 *
 * Positions are code point indices in `1 .. n-1`: position `k` means a line
 * ending after code point `k-1`. The end of the text is not included - it is
 * not a choice.
 */
export function breakOpportunities(text: string, options: BreakOptions = {}): readonly number[] {
  const points = toCodePoints(text);
  return opportunitiesOf(points, options);
}

function opportunitiesOf(points: readonly string[], options: BreakOptions): readonly number[] {
  const latinLnBrk = options.tailoring?.latinLnBrk ?? false;
  const kinsoku = options.kinsoku ?? BUILT_IN_KINSOKU;
  const out: number[] = [];

  for (let k = 1; k < points.length; k += 1) {
    const before = points[k - 1];
    const after = points[k];
    if (before === undefined || after === undefined) continue;
    if (GLUE.has(before) || GLUE.has(after)) continue;

    const eastAsian = isEastAsian(before) || isEastAsian(after);
    const open =
      (before === SPACE && after !== SPACE) || BREAK_AFTER.has(before) || eastAsian || latinLnBrk;
    if (!open) continue;

    // The kinsoku filter reaches only opportunities the East Asian rule
    // created. A package stating sets that name Latin characters, with the
    // strict list turned off, changes nothing about English text - measured,
    // and the reading ECMA supports scores 2162/2179 against this one's 2179.
    if (eastAsian && (kinsoku.noStart.has(after) || kinsoku.noEnd.has(before))) continue;

    out.push(k);
  }
  return out;
}

/** The width in points of the code points in `[start, end)`. */
export type RangeMeasurer = (start: number, end: number) => number;

export interface WrapInput {
  readonly text: string;
  /** The usable width, in points: the frame width less `lIns` and `rIns`. */
  readonly widthPt: number;
  readonly measure: RangeMeasurer;
  /**
   * The advance of a hyphen in the run's face, in points.
   *
   * Charged to a line that ends on a soft hyphen, because that line draws one.
   * Measured: at a box sized for exactly the code points up to and including a
   * soft hyphen, PowerPoint refused the break and fell back to the space before
   * it. The soft hyphen has no advance of its own, so the only thing that can
   * have overflowed is the hyphen it drew.
   */
  readonly hyphenWidthPt: number;
  readonly tailoring?: BreakTailoring | undefined;
  readonly kinsoku?: KinsokuSets | undefined;
}

export interface LineBox {
  /** Index of the first code point on the line. */
  readonly start: number;
  /** One past the last code point on the line, including anything that hangs. */
  readonly end: number;
  /**
   * One past the last code point charged to the width.
   *
   * Equal to `end` unless something hangs. The difference is what a renderer
   * must let overflow the frame, and what an alignment calculation must ignore.
   */
  readonly measuredEnd: number;
  /** Whether the line draws a hyphen after its last code point. */
  readonly hyphen: boolean;
  /** The text on the line, hanging characters included. */
  readonly text: string;
}

/**
 * How many code points at the end of `[from, k)` are not charged to the width.
 *
 * Trailing spaces always, however many; otherwise one hanging punctuation mark,
 * unless the paragraph turned hanging off. The two are exclusive because they
 * were measured that way: a line ending `", "` hangs the space, not the comma.
 */
function hangCount(
  points: readonly string[],
  from: number,
  k: number,
  hangingPunct: boolean,
): number {
  // The `>= from` bound is currently unobservable - every break, chosen or
  // forced, absorbs the space run that follows it, so no line but the first can
  // begin with a space, and for the first `from` is zero. It stays because a
  // helper that can walk into the previous line's characters is one invariant
  // change away from being wrong, and the invariant is not this function's.
  let spaces = 0;
  while (k - 1 - spaces >= from && points[k - 1 - spaces] === SPACE) spaces += 1;
  if (spaces > 0) return spaces;
  const last = points[k - 1];
  if (hangingPunct && last !== undefined && HANGING_PUNCTUATION.has(last)) return 1;
  return 0;
}

/**
 * The largest `k` in `(from, n]` whose raw width fits, never fewer than one
 * code point. Used only when no break opportunity fits, so the line has to be
 * cut mid-cluster.
 *
 * Scanned, not binary-searched, and that is a correction rather than a
 * preference. **Prefix widths are not monotone in a shaped script.** In Arabic,
 * `مرحبا` measures 45.3pt where
 * `مرحب` - one letter shorter - measures 51.0pt, because
 * the fifth letter joins the fourth into a shorter connected form. A binary
 * search over a sequence that dips like that returns an arbitrary element of
 * it, and T3's one Arabic probe is what caught it.
 *
 * The scan stops once the width has passed twice the box, which no shaping dip
 * recovers from - a ligature shortens a run, it does not halve it.
 */
function widestFit(measure: RangeMeasurer, from: number, n: number, widthPt: number): number {
  let best = from + 1;
  const cap = widthPt * 2;
  for (let k = from + 1; k <= n; k += 1) {
    const w = measure(from, k);
    if (w <= widthPt) best = k;
    else if (w > cap) break;
  }
  return best;
}

/**
 * Break `text` into lines.
 *
 * Greedy, and that is measured rather than assumed: every probe in T3 came back
 * as the longest prefix that ends at an opportunity and fits, with no
 * lookahead and no penalty minimisation.
 */
export function wrapText(input: WrapInput): readonly LineBox[] {
  if (!Number.isFinite(input.widthPt) || input.widthPt <= 0) {
    throw new TextError(
      'TEXT_BREAK_WIDTH',
      `wrap width ${String(input.widthPt)} is not a positive width`,
      String(input.widthPt),
    );
  }
  if (!Number.isFinite(input.hyphenWidthPt) || input.hyphenWidthPt < 0) {
    throw new TextError(
      'TEXT_BREAK_WIDTH',
      `hyphen width ${String(input.hyphenWidthPt)} is not a width`,
      String(input.hyphenWidthPt),
    );
  }

  const points = toCodePoints(input.text);
  const n = points.length;
  if (n === 0) return [];

  const hangingPunct = input.tailoring?.hangingPunct ?? true;
  const opportunities = opportunitiesOf(points, {
    ...(input.tailoring === undefined ? {} : { tailoring: input.tailoring }),
    ...(input.kinsoku === undefined ? {} : { kinsoku: input.kinsoku }),
  });

  const lines: LineBox[] = [];
  let from = 0;

  /** What ending a line that started at `from` at `k` costs: charged text plus any hyphen. */
  const charged = (from: number, k: number): number => {
    const width = input.measure(from, k - hangCount(points, from, k, hangingPunct));
    return points[k - 1] === SOFT_HYPHEN ? width + input.hyphenWidthPt : width;
  };

  while (from < n) {
    // Candidates are the opportunities after `from`, then the end of the text.
    // Each is measured, which is a handful of measurements per line - one per
    // word - rather than one per code point.
    //
    // The scan stops at the first candidate that does not fit, which relies on
    // the charged width increasing from one opportunity to the next. That holds
    // where per-code-point widths do not: an opportunity always falls on a
    // cluster boundary - after a space, a dash, or between scripts - and
    // shaping does not join across one. A soft hyphen is the exception, because
    // the hyphen it adds can make one candidate cost more than a later one, so
    // its failure does not end the scan.
    let best = 0;
    for (const k of opportunities) {
      if (k <= from) continue;
      if (charged(from, k) <= input.widthPt) best = k;
      else if (points[k - 1] !== SOFT_HYPHEN) break;
    }
    if (charged(from, n) <= input.widthPt) best = n;

    // Nothing fits: break at exactly what does. PowerPoint force-breaks an
    // unbreakable word rather than overflowing, and `latinLnBrk="0"` does not
    // stop it - measured on a 25-letter word with no opportunity anywhere,
    // which broke at every one of 24 box widths in all three variants.
    //
    // A forced break still takes any spaces that follow it. Measured on
    // `aa 、bb。cc` in a two-character box: the ideographic comma may not begin
    // a line so no opportunity was available, and PowerPoint returned `aa ` -
    // three code points - rather than `aa`. Hanging a trailing space is a
    // property of spaces, not of the break that produced them.
    let end = best;
    if (end <= from) {
      end = widestFit(input.measure, from, n, input.widthPt);
      while (end < n && points[end] === SPACE) end += 1;
    }
    const measuredEnd = end - hangCount(points, from, end, hangingPunct);
    lines.push({
      start: from,
      end,
      measuredEnd,
      hyphen: points[end - 1] === SOFT_HYPHEN,
      text: points.slice(from, end).join(''),
    });
    from = end;
  }

  return lines;
}

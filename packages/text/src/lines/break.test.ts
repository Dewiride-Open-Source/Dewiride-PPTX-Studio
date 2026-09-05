/**
 * Line breaking, re-derived from the T3 fixture.
 *
 * Every assertion below reads `corpus/ground-truth/line-breaks.json` and works
 * out what the answer should be from the measurement. None of them was written
 * by reading `break.ts`: a test that re-derives the rule fails when the code is
 * wrong, and a test written from the code passes forever and proves nothing.
 *
 * The centrepiece is the replay. The fixture carries the prefix widths that
 * sized every probe box, so the whole experiment - 2212 scored readings across twelve
 * packages - can be run again here with no browser and no PowerPoint, and the
 * implementation has to produce the same line 1 that PowerPoint did, on every
 * one of them.
 */

import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/line-breaks.json' with { type: 'json' };

import {
  breakOpportunities,
  isEastAsian,
  toCodePoints,
  wrapText,
  type BreakTailoring,
} from './break.js';
import { TextError } from './errors.js';
import {
  BUILT_IN_KINSOKU,
  HANGING_PUNCTUATION,
  kinsokuInForce,
  type KinsokuSets,
} from './kinsoku.js';

interface FixtureString {
  readonly id: string;
  readonly category: string;
  readonly face: string;
  readonly text: string;
  readonly widths: readonly number[];
}
interface FixtureDeck {
  readonly deck: string;
  readonly statesKinsoku: boolean;
  readonly invalStChars: string | null;
  readonly invalEndChars: string | null;
  readonly strictFirstAndLastChars: boolean | null;
  readonly opened: boolean;
}
/** One width sweep: a deck, a string, a flag variant, and three parallel columns. */
interface FixtureSweep {
  readonly deck: string;
  readonly string: string;
  readonly variant: number;
  readonly flags: BreakTailoring;
  readonly boxPt: readonly number[];
  readonly fitMax: readonly number[];
  readonly lineOne: readonly number[];
}

/** One probe, flattened out of a sweep. */
interface FixtureReading {
  readonly deck: string;
  readonly string: string;
  readonly variant: number;
  readonly face: string;
  readonly flags: BreakTailoring;
  readonly boxPt: number;
  readonly fitMax: number;
  readonly lineOne: number;
  readonly label: string;
}

const strings = fixture.strings as readonly FixtureString[];
const decks = fixture.decks as readonly FixtureDeck[];
const sweeps = fixture.readings as readonly FixtureSweep[];
const faceMetrics = fixture.faceMetrics as Readonly<
  Record<string, { family: string; hyphenPt: number; spacePt: number }>
>;

const stringById = new Map(strings.map((s) => [s.id, s]));
const deckById = new Map(decks.map((d) => [d.deck, d]));

/**
 * The sweeps flattened back into one entry per probe.
 *
 * The fixture is columnar - a sweep is a table of box widths against line
 * lengths, and storing it that way keeps it under the corpus size cap - but a
 * test reads more clearly one probe at a time.
 */
const readings: readonly FixtureReading[] = sweeps.flatMap((sweep) => {
  const str = stringById.get(sweep.string);
  if (str === undefined) throw new Error(`no string ${sweep.string}`);
  if (sweep.boxPt.length !== sweep.lineOne.length || sweep.boxPt.length !== sweep.fitMax.length) {
    throw new Error(`${sweep.deck}/${sweep.string}: sweep columns are ragged`);
  }
  return sweep.boxPt.map((boxPt, i) => ({
    deck: sweep.deck,
    string: sweep.string,
    variant: sweep.variant,
    face: str.face,
    flags: sweep.flags,
    boxPt,
    fitMax: sweep.fitMax[i] ?? 0,
    lineOne: sweep.lineOne[i] ?? 0,
    label: `${sweep.deck}/${sweep.string}/v${String(sweep.variant)} box=${String(boxPt)}`,
  }));
});

/** The sets a deck's probes were laid out under, from what the deck stated. */
function kinsokuForDeck(deck: FixtureDeck): KinsokuSets {
  const stated =
    deck.invalStChars === null || deck.invalEndChars === null
      ? undefined
      : {
          noStart: new Set(toCodePoints(deck.invalStChars)),
          noEnd: new Set(toCodePoints(deck.invalEndChars)),
        };
  return kinsokuInForce(stated, deck.strictFirstAndLastChars ?? undefined);
}

/**
 * A measurer over the fixture's prefix widths.
 *
 * These are exactly the numbers that decided how wide each probe box was
 * written, so replaying with them puts the implementation in front of the same
 * geometry PowerPoint saw.
 */
function measurerFor(str: FixtureString) {
  return (start: number, end: number): number => {
    const a = str.widths[start];
    const b = str.widths[end];
    if (a === undefined || b === undefined) {
      throw new Error(`${str.id}: no width for [${String(start)}, ${String(end)})`);
    }
    return b - a;
  };
}

/** Replay one reading and return the length of line 1 the implementation produced. */
function replay(reading: FixtureReading): number {
  const str = stringById.get(reading.string);
  const deck = deckById.get(reading.deck);
  if (str === undefined || deck === undefined) {
    throw new Error(`${reading.label}: missing string or deck`);
  }
  const metrics = faceMetrics[str.face];
  if (metrics === undefined) throw new Error(`no face metrics for ${str.face}`);
  const lines = wrapText({
    text: str.text,
    widthPt: reading.boxPt,
    measure: measurerFor(str),
    hyphenWidthPt: metrics.hyphenPt,
    tailoring: reading.flags,
    kinsoku: kinsokuForDeck(deck),
  });
  const first = lines[0];
  if (first === undefined) throw new Error(`${reading.label}: produced no lines`);
  return first.end;
}

const isThai = (r: FixtureReading): boolean => r.face === 'thai';

describe('the T3 replay', () => {
  it('reproduces PowerPoint on every probe that is not Thai', () => {
    const misses: string[] = [];
    let checked = 0;
    for (const reading of readings) {
      if (isThai(reading)) continue;
      checked += 1;
      const got = replay(reading);
      if (got !== reading.lineOne) {
        misses.push(`${reading.label} expected ${String(reading.lineOne)} got ${String(got)}`);
      }
    }
    expect(checked).toBeGreaterThan(2000);
    expect(misses).toEqual([]);
  });

  it('produces the exact text PowerPoint put on line 1', () => {
    // A length can agree by accident where a slice agrees only if the break is
    // in the same place, so the whole line is checked and not just its size.
    for (const reading of readings) {
      if (isThai(reading) || reading.deck !== 'latin') continue;
      const str = stringById.get(reading.string);
      const deck = deckById.get(reading.deck);
      if (str === undefined || deck === undefined) throw new Error('missing');
      const metrics = faceMetrics[str.face];
      if (metrics === undefined) throw new Error('missing metrics');
      const lines = wrapText({
        text: str.text,
        widthPt: reading.boxPt,
        measure: measurerFor(str),
        hyphenWidthPt: metrics.hyphenPt,
        tailoring: reading.flags,
        kinsoku: kinsokuForDeck(deck),
      });
      const points = [...str.text];
      expect(lines[0]?.text, reading.label).toBe(points.slice(0, reading.lineOne).join(''));
    }
  });

  it('over-runs Thai in exactly the places the fixture records, and no more', () => {
    // PowerPoint segments Thai with a dictionary. We do not, and no JavaScript
    // library carries one, so this is a scope boundary rather than a bug - see
    // ADR 0029. The count is pinned so the gap fails the suite if it grows.
    const thai = readings.filter(isThai);
    const misses = thai.filter((r) => replay(r) !== r.lineOne);
    const recorded = fixture.rules.thai as { fits: number; total: number };
    expect(thai.length).toBe(recorded.total);
    expect(thai.length - misses.length).toBe(recorded.fits);
  });
});

describe('the opportunity set, against the alternatives', () => {
  /** What PowerPoint's opportunities were for one string, derived from the readings. */
  function measuredOpportunities(stringId: string, deck = 'latin'): ReadonlySet<number> {
    // Variant 0 only. A string with flag variants appears in the fixture three
    // times, and the latinLnBrk="1" run opens every position - mixing them in
    // would report an opportunity set no single paragraph ever had.
    const rows = readings.filter(
      (r) => r.string === stringId && r.deck === deck && r.variant === 0,
    );
    expect(rows.length).toBeGreaterThan(0);
    const out = new Set<number>();
    for (const r of rows) {
      // A reading shorter than the box could hold is a break PowerPoint
      // *preferred*, so the position is a real opportunity. A reading equal to
      // the box is either an opportunity or a forced break, and says nothing.
      if (r.lineOne < r.fitMax || r.lineOne > r.fitMax) out.add(r.lineOne);
    }
    return out;
  }

  it('breaks after a space run, and only after the last space of it', () => {
    const opps = breakOpportunities('aaaa  bbbb  cccc');
    expect(opps).toContain(6);
    expect(opps).not.toContain(5);
    expect(measuredOpportunities('sp-double')).toContain(6);
  });

  it('breaks after a hyphen but never before one', () => {
    const opps = breakOpportunities('aa well-known-thing');
    expect(opps).toEqual([3, 8, 14]);
    expect([...measuredOpportunities('hy-hyphen')].sort((a, b) => a - b)).toEqual([3, 8, 14]);
  });

  it('does not break after a solidus, which UAX#14 would allow', () => {
    expect(breakOpportunities('aa alpha/beta')).toEqual([3]);
    expect([...measuredOpportunities('hy-slash')]).toEqual([3]);
  });

  it('does not break anywhere in a URL or a Windows path', () => {
    expect(breakOpportunities('aa http://a.example/b')).toEqual([3]);
    expect(breakOpportunities('aa C:\\dir\\file.txt')).toEqual([3]);
    expect([...measuredOpportunities('uri-http')]).toEqual([3]);
    expect([...measuredOpportunities('uri-path')]).toEqual([3]);
  });

  it('ignores ZERO WIDTH SPACE, which UAX#14 makes a break', () => {
    expect(breakOpportunities('aa bbbb\u200bcccc')).toEqual([3]);
    expect([...measuredOpportunities('iv-zwsp')]).toEqual([3]);
  });

  it('breaks at a soft hyphen', () => {
    expect(breakOpportunities('aa bbbb\u00adcccc')).toEqual([3, 8]);
    expect([...measuredOpportunities('iv-shy')].sort((a, b) => a - b)).toEqual([3, 8]);
  });

  it('never breaks at a NO-BREAK SPACE or a NON-BREAKING HYPHEN', () => {
    expect(breakOpportunities('aa bbbb\u00a0cccc')).toEqual([3]);
    expect(breakOpportunities('aa well\u2011known')).toEqual([3]);
    expect([...measuredOpportunities('iv-nbsp')]).toEqual([3]);
    expect([...measuredOpportunities('hy-nobreak')]).toEqual([3]);
  });

  it('breaks after an em dash but not before it', () => {
    // The one place this departs from UAX#14, whose B2 class breaks on both
    // sides. Position 8 is before the dash and 9 is after it.
    const opps = breakOpportunities('aa alpha\u2014beta');
    expect(opps).toContain(9);
    expect(opps).not.toContain(8);
    expect([...measuredOpportunities('hy-emdash')].sort((a, b) => a - b)).toEqual([3, 9]);
  });

  it('does not break inside a number, at a comma, a point, a currency sign or a percent', () => {
    for (const id of ['nu-thousands', 'nu-decimal', 'nu-currency', 'nu-percent']) {
      const str = stringById.get(id);
      if (str === undefined) throw new Error(id);
      expect(breakOpportunities(str.text), id).toEqual([3]);
      expect([...measuredOpportunities(id)], id).toEqual([3]);
    }
  });

  it('breaks a hyphenated date, where the number rule does not reach', () => {
    expect(breakOpportunities('aa 2026-09-04')).toEqual([3, 8, 11]);
  });

  it('breaks between two ideographs and at a script transition', () => {
    const opps = breakOpportunities('\u65e5\u672c\u8a9eabcdef\u3067\u3059');
    // 1,2,3 are ideograph boundaries and the transition; 4..8 are inside the
    // Latin word; 9 is the transition back and 10 is between two kana.
    expect(opps).toEqual([1, 2, 3, 9, 10]);
  });
});

describe('what an absent attribute means', () => {
  const defaults = fixture.rules.flagDefaults as Readonly<Record<string, string>>;

  it('treats an absent latinLnBrk as false, which ECMA does not', () => {
    expect(defaults['latinLnBrk']).toBe('false');
    const word = 'ab antidisestablishmentarian';
    expect(breakOpportunities(word)).toEqual([3]);
    expect(breakOpportunities(word, { tailoring: { latinLnBrk: false } })).toEqual([3]);
    expect(breakOpportunities(word, { tailoring: { latinLnBrk: true } }).length).toBeGreaterThan(
      20,
    );
  });

  it('treats an absent hangingPunct as true', () => {
    expect(defaults['hangingPunct']).toBe('true');
  });

  it('records eaLnBrk as measured inert', () => {
    expect(defaults['eaLnBrk']).toBe('inert');
    const jp = '\u65e5\u672c\u8a9e\u306e\u30c6\u30b9\u30c8\u3067\u3059';
    const off = breakOpportunities(jp, { tailoring: { eaLnBrk: false } });
    const on = breakOpportunities(jp, { tailoring: { eaLnBrk: true } });
    expect(off).toEqual(on);
    expect(off).toEqual(breakOpportunities(jp));
  });

  it('leaves latinLnBrk="1" unable to break a NON-BREAKING HYPHEN', () => {
    const opps = breakOpportunities('aa well\u2011known', { tailoring: { latinLnBrk: true } });
    expect(opps).not.toContain(7);
    expect(opps).not.toContain(8);
    expect(opps).toContain(6);
  });

  it('leaves latinLnBrk="1" unable to break a NO-BREAK SPACE either', () => {
    // The harder half of the no-break question, and measured rather than
    // generalised from the hyphen: with every other position in the string open,
    // the two around the NBSP stay closed.
    const opps = breakOpportunities('aa bbbb\u00a0cccc', {
      tailoring: { latinLnBrk: true },
    });
    expect(opps).not.toContain(7);
    expect(opps).not.toContain(8);
    expect(opps).toContain(6);
    expect(opps).toContain(9);
  });
});

describe('the kinsoku sets', () => {
  it('matches the measured table exactly', () => {
    const measured = fixture.kinsoku as {
      mayNotBeginLine: readonly string[];
      mayNotEndLine: readonly string[];
      hangOutsideTheMeasure: readonly string[];
    };
    expect([...BUILT_IN_KINSOKU.noStart].sort()).toEqual([...measured.mayNotBeginLine].sort());
    expect([...BUILT_IN_KINSOKU.noEnd].sort()).toEqual([...measured.mayNotEndLine].sort());
    expect([...HANGING_PUNCTUATION].sort()).toEqual([...measured.hangOutsideTheMeasure].sort());
  });

  it('leaves the six control characters unrestricted', () => {
    const controls = (fixture.kinsoku as { unrestrictedControls: readonly string[] })
      .unrestrictedControls;
    for (const ch of controls) {
      expect(BUILT_IN_KINSOKU.noStart.has(ch), ch).toBe(false);
      expect(BUILT_IN_KINSOKU.noEnd.has(ch), ch).toBe(false);
    }
  });

  it('disagrees with the widely repeated list in exactly four places', () => {
    const wrong = (fixture.kinsoku as { disagreesWithA10Comment: readonly string[] })
      .disagreesWithA10Comment;
    expect(wrong).toEqual(['\u00a2', '\u00a3', '\u00a5', '\uff03']);
    for (const ch of wrong) {
      expect(BUILT_IN_KINSOKU.noStart.has(ch), ch).toBe(false);
      expect(BUILT_IN_KINSOKU.noEnd.has(ch), ch).toBe(false);
    }
    // ...while restricting the fullwidth forms the list leaves out.
    expect(BUILT_IN_KINSOKU.noEnd.has('\uffe1')).toBe(true);
    expect(BUILT_IN_KINSOKU.noEnd.has('\uffe5')).toBe(true);
  });

  it('ignores the file sets unless strictFirstAndLastChars is explicitly 0', () => {
    const stated: KinsokuSets = { noStart: new Set(['x']), noEnd: new Set(['y']) };
    expect(kinsokuInForce(stated, undefined)).toBe(BUILT_IN_KINSOKU);
    expect(kinsokuInForce(stated, true)).toBe(BUILT_IN_KINSOKU);
    expect(kinsokuInForce(stated, false)).toBe(stated);
    expect(kinsokuInForce(undefined, false)).toBe(BUILT_IN_KINSOKU);
  });

  it('replaces the built-in list entirely, leaving no hard core', () => {
    // Measured with empty sets: every position in every East Asian probe
    // became an opportunity, including the one before an ideographic comma.
    const empty: KinsokuSets = { noStart: new Set(), noEnd: new Set() };
    const jp = '\u65e5\u672c\u8a9e\u3001\u30c6\u30b9\u30c8\u3002\u3067\u3059';
    expect(breakOpportunities(jp, { kinsoku: empty })).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(breakOpportunities(jp)).not.toContain(3);
    expect(breakOpportunities(jp)).not.toContain(7);
  });

  it('applies only to East Asian positions, not to every break', () => {
    // The reading ECMA supports would suppress the space break here. Measured:
    // a package forbidding `b` at the start of a line changed nothing about
    // English text.
    const latin: KinsokuSets = { noStart: new Set(['b']), noEnd: new Set(['a']) };
    expect(breakOpportunities('aaaa bbbb', { kinsoku: latin })).toEqual([5]);
  });
});

describe('the fit test', () => {
  /** A monospaced measurer: one point per code point, so widths are exact. */
  const mono = (start: number, end: number): number => end - start;

  it('never charges a trailing space, whatever hangingPunct says', () => {
    for (const hangingPunct of [true, false, undefined]) {
      const lines = wrapText({
        text: 'aaaa bbbb',
        widthPt: 4,
        measure: mono,
        hyphenWidthPt: 1,
        tailoring: { hangingPunct },
      });
      expect(lines[0]?.end, String(hangingPunct)).toBe(5);
      expect(lines[0]?.measuredEnd, String(hangingPunct)).toBe(4);
    }
  });

  it('hangs one comma or full stop when hangingPunct is on, and not when it is off', () => {
    const jp = '\u65e5\u672c\u8a9e\u3001\u30c6\u30b9\u30c8';
    const on = wrapText({ text: jp, widthPt: 3, measure: mono, hyphenWidthPt: 1 });
    expect(on[0]?.end).toBe(4);
    expect(on[0]?.measuredEnd).toBe(3);
    const off = wrapText({
      text: jp,
      widthPt: 3,
      measure: mono,
      hyphenWidthPt: 1,
      tailoring: { hangingPunct: false },
    });
    // Position 3 is closed - the comma may not begin a line - so it falls to 2.
    expect(off[0]?.end).toBe(2);
  });

  it('charges a soft-hyphen break for the hyphen it draws', () => {
    const text = 'aa bbbb\u00adcccc';
    const points = [...text];
    // A soft hyphen has no advance until a line ends on it, so the measurer
    // must not charge for it either - which `mono` would, being one point per
    // code point. This is the real shape of the question.
    const shy = (start: number, end: number): number => {
      let w = 0;
      for (let i = start; i < end; i += 1) if (points[i] !== '\u00ad') w += 1;
      return w;
    };
    // Eight code points measure seven, so with no hyphen to draw the break at
    // the soft hyphen fits a seven-point box.
    const free = wrapText({ text, widthPt: 7, measure: shy, hyphenWidthPt: 0 });
    expect(free[0]?.end).toBe(8);
    expect(free[0]?.hyphen).toBe(true);
    // Charge for the hyphen and the same break no longer fits, so it falls all
    // the way back to the space - which is what PowerPoint did.
    const charged = wrapText({ text, widthPt: 7, measure: shy, hyphenWidthPt: 1 });
    expect(charged[0]?.end).toBe(3);
    expect(charged[0]?.hyphen).toBe(false);
  });

  it('force-breaks an unbreakable word, even with latinLnBrk="0"', () => {
    for (const latinLnBrk of [true, false, undefined]) {
      const lines = wrapText({
        text: 'antidisestablishmentarian',
        widthPt: 5,
        measure: mono,
        hyphenWidthPt: 1,
        tailoring: { latinLnBrk },
      });
      expect(lines[0]?.end, String(latinLnBrk)).toBe(5);
    }
  });

  it('prefers a real opportunity over a forced break', () => {
    // With a space at 3, a box that could hold ten code points still stops at
    // the space rather than force-breaking inside the word.
    const lines = wrapText({
      text: 'ab antidisestablishmentarian',
      widthPt: 10,
      measure: mono,
      hyphenWidthPt: 1,
    });
    expect(lines[0]?.end).toBe(3);
  });

  it('always advances, so a box narrower than one code point still terminates', () => {
    const lines = wrapText({ text: 'abc', widthPt: 0.25, measure: mono, hyphenWidthPt: 1 });
    expect(lines.map((l) => l.text)).toEqual(['a', 'b', 'c']);
  });

  it('gives one line when the whole text fits', () => {
    // The end of the text has to be a candidate in its own right. Without it the
    // last opportunity wins, and text that fits comfortably still wraps - a
    // mutation that dropped the candidate survived every other assertion here.
    const lines = wrapText({ text: 'aa bb cc', widthPt: 100, measure: mono, hyphenWidthPt: 1 });
    expect(lines).toHaveLength(1);
    expect(lines[0]?.text).toBe('aa bb cc');
  });

  it('looks past a soft-hyphen candidate that does not fit', () => {
    // The candidate scan stops at the first opportunity that does not fit,
    // which is sound because the charged width rises from one opportunity to
    // the next - except at a soft hyphen, whose drawn hyphen can make it cost
    // more than a *later* break where a space hangs instead. So a failed soft
    // hyphen must not end the scan.
    const text = 'ab\u00adcd efghi';
    const points = [...text];
    const measure = (start: number, end: number): number => {
      let w = 0;
      for (let i = start; i < end; i += 1) if (points[i] !== '\u00ad') w += 1;
      return w;
    };
    // Opportunities are 3 (the soft hyphen) and 6 (the space). At a width of 8
    // the soft hyphen costs 2 plus a 7-point hyphen and fails; the space break
    // charges 4 and fits; the whole text costs 9 and does not. Ending the scan
    // at the soft hyphen loses the space break, falls through to the forced
    // path, and returns nine code points instead of six.
    const lines = wrapText({ text, widthPt: 8, measure, hyphenWidthPt: 7 });
    expect(lines[0]?.end).toBe(6);
    expect(lines[0]?.hyphen).toBe(false);
  });

  it('scans past a dip when forced to break mid-word', () => {
    // Prefix widths are not monotone in a shaped script: T3 measured Arabic
    // مرحبا at 45.3pt where مرحب, one letter shorter, is 51.0pt. So the forced
    // path cannot stop at the first prefix that overflows - it has to keep
    // looking, and the measured dip is what says how far. Eleven percent there;
    // the scan gives up at twice the box, which no ligature recovers from.
    const widths = [0, 1, 2, 3, 5, 4, 6];
    const measure = (start: number, end: number): number => {
      const a = widths[start];
      const b = widths[end];
      if (a === undefined || b === undefined) throw new Error('out of range');
      return b - a;
    };
    const lines = wrapText({ text: 'abcdef', widthPt: 4, measure, hyphenWidthPt: 1 });
    // Five code points measure 4 and fit; four measure 5 and do not. A scan
    // that stopped at the first overflow would return three.
    expect(lines[0]?.end).toBe(5);
  });

  it('takes the trailing spaces with a forced break too', () => {
    // Measured: an ideographic comma may not begin a line, so a two-character
    // box on `aa \u3001bb` has no opportunity at all - and PowerPoint still
    // returned three code points, the space among them.
    const lines = wrapText({
      text: 'aa \u3001bb',
      widthPt: 2,
      measure: mono,
      hyphenWidthPt: 1,
    });
    expect(lines[0]?.text).toBe('aa ');
    expect(lines[0]?.measuredEnd).toBe(2);
  });

  it('covers the text exactly once, with no gap and no overlap', () => {
    const text = 'alpha beta gamma delta epsilon';
    const lines = wrapText({ text, widthPt: 9, measure: mono, hyphenWidthPt: 1 });
    expect(lines.map((l) => l.text).join('')).toBe(text);
    for (let i = 1; i < lines.length; i += 1) {
      expect(lines[i]?.start).toBe(lines[i - 1]?.end);
    }
  });
});

describe('input validation', () => {
  const mono = (start: number, end: number): number => end - start;

  it('refuses a width that is not positive', () => {
    for (const widthPt of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => wrapText({ text: 'a', widthPt, measure: mono, hyphenWidthPt: 0 })).toThrow(
        TextError,
      );
    }
  });

  it('refuses a hyphen advance that is not a width', () => {
    expect(() =>
      wrapText({ text: 'a', widthPt: 10, measure: mono, hyphenWidthPt: -1 }),
    ).toThrowError(/hyphen width/);
  });

  it('returns no lines for empty text', () => {
    expect(wrapText({ text: '', widthPt: 10, measure: mono, hyphenWidthPt: 0 })).toEqual([]);
  });
});

describe('East Asian classification', () => {
  it('puts every restricted character inside an East Asian position', () => {
    // The filter runs on positions, not characters, and a restricted character
    // need not itself be East Asian - the sets include ASCII punctuation, which
    // is restricted only because of the ideographs around it. So the sweep is
    // replayed: the same string it was measured in, and the same two positions.
    const measured = fixture.kinsoku as {
      mayNotBeginLine: readonly string[];
      mayNotEndLine: readonly string[];
    };
    const at = (ch: string): readonly number[] => breakOpportunities(`日本語${ch}テスト`);
    const name = (ch: string): string => `U+${(ch.codePointAt(0) ?? 0).toString(16)}`;
    for (const ch of measured.mayNotBeginLine) {
      expect(at(ch), `${name(ch)} may not begin a line`).not.toContain(3);
      // ...and it must still be free on the other side, or the probe proves
      // nothing about which of the two sets it is in.
      expect(at(ch), `${name(ch)} may still end a line`).toContain(4);
    }
    for (const ch of measured.mayNotEndLine) {
      expect(at(ch), `${name(ch)} may not end a line`).not.toContain(4);
      expect(at(ch), `${name(ch)} may still begin a line`).toContain(3);
    }
  });

  it('leaves an unrestricted character both positions', () => {
    for (const ch of (fixture.kinsoku as { unrestrictedControls: readonly string[] })
      .unrestrictedControls) {
      const opps = breakOpportunities(`日本語${ch}テスト`);
      expect(opps, ch).toContain(3);
      expect(opps, ch).toContain(4);
    }
  });

  it('does not classify ordinary Latin as East Asian', () => {
    for (const ch of [...'abcXYZ019 -.,']) expect(isEastAsian(ch), ch).toBe(false);
  });

  it('counts code points, not UTF-16 units', () => {
    // A kanji outside the BMP: two UTF-16 units, one code point, one column.
    expect(toCodePoints('\u{20000}a')).toEqual(['\u{20000}', 'a']);
    expect(isEastAsian('\u{20000}')).toBe(true);
  });
});

/**
 * Experiment T3, step 4 - score the candidate models and emit the fixture.
 *
 * ```
 * node tools/ground-truth/text/line-breaks/analyse.ts <work-dir>
 * ```
 *
 * Reads `break-inputs.json` and `break-readings.json`, scores every candidate
 * reading against every probe, and writes `corpus/ground-truth/line-breaks.json`
 * only if exactly one candidate fits each question perfectly. A model that
 * fits 2533 of 2534 is not the rule; it is a rule with a counterexample, and
 * the counterexample is the finding.
 *
 * ## How a model is scored
 *
 * One simulator, parameterised. For a probe whose box is `W` points wide, the
 * simulator asks the model for the break opportunities in the string and for
 * the width each candidate prefix is *charged*, then applies the greedy rule
 * and compares the predicted length of line 1 against what PowerPoint laid out.
 *
 * The prefix widths come from Chromium, measured in step 1. They are used here
 * rather than PowerPoint's own because they are what sized the box in the first
 * place - so a model is scored against exactly the geometry the deck states.
 * Section G checks the two against each other independently.
 *
 * ## What is deliberately not imported
 *
 * Nothing from `@pptx-studio/text`. The models below are written from the
 * measurement, and the package is written from the fixture this emits; if the
 * scorer imported the implementation, a wrong implementation would score itself
 * correct and the tests derived from the fixture would agree with it forever.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import {
  FACES,
  PROBE_SIZE,
  allStrings,
  codePoints,
  type BreakFlags,
  type Probe,
} from './probes.ts';

const dir = process.argv[2];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/text/line-breaks/analyse.ts <work-dir>');
const work = resolve(dir);

// ------------------------------------------------------------------- inputs

interface DeckInput {
  readonly deck: string;
  readonly file: string;
  readonly kinsoku: boolean;
  readonly strictFirstAndLastChars: boolean | null;
  readonly invalStChars: string | null;
  readonly invalEndChars: string | null;
  readonly flagOverride: BreakFlags | null;
}
interface StringInput {
  readonly id: string;
  readonly category: string;
  readonly face: string;
  readonly text: string;
  readonly note: string;
  readonly points: readonly string[];
  readonly widths: readonly number[];
}
interface Inputs {
  readonly sizePt: number;
  readonly chromium: string;
  readonly fontsCheck: Readonly<Record<string, boolean>>;
  /** The advance of a real hyphen in each face, for the soft-hyphen fit cost. */
  readonly hyphenWidth: Readonly<Record<string, number>>;
  /** The space advance in each face: the width of the paragraph mark T2 found. */
  readonly spaceWidth: Readonly<Record<string, number>>;
  readonly strings: readonly StringInput[];
  readonly decks: readonly DeckInput[];
  readonly probes: readonly Probe[];
}

interface LineReading {
  readonly index: number;
  readonly text: string | null;
  readonly width: number | null;
  readonly length: number | null;
}
interface ShapeReading {
  readonly id: string;
  readonly width: number | null;
  readonly lineCount: number | null;
  readonly text: string | null;
  readonly lines: readonly LineReading[];
}
interface DeckReading {
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly error: string | null;
  readonly shapes: readonly ShapeReading[];
}

const inputs = JSON.parse(readFileSync(join(work, 'break-inputs.json'), 'utf8')) as Inputs;
const readings = JSON.parse(readFileSync(join(work, 'break-readings.json'), 'utf8')) as {
  readonly decks: readonly DeckReading[];
};

const stringById = new Map(inputs.strings.map((s) => [s.id, s]));
const probeById = new Map(inputs.probes.map((p) => [p.id, p]));
const deckInputById = new Map(inputs.decks.map((d) => [d.deck, d]));

/** One measured probe: what it was asked, and what PowerPoint answered. */
interface Case {
  readonly deck: string;
  readonly probe: Probe;
  readonly string: StringInput;
  readonly observed: number;
  readonly lineText: string;
}

const cases: Case[] = [];
const refused: string[] = [];
const repaired: string[] = [];
/** Every `wrap="none"` control, for the width comparison in section G. */
const controls: { deck: string; id: string; face: string; powerpoint: number; chromium: number }[] =
  [];

for (const deck of readings.decks) {
  if (!deck.opened) {
    refused.push(`${deck.deck}: ${deck.error ?? 'no reason given'}`);
    continue;
  }
  if (deck.repaired === true) repaired.push(deck.deck);
  for (const shape of deck.shapes) {
    const probe = probeById.get(shape.id);
    if (probe === undefined) throw new Error(`${deck.deck}: shape ${shape.id} is not a probe`);
    const str = stringById.get(probe.stringId);
    if (str === undefined) throw new Error(`no string ${probe.stringId}`);

    if (probe.kind === 'natural') {
      const pp = shape.width;
      const cr = str.widths.at(-1);
      if (pp !== null && cr !== undefined) {
        controls.push({
          deck: deck.deck,
          id: probe.id,
          face: str.face,
          powerpoint: pp,
          chromium: cr,
        });
      }
      continue;
    }

    const first = shape.lines[0];
    if (first === undefined || first.length === null || first.text === null) {
      throw new Error(`${deck.deck}/${shape.id}: no line 1 reading`);
    }
    // Two independent readings of the same fact. `Length` is PowerPoint's own
    // count; the text is what it put there. They must agree, or the reader is
    // wrong about something and nothing below can be trusted.
    if (first.length !== codePoints(first.text).length) {
      throw new Error(
        `${deck.deck}/${shape.id}: Lines(1).Length is ${String(first.length)} but the text is ` +
          `${String(codePoints(first.text).length)} code points (${JSON.stringify(first.text)})`,
      );
    }
    // A deck may force paragraph flags on every probe in it - the kinsoku sweep
    // is run twice, once with hanging punctuation off. The probe table records
    // what the *string* states, so the deck's override is folded in here rather
    // than left to be forgotten at the point of use.
    const override = deckInputById.get(deck.deck)?.flagOverride ?? null;
    cases.push({
      deck: deck.deck,
      probe: override === null ? probe : { ...probe, flags: override },
      string: str,
      observed: first.length,
      lineText: first.text,
    });
  }
}

// ------------------------------------------------------------------- models

/**
 * The kinsoku sets in force for a deck, as characters.
 *
 * Filled in from section F's measurement for the built-in list, and from what
 * the deck states for the ones that override it.
 */
interface Kinsoku {
  readonly noStart: ReadonlySet<string>;
  readonly noEnd: ReadonlySet<string>;
}
const EMPTY_KINSOKU: Kinsoku = { noStart: new Set(), noEnd: new Set() };

/**
 * East Asian, for the purposes of "may a line break here".
 *
 * Han, kana and Hangul come from Unicode script properties, which the JavaScript
 * engine already carries - no data file, no dependency. CJK punctuation and the
 * fullwidth and halfwidth forms are Script=Common and have to be named by range;
 * every character the kinsoku sweep restricted falls in one of them.
 */
const EA_SCRIPT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
function isEastAsian(ch: string): boolean {
  const cp = ch.codePointAt(0) ?? 0;
  if (cp >= 0x3000 && cp <= 0x303f) return true; // CJK Symbols and Punctuation
  if (cp >= 0xff00 && cp <= 0xffef) return true; // Halfwidth and Fullwidth Forms
  return EA_SCRIPT.test(ch);
}

const SPACE = ' ';
/** Characters after which PowerPoint breaks, measured one at a time in the `latin` deck. */
const BREAK_AFTER = new Set(['-', '–', '—', '­']);
/** Characters that suppress a break on both sides even when latinLnBrk is on. */
const GLUE = new Set(['‑', ' ']);

interface Model {
  readonly name: string;
  /** Positions 1..n-1 at which a line may end. */
  opportunities(
    points: readonly string[],
    flags: BreakFlags,
    kinsoku: Kinsoku,
  ): ReadonlySet<number>;
  /** Code points at the end of a prefix that are not charged to the fit test. */
  hangCount(points: readonly string[], k: number, flags: BreakFlags): number;
  /**
   * Whether ending a line at `k` renders a hyphen the line has to pay for.
   *
   * True only for a soft hyphen, and the reason this exists at all is one
   * probe: at a box sized for exactly the eight code points up to and
   * including the soft hyphen, PowerPoint refused the break and fell all the
   * way back to the space. The soft hyphen has no advance of its own, so the
   * only thing that can have overflowed is the hyphen it draws in its place.
   */
  rendersHyphen(points: readonly string[], k: number): boolean;
  /** Whether a box too narrow for any opportunity breaks anyway. */
  readonly emergency: boolean;
}

/** The six characters section F measured as hanging. Nothing else does. */
const HANGABLE = new Set([',', '.', '、', '。', '，', '．']);

function trailingSpaces(points: readonly string[], k: number): number {
  let n = 0;
  while (n < k && points[k - 1 - n] === SPACE) n += 1;
  return n;
}

/** The measured reading. Every clause below is one row of the report. */
function measuredOpportunities(
  points: readonly string[],
  flags: BreakFlags,
  kinsoku: Kinsoku,
): ReadonlySet<number> {
  const out = new Set<number>();
  const n = points.length;
  for (let k = 1; k < n; k += 1) {
    const before = points[k - 1];
    const after = points[k];
    if (before === undefined || after === undefined) continue;
    if (GLUE.has(before) || GLUE.has(after)) continue;

    const eastAsian = isEastAsian(before) || isEastAsian(after);
    let ok = false;
    // A run of spaces gives one opportunity, after the last of them.
    if (before === SPACE && after !== SPACE) ok = true;
    else if (BREAK_AFTER.has(before)) ok = true;
    else if (eastAsian) ok = true;
    else if (flags.latinLnBrk === true) ok = true;

    if (!ok) continue;
    // The kinsoku filter reaches only the opportunities the East Asian rule
    // created. `latin-custom` states sets naming `b` and `a`, turns the strict
    // list off, and changes nothing about English text - so the filter is not a
    // general character filter over every break. Whether the gate is on the
    // character's script or on the run's language this cannot say, because that
    // deck's runs are `en-US` and its opportunity is a space; but PowerPoint
    // refuses a package whose `p:kinsoku/@lang` is not East Asian outright,
    // which points the same way.
    if (eastAsian) {
      if (kinsoku.noStart.has(after)) continue;
      if (kinsoku.noEnd.has(before)) continue;
    }
    out.add(k);
  }
  return out;
}

/** The soft hyphen, which is invisible until a line ends on it. */
const SHY = '­';

const MEASURED: Model = {
  name: 'measured',
  opportunities: measuredOpportunities,
  rendersHyphen: (points, k) => points[k - 1] === SHY,
  hangCount(points, k, flags) {
    const spaces = trailingSpaces(points, k);
    if (spaces > 0) return spaces;
    // Hanging punctuation is one character, and only when it is switched on -
    // which it is unless the file says otherwise.
    const last = points[k - 1];
    if (flags.hangingPunct !== false && last !== undefined && HANGABLE.has(last)) return 1;
    return 0;
  },
  emergency: true,
};

/**
 * The reading anyone would write first, following UAX#14 and ECMA.
 *
 * Breaks after the solidus and the other symbol classes, honours ZERO WIDTH
 * SPACE and WORD JOINER, and takes ECMA at its word that `latinLnBrk` defaults
 * to true. Every one of those is wrong, and this row is here to say by how much.
 */
const UAX14ISH: Model = {
  name: 'uax14-ish',
  rendersHyphen: () => false,
  opportunities(points, flags, kinsoku) {
    const out = new Set<number>();
    const n = points.length;
    for (let k = 1; k < n; k += 1) {
      const before = points[k - 1];
      const after = points[k];
      if (before === undefined || after === undefined) continue;
      if (before === '⁠' || after === '⁠') continue;
      if (GLUE.has(before) || GLUE.has(after)) continue;
      let ok = false;
      if (before === SPACE && after !== SPACE) ok = true;
      else if (before === '​') ok = true;
      else if (BREAK_AFTER.has(before)) ok = true;
      else if (before === '/' || before === '\\') ok = true;
      else if (isEastAsian(before) || isEastAsian(after)) ok = true;
      else if (flags.latinLnBrk !== false) ok = true;
      if (!ok) continue;
      if (kinsoku.noStart.has(after)) continue;
      if (kinsoku.noEnd.has(before)) continue;
      out.add(k);
    }
    return out;
  },
  hangCount: (points, k) => trailingSpaces(points, k),
  emergency: true,
};

/** The measured opportunities, but a trailing space is charged to the fit test. */
const NO_HANG: Model = {
  name: 'no-hanging',
  opportunities: measuredOpportunities,
  rendersHyphen: (points, k) => MEASURED.rendersHyphen(points, k),
  hangCount: () => 0,
  emergency: true,
};

/** The measured rule with no emergency break: an unbreakable word overflows. */
const NO_EMERGENCY: Model = { ...MEASURED, name: 'no-emergency', emergency: false };

/** Break only after spaces. What a first implementation does. */
const SPACES_ONLY: Model = {
  name: 'spaces-only',
  rendersHyphen: (points, k) => MEASURED.rendersHyphen(points, k),
  opportunities(points) {
    const out = new Set<number>();
    for (let k = 1; k < points.length; k += 1) {
      if (points[k - 1] === SPACE && points[k] !== SPACE) out.add(k);
    }
    return out;
  },
  hangCount: (points, k) => trailingSpaces(points, k),
  emergency: true,
};

/** The measured rule with kinsoku switched off entirely. */
const NO_KINSOKU: Model = {
  name: 'no-kinsoku',
  rendersHyphen: (points, k) => MEASURED.rendersHyphen(points, k),
  opportunities: (points, flags) => measuredOpportunities(points, flags, EMPTY_KINSOKU),
  hangCount: (points, k, flags) => MEASURED.hangCount(points, k, flags),
  emergency: true,
};

/**
 * The measured rule with the kinsoku filter applied to *every* opportunity
 * rather than only to East Asian ones.
 *
 * This is the reading ECMA supports - `p:kinsoku` says nothing about script -
 * and the one a renderer would write from the specification alone. It is
 * refuted by `latin-custom`, which forbids `b` at the start of a line and `a`
 * at the end of one and changes nothing about English text.
 */
const KINSOKU_UNGATED: Model = {
  name: 'kinsoku-ungated',
  rendersHyphen: (points, k) => MEASURED.rendersHyphen(points, k),
  hangCount: (points, k, flags) => MEASURED.hangCount(points, k, flags),
  emergency: true,
  opportunities(points, flags, kinsoku) {
    const out = new Set<number>();
    for (const k of measuredOpportunities(points, flags, EMPTY_KINSOKU)) {
      const before = points[k - 1];
      const after = points[k];
      if (before === undefined || after === undefined) continue;
      if (kinsoku.noStart.has(after) || kinsoku.noEnd.has(before)) continue;
      out.add(k);
    }
    return out;
  },
};

// --------------------------------------------------------------- simulation

/**
 * What the model predicts for line 1.
 *
 * The greedy rule: the longest prefix that ends at an opportunity and whose
 * charged width fits. If none does and the model allows an emergency break,
 * exactly as many code points as fit, but never fewer than one - a line that
 * held nothing would not terminate.
 */
function predict(model: Model, c: Case, kinsoku: Kinsoku): number {
  const points = c.string.points;
  const widths = c.string.widths;
  const box = c.probe.widthPt ?? 0;
  const flags = c.probe.flags;
  const opps = model.opportunities(points, flags, kinsoku);

  const hyphen = inputs.hyphenWidth[FACES[c.string.face as keyof typeof FACES]] ?? 0;
  const charged = (k: number): number => {
    const w = widths[k - model.hangCount(points, k, flags)];
    if (w === undefined) return Number.POSITIVE_INFINITY;
    return model.rendersHyphen(points, k) ? w + hyphen : w;
  };

  let best = 0;
  for (const k of opps) {
    if (k > best && charged(k) <= box) best = k;
  }
  if (best > 0) return best;
  if (!model.emergency) return points.length;

  let fit = 1;
  for (let k = 1; k <= points.length; k += 1) {
    const w = widths[k];
    if (w !== undefined && w <= box) fit = k;
  }
  // A forced break still takes the spaces that follow it. Measured on
  // `aa 、bb。cc` in a two-character box, where no opportunity was
  // available and PowerPoint returned three code points rather than two.
  while (fit < points.length && points[fit] === SPACE) fit += 1;
  return fit;
}

/** The kinsoku in force for each deck, given what section E concludes. */
function kinsokuFor(deck: string, builtIn: Kinsoku, stated: Map<string, Kinsoku>): Kinsoku {
  const input = deckInputById.get(deck);
  if (input === undefined) throw new Error(`no deck input for ${deck}`);
  // The file's sets are honoured only when the file both states them and turns
  // the strict list off - section E.
  if (input.kinsoku && input.strictFirstAndLastChars === false) {
    return stated.get(deck) ?? EMPTY_KINSOKU;
  }
  return builtIn;
}

interface Score {
  readonly name: string;
  readonly fits: number;
  readonly total: number;
  readonly misses: readonly string[];
}

function score(
  model: Model,
  subset: readonly Case[],
  builtIn: Kinsoku,
  stated: Map<string, Kinsoku>,
): Score {
  let fits = 0;
  const misses: string[] = [];
  for (const c of subset) {
    const got = predict(model, c, kinsokuFor(c.deck, builtIn, stated));
    if (got === c.observed) fits += 1;
    else if (misses.length < 12) {
      misses.push(
        `${c.deck}/${c.probe.id} box=${String(c.probe.widthPt)}pt ` +
          `predicted ${String(got)} observed ${String(c.observed)} ` +
          `${JSON.stringify(c.lineText)}`,
      );
    }
  }
  return { name: model.name, fits, total: subset.length, misses };
}

// ---------------------------------------------------------------- report

const report: string[] = [];
const findings: Record<string, unknown> = {};

function say(line = ''): void {
  report.push(line);
  console.log(line);
}

function table(rows: readonly Score[]): void {
  const width = Math.max(...rows.map((r) => r.name.length));
  for (const r of rows) {
    say(`    ${r.name.padEnd(width)}  ${String(r.fits).padStart(5)} / ${String(r.total)}`);
  }
}

/** Refuse to emit a fixture unless exactly one candidate is perfect. */
function chooseUnique(what: string, rows: readonly Score[]): Score {
  const perfect = rows.filter((r) => r.fits === r.total);
  if (perfect.length !== 1) {
    say();
    for (const r of rows) {
      if (r.misses.length > 0) {
        say(`  ${r.name} misses:`);
        for (const m of r.misses) say(`    ${m}`);
      }
    }
    throw new Error(
      `${what}: expected exactly one candidate to fit every case, got ${String(perfect.length)}` +
        ` (${perfect.map((r) => r.name).join(', ') || 'none'})`,
    );
  }
  const winner = perfect[0];
  if (winner === undefined) throw new Error('unreachable');
  return winner;
}

say(`T3 - line breaking, measured against PowerPoint`);
say(`  ${String(cases.length)} cut probes, ${String(controls.length)} width controls`);
say(`  chromium ${inputs.chromium}, ${String(inputs.sizePt)}pt`);
say();

// -------------------------------------------------- A: did the decks open

say('A. Packages');
if (repaired.length > 0) {
  throw new Error(
    `repaired on open, so the measurement is of a repaired file: ${repaired.join(', ')}`,
  );
}
for (const r of refused) say(`  REFUSED  ${r}`);
say(`  opened without repair: ${String(readings.decks.filter((d) => d.opened).length)}`);
// `latin-custom-lang` differs from `latin-custom` in one attribute. Its refusal
// is the measurement, not a failure, so it is asserted rather than tolerated.
const langRefusal = refused.find((r) => r.startsWith('latin-custom-lang'));
if (langRefusal === undefined) {
  throw new Error(
    'latin-custom-lang opened. It states p:kinsoku/@lang="en-US" and differs from ' +
      'latin-custom in nothing else; the earlier measurement had PowerPoint refuse it.',
  );
}
findings['kinsokuLangRefused'] = {
  deck: 'latin-custom-lang',
  differsFrom: 'latin-custom',
  inOneAttribute: 'p:kinsoku/@lang',
  value: 'en-US',
  error: langRefusal,
};
say(`  p:kinsoku/@lang="en-US" is refused outright: ${langRefusal}`);
say();

// ------------------------------------------------- F: the built-in kinsoku
// Measured first, because every other section needs it.

say('F. The built-in kinsoku sets');

const sweep = cases.filter((c) => c.string.category === 'kinsoku');
function sweepTable(deck: string): Map<string, Map<number, number>> {
  const out = new Map<string, Map<number, number>>();
  for (const c of sweep) {
    if (c.deck !== deck) continue;
    let m = out.get(c.string.id);
    if (m === undefined) {
      m = new Map();
      out.set(c.string.id, m);
    }
    m.set(c.probe.fitMax ?? -1, c.observed);
  }
  return out;
}
const hangRun = sweepTable('kinsoku');
const nohangRun = sweepTable('kinsoku-nohang');

const noStart = new Set<string>();
const noEnd = new Set<string>();
const hangs = new Set<string>();
const unrestricted: string[] = [];

for (const [id, m] of nohangRun) {
  const str = stringById.get(id);
  if (str === undefined) throw new Error(`no string ${id}`);
  const ch = str.points[3];
  if (ch === undefined) throw new Error(`${id}: no probe character`);
  // The string is three ideographs, the character, then three more. Positions 2
  // and 5 are ideograph boundaries and must be free, or the probe is measuring
  // something other than the character in the middle.
  if (m.get(2) !== 2 || m.get(5) !== 5) {
    throw new Error(
      `${id}: control positions are not free (2 -> ${String(m.get(2))}, 5 -> ${String(m.get(5))})`,
    );
  }
  const startBlocked = m.get(3) !== 3;
  const endBlocked = m.get(4) !== 4;
  if (startBlocked) noStart.add(ch);
  if (endBlocked) noEnd.add(ch);
  if (!startBlocked && !endBlocked) unrestricted.push(ch);
  // Hanging shows up as the character joining line 1 at a box too narrow for
  // it, but only in the run that has hanging switched on.
  if (hangRun.get(id)?.get(3) === 4 && m.get(3) !== 4) hangs.add(ch);
}

// The controls. If any of these came back restricted the sweep is measuring
// something other than kinsoku and every table below is void.
for (const control of ['あ', 'ア', '字', '漢', 'A', '1']) {
  if (noStart.has(control) || noEnd.has(control)) {
    throw new Error(`control character ${JSON.stringify(control)} came back restricted`);
  }
}

const cp = (ch: string): string =>
  `U+${(ch.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}`;
say(`  may not begin a line: ${String(noStart.size)}`);
say(`  may not end a line:   ${String(noEnd.size)}`);
say(`  hang when hangingPunct is on: ${[...hangs].map((c) => `${cp(c)} ${c}`).join('  ')}`);
say(`  unrestricted (incl. all six controls): ${unrestricted.map((c) => cp(c)).join(' ')}`);
const BUILT_IN: Kinsoku = { noStart, noEnd };
if (![...hangs].every((h) => noStart.has(h))) {
  throw new Error('a character hangs but is not start-forbidden, which the model does not allow');
}
say();

// The sets each deck states, as the builder recorded them.
const stated = new Map<string, Kinsoku>();
for (const d of inputs.decks) {
  if (d.invalStChars === null || d.invalEndChars === null) continue;
  stated.set(d.deck, {
    noStart: new Set(codePoints(d.invalStChars)),
    noEnd: new Set(codePoints(d.invalEndChars)),
  });
}

// ------------------------------------------- B/C/D: the rule, scored whole

// ------------------------------------------------ B: the rule, scored whole

/**
 * Thai, held apart from the rest.
 *
 * Thai is written without spaces and PowerPoint segments it with a dictionary:
 * it breaks `สวัสดี` from
 * `ชาวโลก` - two words - and nowhere else. No
 * rule over character classes can produce that, and no JavaScript library
 * carries the dictionary, so this is a gap rather than a miss.
 *
 * It is separated rather than deleted. Section E scores the model over Thai on
 * its own and pins the number, so the gap is a measured quantity that fails the
 * suite if it grows, instead of an absence nobody notices.
 */
const isThai = (c: Case): boolean => c.string.face === 'thai';
const scorable = cases.filter((c) => !isThai(c));
const thaiCases = cases.filter(isThai);

say('B. The break rule, scored over every probe that is not Thai');
const candidates = [
  MEASURED,
  UAX14ISH,
  NO_HANG,
  NO_EMERGENCY,
  SPACES_ONLY,
  NO_KINSOKU,
  KINSOKU_UNGATED,
];
const scored = candidates.map((m) => score(m, scorable, BUILT_IN, stated));
table(scored);
const winner = chooseUnique('the break rule', scored);
say(`  -> ${winner.name}, ${String(winner.fits)}/${String(winner.total)}`);
say();
findings['ruleScores'] = scored.map((s) => ({ model: s.name, fits: s.fits, total: s.total }));

// ------------------------------------------------------- C: the flag defaults

say('C. What an absent attribute means');

/**
 * Two probes of the same string differ only in a flag. If the variant with the
 * attribute absent lays out identically to one of the two stated values and
 * differently from the other, the default is settled by observation.
 */
function compareVariants(stringId: string, deck: string): Map<number, Map<number, number>> {
  const byVariant = new Map<number, Map<number, number>>();
  for (const c of cases) {
    if (c.deck !== deck || c.probe.stringId !== stringId) continue;
    let m = byVariant.get(c.probe.variant);
    if (m === undefined) {
      m = new Map();
      byVariant.set(c.probe.variant, m);
    }
    m.set(c.probe.fitMax ?? -1, c.observed);
  }
  return byVariant;
}
function sameLayout(
  a: Map<number, number> | undefined,
  b: Map<number, number> | undefined,
): boolean {
  if (a === undefined || b === undefined) return false;
  if (a.size !== b.size) return false;
  for (const [k, v] of a) if (b.get(k) !== v) return false;
  return true;
}

/** variant 0 is the attribute absent, 1 is `="1"`, 2 is `="0"`. */
function defaultOf(stringId: string, deck: string, attribute: string): string {
  const v = compareVariants(stringId, deck);
  const absent = v.get(0);
  const asTrue = sameLayout(absent, v.get(1));
  const asFalse = sameLayout(absent, v.get(2));
  if (asTrue && asFalse) {
    say(`  ${attribute}: no effect at all - all three variants lay out identically`);
    return 'inert';
  }
  if (asTrue) {
    say(`  ${attribute}: absent behaves as "1"`);
    return 'true';
  }
  if (asFalse) {
    say(`  ${attribute}: absent behaves as "0"`);
    return 'false';
  }
  throw new Error(`${attribute}: absent matches neither stated value on ${stringId}`);
}

const latinLnBrkDefault = defaultOf('un-longpair', 'latin', 'latinLnBrk');
if (latinLnBrkDefault !== 'false') {
  throw new Error(
    `latinLnBrk came back as ${latinLnBrkDefault}. ECMA gives it a default of true and the plan ` +
      'claims Office behaves as false; a third answer means the probe is wrong.',
  );
}
const hangingPunctDefault = defaultOf('ea-kinsoku-start', 'ea-kinsoku', 'hangingPunct');
if (hangingPunctDefault !== 'true') {
  throw new Error(`hangingPunct came back as ${hangingPunctDefault}, not true`);
}
const eaLnBrkDefault = defaultOf('ea-plain', 'ea-kinsoku', 'eaLnBrk');
const hangingPunctLatin = defaultOf('pu-comma', 'latin', 'hangingPunct (Latin comma)');
say();
findings['flagDefaults'] = {
  latinLnBrk: latinLnBrkDefault,
  hangingPunct: hangingPunctDefault,
  eaLnBrk: eaLnBrkDefault,
  hangingPunctOnLatinComma: hangingPunctLatin,
};

// --------------------------------------------- D: where the kinsoku comes from

say('D. Where the kinsoku sets come from');

/** The layout of one deck's shared strings, keyed by string and box. */
function layoutOf(deck: string): Map<string, number> {
  const out = new Map<string, number>();
  for (const c of cases) {
    if (c.deck !== deck || c.probe.variant !== 0) continue;
    out.set(`${c.probe.stringId}/${String(c.probe.fitMax)}`, c.observed);
  }
  return out;
}
function differs(a: string, b: string): number {
  const x = layoutOf(a);
  const y = layoutOf(b);
  let n = 0;
  for (const [k, v] of x) if (y.has(k) && y.get(k) !== v) n += 1;
  return n;
}

const kinsokuComparisons = {
  /** Stating the standard sets against stating none: if these agree, the list is built in. */
  'ea-kinsoku vs ea-nokinsoku': differs('ea-kinsoku', 'ea-nokinsoku'),
  /** Turning strict off while the stated sets equal the built-in ones: nothing to choose. */
  'ea-kinsoku vs ea-strict0': differs('ea-kinsoku', 'ea-strict0'),
  /** Turning strict off with sets no built-in list contains. */
  'ea-kinsoku vs ea-custom': differs('ea-kinsoku', 'ea-custom'),
  /** Turning strict off with empty sets: how much of the built-in list survives. */
  'ea-kinsoku vs ea-empty-nohang': differs('ea-kinsoku', 'ea-empty-nohang'),
  /** A Latin kinsoku on English text, with and without the toggle. */
  'latin vs latin-kinsoku': differs('latin', 'latin-kinsoku'),
  'latin vs latin-custom': differs('latin', 'latin-custom'),
};
for (const [what, n] of Object.entries(kinsokuComparisons)) {
  say(`  ${what.padEnd(34)} ${n === 0 ? 'identical' : `${String(n)} probes differ`}`);
}
if (kinsokuComparisons['ea-kinsoku vs ea-nokinsoku'] !== 0) {
  throw new Error('stating the standard sets changed the layout; the list is not built in');
}
if (kinsokuComparisons['ea-kinsoku vs ea-custom'] === 0) {
  throw new Error('a custom set with strictFirstAndLastChars="0" had no effect');
}
if (kinsokuComparisons['latin vs latin-custom'] !== 0) {
  throw new Error('a Latin kinsoku changed the layout of English text');
}

// With empty sets, strict off and no hanging, every position in every East
// Asian probe must break. That is what refutes a hard core.
const emptyFree = cases.filter((c) => c.deck === 'ea-empty-nohang');
const stillBlocked = emptyFree.filter((c) => c.observed !== c.probe.fitMax);
if (stillBlocked.length > 0) {
  throw new Error(
    `${String(stillBlocked.length)} positions still refuse to break with empty kinsoku sets: ` +
      stillBlocked
        .slice(0, 6)
        .map((c) => `${c.probe.id} -> ${String(c.observed)}`)
        .join(', '),
  );
}
say(`  with empty sets and no hanging, all ${String(emptyFree.length)} East Asian positions break`);
say('  -> the built-in list is used unless the file states sets AND turns strict off,');
say('     in which case the file replaces it entirely. Nothing survives replacement.');
say();
findings['kinsokuSource'] = kinsokuComparisons;

// ---------------------------------------------------------------- E: Thai

say('E. Thai');
const thaiScore = score(MEASURED, thaiCases, BUILT_IN, stated);
const thaiString = stringById.get('sc-thai');
const thaiBreaks = [...new Set(thaiCases.map((c) => c.observed))].sort((a, b) => a - b);
say(`  the rule fits ${String(thaiScore.fits)}/${String(thaiScore.total)} Thai probes`);
say(`  PowerPoint's line-1 lengths: ${thaiBreaks.join(', ')}`);
say('  It segments by dictionary. We do not, and no JavaScript library carries one,');
say('  so Thai over-runs rather than breaking. This is a scope boundary, not a bug.');
say();
findings['thai'] = {
  text: thaiString?.text ?? null,
  fits: thaiScore.fits,
  total: thaiScore.total,
  observedLineOneLengths: thaiBreaks,
  note:
    'PowerPoint breaks Thai at dictionary word boundaries. The rule has no dictionary and ' +
    'produces no opportunity inside a Thai run, so a Thai line over-runs its box. Tracked, ' +
    'not fixed: the count is pinned so the gap fails the suite if it grows.',
};

// ------------------------------------------------ G: Chromium vs PowerPoint

say('G. Chromium against PowerPoint, on the wrap="none" controls');

/**
 * PowerPoint's `BoundWidth` includes a trailing paragraph mark one space
 * advance wide - T2's finding, and the reason a raw comparison reports a
 * disagreement that is entirely an artefact of the reader. The mark is
 * subtracted here using the space advance measured for the same face.
 */
const deltas: { id: string; face: string; pct: number }[] = [];
for (const control of controls) {
  const family = FACES[control.face as keyof typeof FACES];
  const mark = inputs.spaceWidth[family];
  if (mark === undefined) throw new Error(`no space advance measured for ${family}`);
  const corrected = control.powerpoint - mark;
  if (corrected <= 0) continue;
  deltas.push({
    id: control.id,
    face: control.face,
    pct: Math.abs(corrected - control.chromium) / corrected,
  });
}
deltas.sort((a, b) => a.pct - b.pct);
const quantile = (q: number): number =>
  deltas[Math.min(deltas.length - 1, Math.floor(q * deltas.length))]?.pct ?? 0;
const agreement = {
  comparisons: deltas.length,
  medianPct: quantile(0.5) * 100,
  p95Pct: quantile(0.95) * 100,
  maxPct: (deltas.at(-1)?.pct ?? 0) * 100,
  worst: deltas.slice(-3).map((d) => ({ id: d.id, face: d.face, pct: d.pct * 100 })),
};
say(`  ${String(agreement.comparisons)} comparisons`);
say(
  `  median ${agreement.medianPct.toFixed(3)}%  p95 ${agreement.p95Pct.toFixed(3)}%  max ${agreement.maxPct.toFixed(3)}%`,
);
say(`  worst: ${agreement.worst.map((w) => `${w.id} ${w.pct.toFixed(2)}%`).join(', ')}`);
say();
findings['browserAgreement'] = agreement;

// ----------------------------------------------------------- the fixture

/**
 * The readings, in the form a test can re-derive a rule from.
 *
 * One row per measured probe: what the deck stated, how wide the box was, and
 * how many code points PowerPoint put on line 1. A test that reads this and
 * recomputes the rule fails when the implementation is wrong. A test written by
 * reading the implementation passes forever.
 */
interface SweepRow {
  readonly deck: string;
  readonly string: string;
  readonly variant: number;
  readonly flags: BreakFlags;
  /** The box width of each probe in the sweep, in points, ascending. */
  readonly boxPt: number[];
  /** How many code points that box holds. */
  readonly fitMax: number[];
  /** How many PowerPoint put on line 1. The measurement. */
  readonly lineOne: number[];
}

/*
 * Columnar, one row per sweep rather than one object per probe.
 *
 * The obvious shape - an object per reading, carrying the deck, the string, the
 * text, the flags and the line - repeats everything but three numbers 2188
 * times and puts the file over the corpus check's per-file cap. It is also the
 * wrong shape: a sweep *is* a table of box widths against line lengths, and one
 * row per sweep is what a reader compares against another row.
 *
 * The text of line 1 is not carried. It is `string` sliced to `lineOne`, and
 * the reader has already asserted that PowerPoint's own two readings of it -
 * `Lines(1).Length` and `Lines(1).Text` - agree, so a third copy would prove
 * nothing and cost 200 KiB.
 */
const rowByKey = new Map<string, SweepRow>();
for (const c of cases) {
  const key = `${c.deck}/${c.probe.stringId}/${String(c.probe.variant)}`;
  let row = rowByKey.get(key);
  if (row === undefined) {
    row = {
      deck: c.deck,
      string: c.probe.stringId,
      variant: c.probe.variant,
      flags: c.probe.flags,
      boxPt: [],
      fitMax: [],
      lineOne: [],
    };
    rowByKey.set(key, row);
  }
  row.boxPt.push(c.probe.widthPt ?? 0);
  row.fitMax.push(c.probe.fitMax ?? 0);
  row.lineOne.push(c.observed);
}
const rows = [...rowByKey.values()];

const fixture = {
  $schema: 'https://pptx-studio.dev/ground-truth/line-breaks.schema.json',
  experiment: 'T3',
  subPhase: '3.3',
  description:
    'Where PowerPoint breaks a line. Measured by sizing a box to hold exactly n code points ' +
    'and reading back the text of line 1, which is the break position itself rather than a ' +
    'number a rule has to be inferred from.',
  measuredWith: {
    application: 'Microsoft PowerPoint, via TextRange2 over COM',
    chromium: inputs.chromium,
    fontSizePt: inputs.sizePt,
    faces: FACES,
    probeSizeHundredths: PROBE_SIZE,
  },
  decks: inputs.decks.map((d) => ({
    deck: d.deck,
    statesKinsoku: d.kinsoku,
    invalStChars: d.invalStChars,
    invalEndChars: d.invalEndChars,
    strictFirstAndLastChars: d.strictFirstAndLastChars,
    flagOverride: d.flagOverride,
    opened: readings.decks.find((r) => r.deck === d.deck)?.opened ?? false,
  })),
  /**
   * Every probe string, with the prefix widths that sized its boxes.
   *
   * The widths are here so a test can replay the whole experiment without a
   * browser: a measurer built from `widths` and the box recorded on each
   * reading reproduces exactly the geometry PowerPoint was shown.
   */
  strings: inputs.strings.map((s) => ({
    id: s.id,
    category: s.category,
    face: s.face,
    text: s.text,
    note: s.note,
    // Rounded to a thousandth of a point. The measurement is nowhere near that
    // precise, and the extra digits are float noise that costs 200 KiB.
    widths: s.widths.map((w) => Math.round(w * 1000) / 1000),
  })),
  /** Per face, in points: the hyphen a soft-hyphen break draws, and the space. */
  faceMetrics: Object.fromEntries(
    Object.entries(FACES).map(([key, family]) => [
      key,
      {
        family,
        hyphenPt: inputs.hyphenWidth[family] ?? 0,
        spacePt: inputs.spaceWidth[family] ?? 0,
      },
    ]),
  ),
  rules: findings,
  kinsoku: {
    note:
      "PowerPoint's built-in sets, measured one character at a time. Used unless the file " +
      'both states its own sets and sets strictFirstAndLastChars="0", in which case the ' +
      "file's sets replace these entirely - nothing survives replacement.",
    mayNotBeginLine: [...noStart].sort(),
    mayNotEndLine: [...noEnd].sort(),
    hangOutsideTheMeasure: [...hangs].sort(),
    unrestrictedControls: unrestricted,
    disagreesWithA10Comment: [...['¢', '£', '¥', '＃']].filter(
      (c) => !noStart.has(c) && !noEnd.has(c),
    ),
  },
  readings: rows,
};

const fixturePath = join(work, '..', '..', '..', 'line-breaks.json');
const repoFixture = resolve(process.cwd(), 'corpus', 'ground-truth', 'line-breaks.json');
writeFileSync(repoFixture, `${JSON.stringify(fixture, null, 2)}\n`, 'utf8');
say(`wrote ${repoFixture}`);
say(
  `  ${String(cases.length)} readings in ${String(rows.length)} sweeps, ` +
    `${String(allStrings().length)} strings, ${String(inputs.decks.length)} decks`,
);
void fixturePath;

writeFileSync(join(work, 'break-report.txt'), report.join('\n'), 'utf8');
console.log(`wrote ${join(work, 'break-report.txt')}`);

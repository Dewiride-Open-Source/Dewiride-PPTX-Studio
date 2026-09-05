/**
 * Experiment T5, step 4 - score the candidate rules and emit the fixture.
 *
 * ```
 * node tools/ground-truth/text/bullets/analyse.ts <work-dir> --fixture corpus/ground-truth/bullets.json
 * ```
 *
 * Imports nothing from `@pptx-studio/text`. Every model scored here is written
 * out again, independently, in this file - so a test that replays the fixture is
 * comparing the implementation against a measurement rather than against itself.
 *
 * ## What is scored
 *
 * Each section enumerates the readings a person would plausibly implement,
 * including the wrong one, and `chooseUnique` throws unless exactly one is
 * perfect. A rule that no probe in the deck can tell from its rival is not a
 * measured rule, and the analysis says so by refusing to emit.
 *
 * ## Three instruments, and they fail differently
 *
 * - **The EMF says what was drawn.** `Slide.Export(path, "EMF")` records
 *   PowerPoint's own GDI calls, so a bullet comes back as the characters it
 *   drew, with the face, size and colour it drew them in.
 * - **COM says where it landed**, and what a field resolved to. `BoundLeft` is
 *   the only instrument for a picture bullet, which draws no text at all, and
 *   `TextRange.Text` is the only one for a field in a complex script, which the
 *   EMF records as shaped glyph ids rather than characters.
 * - **The package says what was asked**, which is how the scheme vocabulary is
 *   checked against what PowerPoint itself wrote.
 *
 * ## Attribution, and why it is not the record index
 *
 * GDI splits a text run at digit and script boundaries: the marker `#1925#`
 * comes back as three records, `"#"`, `"1925"`, `"#"`. So the stream is
 * flattened to one character per entry - each carrying the face, size and colour
 * of the record it came from - and the markers are found in the concatenation.
 * The first version tested each record against `^#(\d+)#$` and lost every field
 * probe silently, because a rendered date splits at every slash.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';

import { FIELD_LANGS, FIELD_TYPES, SCHEMES, SCRIPT_FACES } from './probes.ts';
import { readEmf } from '../../lib/emf.ts';

/* -------------------------------------------------------------------------- */
/* arguments and inputs                                                       */
/* -------------------------------------------------------------------------- */

const args = process.argv.slice(2);
const dir = args[0];
if (dir === undefined)
  throw new Error(
    'usage: tools/ground-truth/text/bullets/analyse.ts <work-dir> [--fixture <path>]',
  );
const fixtureAt = args.indexOf('--fixture');
const fixturePath = fixtureAt < 0 ? undefined : args[fixtureAt + 1];

const readJson = (name: string): unknown =>
  JSON.parse(readFileSync(join(dir, name), 'utf8')) as unknown;

interface InputShape {
  readonly id: string;
  readonly family: string;
  readonly slide: number;
  readonly markers: readonly number[];
  readonly asks: string;
}
interface InputDeck {
  readonly deck: string;
  readonly file: string;
  readonly shapes: readonly InputShape[];
}
interface ReadBox {
  readonly index: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly text: string | null;
}
interface ReadShape {
  readonly id: string;
  readonly slide: number;
  readonly left: number | null;
  readonly text: string | null;
  readonly paragraphs: readonly ReadBox[];
  readonly lines: readonly ReadBox[];
}
interface ReadDeck {
  readonly deck: string;
  /** When the deck was opened and closed, which brackets when its fields were computed. */
  readonly openedAt: string;
  readonly closedAt: string | null;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly emf: readonly { slide: number; file: string | null; error: string | null }[];
  readonly shapes: readonly ReadShape[];
}

const inputs = readJson('bullet-inputs.json') as { decks: readonly InputDeck[] };
const readings = readJson('bullet-readings.json') as { decks: readonly ReadDeck[] };

/* -------------------------------------------------------------------------- */
/* the vocabulary, re-derived from what PowerPoint authored                   */
/* -------------------------------------------------------------------------- */

/** Inflate one entry from a ZIP, using only what `node:zlib` already provides. */
function inflateEntry(zip: Buffer, want: string): string {
  let at = zip.length - 22;
  while (at >= 0 && zip.readUInt32LE(at) !== 0x06054b50) at--;
  if (at < 0) throw new Error('no end-of-central-directory record');
  let entry = zip.readUInt32LE(at + 16);
  const count = zip.readUInt16LE(at + 10);
  for (let i = 0; i < count; i++) {
    const nameLen = zip.readUInt16LE(entry + 28);
    const extraLen = zip.readUInt16LE(entry + 30);
    const commentLen = zip.readUInt16LE(entry + 32);
    const name = zip.subarray(entry + 46, entry + 46 + nameLen).toString('utf8');
    if (name === want) {
      const localAt = zip.readUInt32LE(entry + 42);
      const method = zip.readUInt16LE(entry + 10);
      const compressed = zip.readUInt32LE(entry + 20);
      const start = localAt + 30 + zip.readUInt16LE(localAt + 26) + zip.readUInt16LE(localAt + 28);
      const bytes = zip.subarray(start, start + compressed);
      return method === 0 ? bytes.toString('utf8') : inflateRawSync(bytes).toString('utf8');
    }
    entry += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error(`no ${want} in the package`);
}

/**
 * `SCHEMES` must be what PowerPoint wrote, not what the standard lists.
 *
 * The authoring step set `PpNumberedBulletStyle` 0..47 and saved; 0..40 were
 * accepted and each wrote its own `a:buAutoNum/@type`. Re-reading that here
 * makes the table in `bullets.ts` a measurement with a check on it rather than
 * a transcription somebody could mistype.
 */
const authoredSchemes = new Map<number, string>();
{
  const slide = inflateEntry(readFileSync(join(dir, 'pp-bullets.pptx')), 'ppt/slides/slide1.xml');
  for (const sp of slide.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    const name = /name="num-(\d+)"/.exec(sp[0]);
    const type = /<a:buAutoNum type="([^"]+)"/.exec(sp[0]);
    if (name !== null && type !== null) authoredSchemes.set(Number(name[1]), type[1] ?? '');
  }
}
if (authoredSchemes.size !== SCHEMES.length) {
  throw new Error(
    `PowerPoint authored ${String(authoredSchemes.size)} schemes; SCHEMES has ${String(SCHEMES.length)}`,
  );
}
for (const [style, type] of authoredSchemes) {
  if (SCHEMES[style] !== type) {
    throw new Error(
      `SCHEMES[${String(style)}] is ${String(SCHEMES[style])}, PowerPoint wrote ${type}`,
    );
  }
}

/* -------------------------------------------------------------------------- */
/* segmentation                                                               */
/* -------------------------------------------------------------------------- */

/** One drawn character, carrying the state of the record it came from. */
interface Drawn {
  readonly ch: string;
  readonly face: string | null;
  /** `LOGFONTW.lfHeight`, negative, proportional to the em size. */
  readonly height: number | null;
  readonly colorBgr: number | null;
  readonly glyphIndices: boolean;
}

const drawn = new Map<number, Drawn[]>();
let emfSlides = 0;
const emfErrors: string[] = [];

for (const deck of inputs.decks) {
  const reading = readings.decks.find((d) => d.deck === deck.deck);
  if (reading === undefined || !reading.opened) continue;
  for (const slide of reading.emf) {
    if (slide.file === null) {
      emfErrors.push(`${deck.deck}#${String(slide.slide)}: ${slide.error ?? 'no file'}`);
      continue;
    }
    emfSlides++;
    const emf = readEmf(new Uint8Array(readFileSync(join(dir, slide.file))));
    const chars: Drawn[] = [];
    for (const text of emf.texts) {
      for (const ch of text.text) {
        chars.push({
          ch,
          face: text.font?.face ?? null,
          height: text.font?.height ?? null,
          colorBgr: text.colorBgr,
          glyphIndices: text.glyphIndices,
        });
      }
    }
    const flat = chars.map((c) => c.ch).join('');
    let at = 0;
    for (const found of flat.matchAll(/#(\d+)#/g)) {
      drawn.set(Number(found[1]), chars.slice(at, found.index));
      at = found.index + found[0].length;
    }
  }
}

const probeOf = new Map<string, InputShape & { deck: string }>();
for (const deck of inputs.decks) {
  for (const shape of deck.shapes) probeOf.set(shape.id, { ...shape, deck: deck.deck });
}
const shapeOf = new Map<
  string,
  ReadShape & { deck: string; repaired: boolean; openedAt: string; closedAt: string | null }
>();
for (const deck of readings.decks) {
  for (const shape of deck.shapes) {
    shapeOf.set(shape.id, {
      ...shape,
      deck: deck.deck,
      repaired: deck.repaired === true,
      openedAt: deck.openedAt,
      closedAt: deck.closedAt,
    });
  }
}

const asText = (chars: readonly Drawn[]): string => chars.map((c) => c.ch).join('');
const round = (value: number): number => Math.round(value * 10000) / 10000;

/**
 * The characters drawn for paragraph `index` of probe `id`, minus its body text.
 *
 * A probe paragraph draws its bullet, then its content, then its marker, so the
 * window between markers is `bullet + content`. Returns `null` when the window
 * does not end in the expected body - which is how a probe whose marker was
 * clipped off the slide edge is excluded rather than read as its neighbour.
 */
function bulletChars(id: string, index: number, body: string): Drawn[] | null {
  const probe = probeOf.get(id);
  const marker = probe?.markers[index];
  if (marker === undefined) return null;
  const chars = drawn.get(marker);
  if (chars === undefined) return null;
  if (body !== '' && !asText(chars).endsWith(body)) return null;
  return chars.slice(0, chars.length - body.length);
}

/** The bullet's advance in points: the paragraph's left edge, less the frame's. */
function advance(id: string, index = 0): number | null {
  const shape = shapeOf.get(id);
  if (shape?.left == null) return null;
  const para = shape.paragraphs[index];
  if (para?.left == null) return null;
  return round(para.left - shape.left);
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Model<T> {
  readonly name: string;
  readonly note: string;
  readonly of: T;
}
interface Score {
  readonly name: string;
  readonly note: string;
  readonly fits: number;
  readonly of: number;
}

const allScores: Record<string, Score[]> = {};

/**
 * Score every candidate and insist that exactly one is perfect.
 *
 * The guard T4 used, for the same reason: a section that emits its favourite
 * model without checking that the rivals lost has measured nothing. A tie means
 * the probes do not separate the candidates, which is a defect in the deck
 * rather than a result.
 */
function chooseUnique<T, R>(
  what: string,
  models: readonly Model<T>[],
  rows: readonly R[],
  fits: (model: T, row: R) => boolean,
): Model<T> {
  if (rows.length === 0) throw new Error(`${what}: no rows to score`);
  const scores = models.map((model) => ({
    name: model.name,
    note: model.note,
    fits: rows.filter((row) => fits(model.of, row)).length,
    of: rows.length,
  }));
  allScores[what] = scores;
  console.log(`${what}  (${String(rows.length)} rows)`);
  for (const s of scores) {
    const mark = s.fits === rows.length ? '*' : ' ';
    console.log(`  ${mark} ${s.name.padEnd(26)} ${String(s.fits).padStart(5)}  ${s.note}`);
  }
  const perfect = scores.filter((s) => s.fits === rows.length);
  if (perfect.length === 0) {
    if (process.env['T5_DEBUG'] !== undefined) {
      const first = models[0];
      if (first !== undefined) {
        for (const row of rows.filter((r) => !fits(first.of, r)).slice(0, 12)) {
          console.log(`    MISS ${JSON.stringify(row)}`);
        }
      }
    }
    throw new Error(`${what}: no candidate fits all ${String(rows.length)} rows`);
  }
  if (perfect.length > 1) {
    throw new Error(
      `${what}: ${String(perfect.length)} candidates fit every row ` +
        `(${perfect.map((s) => s.name).join(', ')}); the probes do not separate them`,
    );
  }
  const winner = models.find((m) => m.name === perfect[0]?.name);
  if (winner === undefined) throw new Error(`${what}: winner vanished`);
  return winner;
}

/* -------------------------------------------------------------------------- */
/* section A - the packages                                                   */
/* -------------------------------------------------------------------------- */

const packages = readings.decks.map((d) => ({
  deck: d.deck,
  opened: d.opened,
  repaired: d.repaired === true,
  shapes: d.shapes.length,
}));
const repairedDecks = packages.filter((d) => d.repaired).map((d) => d.deck);
const refusedDecks = packages.filter((d) => !d.opened).map((d) => d.deck);

console.log(
  `A. ${String(packages.length)} packages, ${String(emfSlides)} slides exported; ` +
    `${String(repairedDecks.length)} repaired, ${String(refusedDecks.length)} refused`,
);
if (emfErrors.length > 0) console.log(`   EMF failures: ${emfErrors.join('; ')}`);

/* -------------------------------------------------------------------------- */
/* section B - what each scheme renders                                       */
/* -------------------------------------------------------------------------- */

/**
 * The alphabets, and the three numbers that go with each.
 *
 * All of it is measured. `letters` comes from walking n upwards until the string
 * gains a second character - which is how Thai turns out to use 41 of its 44
 * consonants, skipping three. `wrap` and `bias` come from scoring both against
 * 499 readings apiece, where every alphabet has exactly one pair that fits and
 * the runner-up misses by ten to a hundred.
 *
 * - **`wrap`** is where the counter restarts. 780 for five of the six, which is
 *   26 x 30 and therefore a Latin number that the Thai and Devanagari alphabets
 *   inherited; Hebrew wraps at 392, which no arithmetic on 22 produces.
 * - **`bias`** is an off-by-one that Thai alone has. Its bands begin at 42, 82,
 *   123, 164 - 41k for every band after the second - while Latin's begin at 26k+1
 *   and Devanagari's at 16k+1. Nothing in the shape of the rule predicts it,
 *   which is why it is a measured parameter and not a formula.
 * - **`fillWithLast`** is Hebrew's alone: it writes the *last* letter of the
 *   alphabet for each repeat and then the current one, so 26 is tav-dalet and
 *   not dalet-dalet.
 */
interface AlphabetRule {
  readonly letters: string;
  readonly wrap: number;
  readonly bias: number;
  readonly fillWithLast: boolean;
}

const ALPHABETS: Readonly<Record<string, AlphabetRule>> = {
  alphaLc: { letters: 'abcdefghijklmnopqrstuvwxyz', wrap: 780, bias: 0, fillWithLast: false },
  alphaUc: { letters: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', wrap: 780, bias: 0, fillWithLast: false },
  thaiAlpha: {
    letters: 'กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ',
    wrap: 780,
    bias: 1,
    fillWithLast: false,
  },
  hindiAlpha: { letters: 'अआइईउऊऋऌऍऎएऐऑऒओऔ', wrap: 780, bias: 0, fillWithLast: false },
  hindiAlpha1: {
    letters: 'कखगघङचछजझञटठडढणतथदधनपफबभमयरलळवशषसह',
    wrap: 780,
    bias: 0,
    fillWithLast: false,
  },
  hebrew2: { letters: 'אבגדהוזחטיכלמנסעפצקרשת', wrap: 392, bias: 0, fillWithLast: true },
};

function alphabetic(rule: AlphabetRule, n: number): string {
  const size = rule.letters.length;
  const reduced = ((n - 1) % rule.wrap) + 1;
  // The guard is load-bearing only where the bias is one: without it Thai's
  // forty-first item would be two letters, and it is measured as one.
  const count = reduced <= size ? 1 : Math.floor((reduced - 1 + rule.bias) / size) + 1;
  const letter = rule.letters[(reduced - 1) % size] ?? '';
  const fill = rule.fillWithLast ? (rule.letters[size - 1] ?? '') : letter;
  return fill.repeat(count - 1) + letter;
}

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

function roman(n: number, table: readonly (readonly [number, string])[] = ROMAN): string {
  let left = n;
  let out = '';
  for (const [value, glyph] of table) {
    while (left >= value) {
      out += glyph;
      left -= value;
    }
  }
  return out;
}

/**
 * The East Asian digits, and the one that is not the obvious character.
 *
 * Zero is U+25CB WHITE CIRCLE, not U+3007 IDEOGRAPHIC NUMBER ZERO. The two are
 * indistinguishable on screen in every font that has both, and PowerPoint draws
 * the geometric one: `ea1ChsPlain` at 100 is U+4E00 U+25CB U+25CB. An emitter
 * that reaches for U+3007 - which is what a table of Chinese numerals gives you -
 * is wrong on every hundred, and looks right.
 */
const CJK_DIGITS = '○一二三四五六七八九';
const CJK_TEN = '十';

const cjkPositional = (n: number): string =>
  String(n)
    .split('')
    .map((d) => CJK_DIGITS[Number(d)] ?? '')
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
  return cjkPositional(n);
}

/** Traditional Chinese uses the ten-form only for 10 to 19; 20 is two-circle. */
function cjkTraditional(n: number): string {
  if (n < 10) return CJK_DIGITS[n] ?? '';
  if (n < 20) return CJK_TEN + (n % 10 === 0 ? '' : (CJK_DIGITS[n % 10] ?? ''));
  return cjkPositional(n);
}

const inDigits = (table: string, n: number): string =>
  String(n)
    .split('')
    .map((d) => table[Number(d)] ?? '')
    .join('');

const FULLWIDTH_DIGITS = '０１２３４５６７８９';
const THAI_DIGITS = '๐๑๒๓๔๕๖๗๘๙';
const HINDI_DIGITS = '०१२३४५६७८९';

/** Circled digits U+2460..U+2473 cover 1..20, and above that PowerPoint gives up. */
const circledDb = (n: number): string =>
  n >= 1 && n <= 20 ? String.fromCodePoint(0x245f + n) : String(n);

/**
 * Wingdings circled numbers: white at U+F081, black at U+F08C, ten of each.
 *
 * Above ten they **wrap** rather than falling back to digits - 11 draws the same
 * glyph as 1 and 4000 the same as 10 - which is the opposite of
 * `circleNumDbPlain`, and the reason both are probed to 4000 rather than to 20.
 */
const circledWingdings = (n: number, black: boolean): string =>
  String.fromCodePoint((black ? 0xf08c : 0xf081) + ((n - 1) % 10));

type Numeral = (n: number) => string;

const NUMERALS: Readonly<Record<string, Numeral>> = {
  arabic: (n) => String(n),
  arabicDb: (n) => inDigits(FULLWIDTH_DIGITS, n),
  alphaLc: (n) => alphabetic(ALPHABETS['alphaLc'] as AlphabetRule, n),
  alphaUc: (n) => alphabetic(ALPHABETS['alphaUc'] as AlphabetRule, n),
  romanLc: (n) => roman(n),
  romanUc: (n) => roman(n).toUpperCase(),
  ea1Chs: cjkSimplified,
  ea1Cht: cjkTraditional,
  // Japanese and Korean are positional throughout: ten is one-circle, not the
  // ten glyph. `ea1JpnChsDb` follows them rather than the Chs in its own name.
  ea1JpnKor: cjkPositional,
  ea1JpnChsDb: cjkPositional,
  circleNumDb: circledDb,
  circleNumWdWhite: (n) => circledWingdings(n, false),
  circleNumWdBlack: (n) => circledWingdings(n, true),
  thaiAlpha: (n) => alphabetic(ALPHABETS['thaiAlpha'] as AlphabetRule, n),
  thaiNum: (n) => inDigits(THAI_DIGITS, n),
  hindiAlpha: (n) => alphabetic(ALPHABETS['hindiAlpha'] as AlphabetRule, n),
  hindiAlpha1: (n) => alphabetic(ALPHABETS['hindiAlpha1'] as AlphabetRule, n),
  hindiNum: (n) => inDigits(HINDI_DIGITS, n),
  hebrew2: (n) => alphabetic(ALPHABETS['hebrew2'] as AlphabetRule, n),
};

const FULLWIDTH_PERIOD = '．';

interface SchemeRule {
  readonly numeral: string;
  readonly prefix: string;
  readonly suffix: string;
}

function splitScheme(scheme: string): SchemeRule {
  const decorations: readonly (readonly [string, string, string])[] = [
    ['ParenBoth', '(', ')'],
    ['ParenR', '', ')'],
    ['Period', '', '.'],
    ['Plain', '', ''],
    ['Minus', '', '-'],
  ];
  for (const [tail, prefix, suffix] of decorations) {
    if (scheme.endsWith(tail)) {
      const numeral = scheme.slice(0, scheme.length - tail.length);
      // Both double-byte schemes take the fullwidth stop, a different character
      // from the ASCII one, and one an emitter that concatenates "." gets wrong
      // on every Japanese and Chinese deck.
      const wide = numeral === 'arabicDb' || numeral === 'ea1JpnChsDb';
      return { numeral, prefix, suffix: wide && suffix === '.' ? FULLWIDTH_PERIOD : suffix };
    }
  }
  throw new Error(`no decoration recognised in scheme ${scheme}`);
}

function renderScheme(scheme: string, n: number): string {
  const rule = splitScheme(scheme);
  const numeral = NUMERALS[rule.numeral];
  if (numeral === undefined) throw new Error(`no numeral system for ${rule.numeral}`);
  return rule.prefix + numeral(n) + rule.suffix;
}

interface SchemeRow {
  readonly id: string;
  readonly scheme: string;
  readonly n: number;
  readonly drawn: string;
  readonly advancePt: number | null;
}

const schemeRows: SchemeRow[] = [];
const glyphRows: SchemeRow[] = [];
/**
 * Readings PowerPoint truncated, kept out of the scoring and named in the
 * fixture.
 *
 * A scheme that ends in a full stop draws one; a reading that does not have it
 * is a string PowerPoint stopped emitting part-way, not a shorter answer. Twelve
 * of 5121 are like this - `hindiAlphaPeriod` at 769 to 780, where the bullet is
 * forty-eight Devanagari vowels and the stop after them never appears - and the
 * lengths there wander between 47 and 49 for a value the rule says is constant.
 *
 * The filter is objective and uniform rather than a list of ids: it asks only
 * whether the decoration the scheme's own name promises is present. Nothing else
 * in 5121 readings trips it.
 */
const truncatedRows: SchemeRow[] = [];
for (const id of probeOf.keys()) {
  const match = /^(?:scheme-(?:prop|mono)|count)-([A-Za-z0-9]+)-(\d+)$/.exec(id);
  if (match === null) continue;
  const chars = bulletChars(id, 0, 'X');
  if (chars === null) continue;
  const scheme = match[1] ?? '';
  const row: SchemeRow = {
    id,
    scheme,
    n: Number(match[2]),
    drawn: asText(chars),
    advancePt: advance(id),
  };
  // Glyph runs are classified first. Their records hold glyph ids, so the
  // scheme's own decoration is not a character in them either - testing the
  // suffix first would file all 962 of them as truncations.
  const { suffix } = splitScheme(scheme);
  if (chars.some((c) => c.glyphIndices)) glyphRows.push(row);
  else if (suffix !== '' && !row.drawn.endsWith(suffix)) truncatedRows.push(row);
  else schemeRows.push(row);
}

chooseUnique<(scheme: string, n: number) => string, SchemeRow>(
  'B. what a scheme renders',
  [
    {
      name: 'measured',
      note: 'the tables, wrap points and biases read off the sweep',
      of: renderScheme,
    },
    {
      name: 'no-wrap',
      note: 'alphabetic counts grow without bound, so startAt=900 gives 35 letters',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const alphabet = ALPHABETS[rule.numeral];
        if (alphabet === undefined) return renderScheme(scheme, n);
        return (
          rule.prefix + alphabetic({ ...alphabet, wrap: Number.MAX_SAFE_INTEGER }, n) + rule.suffix
        );
      },
    },
    {
      name: 'wrap-780-everywhere',
      note: 'every alphabet wraps at 780, Hebrew included',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const alphabet = ALPHABETS[rule.numeral];
        if (alphabet === undefined) return renderScheme(scheme, n);
        return rule.prefix + alphabetic({ ...alphabet, wrap: 780 }, n) + rule.suffix;
      },
    },
    {
      name: 'bias-zero-everywhere',
      note: 'no alphabet has the off-by-one, so Thai bands begin at 41k+1 like the rest',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const alphabet = ALPHABETS[rule.numeral];
        if (alphabet === undefined) return renderScheme(scheme, n);
        return rule.prefix + alphabetic({ ...alphabet, bias: 0 }, n) + rule.suffix;
      },
    },
    {
      name: 'bijective',
      note: 'alphabetic numbering is bijective base-26, so 27 is "aa" and 53 is "ba"',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const alphabet = ALPHABETS[rule.numeral];
        if (alphabet === undefined) return renderScheme(scheme, n);
        let left = n;
        let out = '';
        while (left > 0) {
          out = (alphabet.letters[(left - 1) % alphabet.letters.length] ?? '') + out;
          left = Math.floor((left - 1) / alphabet.letters.length);
        }
        return rule.prefix + out + rule.suffix;
      },
    },
    {
      name: 'hebrew-repeats',
      note: 'Hebrew repeats its current letter like Latin, so 26 is dalet-dalet',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const alphabet = ALPHABETS[rule.numeral];
        if (alphabet === undefined || !alphabet.fillWithLast) return renderScheme(scheme, n);
        return rule.prefix + alphabetic({ ...alphabet, fillWithLast: false }, n) + rule.suffix;
      },
    },
    {
      name: 'ascii-stop',
      note: 'the double-byte schemes take an ASCII full stop rather than U+FF0E',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        const numeral = NUMERALS[rule.numeral];
        if (numeral === undefined) return renderScheme(scheme, n);
        return rule.prefix + numeral(n) + (rule.suffix === FULLWIDTH_PERIOD ? '.' : rule.suffix);
      },
    },
    {
      name: 'additive-roman',
      note: 'roman numerals are additive, so 4 is "iiii" and 9 is "viiii"',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        if (rule.numeral !== 'romanLc' && rule.numeral !== 'romanUc')
          return renderScheme(scheme, n);
        const out = roman(
          n,
          ROMAN.filter(([, glyph]) => glyph.length === 1),
        );
        return rule.prefix + (rule.numeral === 'romanUc' ? out.toUpperCase() : out) + rule.suffix;
      },
    },
    {
      name: 'cjk-ideographic-zero',
      note: 'the East Asian zero is U+3007, the character a table of Chinese numerals gives',
      of: (scheme, n) => renderScheme(scheme, n).replace(/○/g, '〇'),
    },
    {
      name: 'cjk-one-ten-rule',
      note: 'all three East Asian families share one ten-form rule',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        if (!rule.numeral.startsWith('ea1')) return renderScheme(scheme, n);
        return rule.prefix + cjkTraditional(n) + rule.suffix;
      },
    },
    {
      name: 'wingdings-no-wrap',
      note: 'the Wingdings circled numbers fall back to digits past ten instead of wrapping',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        if (rule.numeral !== 'circleNumWdWhite' && rule.numeral !== 'circleNumWdBlack') {
          return renderScheme(scheme, n);
        }
        const base = rule.numeral === 'circleNumWdBlack' ? 0xf08c : 0xf081;
        return n >= 1 && n <= 10 ? String.fromCodePoint(base + n - 1) : String(n);
      },
    },
    {
      name: 'circled-past-twenty',
      note: 'circled digits continue past 20 rather than falling back to ASCII',
      of: (scheme, n) => {
        const rule = splitScheme(scheme);
        if (rule.numeral !== 'circleNumDb') return renderScheme(scheme, n);
        return rule.prefix + String.fromCodePoint(0x245f + n) + rule.suffix;
      },
    },
  ],
  schemeRows,
  (model, row) => model(row.scheme, row.n) === row.drawn,
);

/* -------------------------------------------------------------------------- */
/* section C - the numbering pass                                             */
/* -------------------------------------------------------------------------- */

/**
 * What a reader has to count, since the file records nothing.
 *
 * A six-paragraph numbered list PowerPoint authored itself carries no `startAt`
 * anywhere, so every number on every slide is computed at layout time. These
 * probes are the only place the rule is visible.
 */
interface SeqPara {
  readonly level: number;
  /** `null` when the paragraph carries no autonumber at all. */
  readonly scheme: string | null;
  readonly startAt: number;
}
interface SeqRow {
  readonly id: string;
  readonly paras: readonly SeqPara[];
  readonly numbers: readonly (number | null)[];
}

const num = (level: number, scheme: string | null, startAt: number): SeqPara => ({
  level,
  scheme,
  startAt,
});

/** The probe shapes, restated so the analysis does not read the builder's table. */
const SEQUENCES: Readonly<Record<string, readonly SeqPara[]>> = {
  'seq-plain': [0, 0, 0, 0, 0].map(() => num(0, 'arabicPeriod', 1)),
  'seq-startat-first': [7, 1, 1, 1].map((s, i) => num(0, 'arabicPeriod', i === 0 ? s : 1)),
  'seq-startat-every': [7, 7, 7, 7].map((s) => num(0, 'arabicPeriod', s)),
  'seq-startat-third': [1, 1, 20, 1].map((s) => num(0, 'arabicPeriod', s)),
  'seq-buNone-middle': [1, 1, 0, 1, 1].map((on) => num(0, on === 1 ? 'arabicPeriod' : null, 1)),
  'seq-buChar-middle': [1, 1, 0, 1, 1].map((on) => num(0, on === 1 ? 'arabicPeriod' : null, 1)),
  'seq-scheme-change': [
    'arabicPeriod',
    'arabicPeriod',
    'alphaLcPeriod',
    'alphaLcPeriod',
    'arabicPeriod',
  ].map((s) => num(0, s, 1)),
  'seq-level-out-and-back': [0, 0, 1, 1, 0, 0].map((l) => num(l, 'arabicPeriod', 1)),
  'seq-level-twice': [0, 1, 0, 1].map((l) => num(l, 'arabicPeriod', 1)),
  'seq-level-deep': [0, 1, 2, 2, 1, 0].map((l) => num(l, 'arabicPeriod', 1)),
  'seq-empty-para': [0, 0, 0].map(() => num(0, 'arabicPeriod', 1)),
  'seq-hard-break': [0, 0, 0].map(() => num(0, 'arabicPeriod', 1)),
  'seq-startat-zero-then': [1, 1, 1, 1].map((s) => num(0, 'arabicPeriod', s)),
  'seq-second-body-a': [0, 0].map(() => num(0, 'arabicPeriod', 1)),
  'seq-second-body-b': [0, 0].map(() => num(0, 'arabicPeriod', 1)),
};

/** Where a probe's body text is not the default `X`. */
const BODY_TEXT: Readonly<Record<string, readonly string[]>> = {
  'seq-empty-para': ['X', '', 'X'],
  'seq-hard-break': ['X', 'wwww', 'X'],
};

/** Read a drawn `arabicPeriod` or `alphaLcPeriod` bullet back as a number. */
function bulletNumber(text: string): number | null {
  const digits = /^(\d+)\.$/.exec(text);
  if (digits !== null) return Number(digits[1]);
  const letters = /^([a-z]+)\.$/.exec(text);
  if (letters === null) return null;
  const word = letters[1] ?? '';
  return word.charCodeAt(0) - 96 + (word.length - 1) * 26;
}

const seqRows: SeqRow[] = [];
for (const [id, paras] of Object.entries(SEQUENCES)) {
  const numbers = paras.map((_, index) => {
    const chars = bulletChars(id, index, BODY_TEXT[id]?.[index] ?? 'X');
    return chars === null ? null : bulletNumber(asText(chars));
  });
  seqRows.push({ id, paras, numbers });
}

type Numbering = (paras: readonly SeqPara[]) => (number | null)[];

/**
 * Walk back, skipping deeper levels; stop at anything shallower, and at anything
 * at this level that is not the same autonumber run.
 */
const measuredNumbering: Numbering = (paras) =>
  paras.map((para, index) => {
    if (para.scheme === null) return null;
    let count = 0;
    for (let back = index - 1; back >= 0; back--) {
      const other = paras[back];
      if (other === undefined) break;
      if (other.level > para.level) continue;
      if (other.level < para.level) break;
      if (other.scheme !== para.scheme || other.startAt !== para.startAt) break;
      count++;
    }
    return para.startAt + count;
  });

chooseUnique<Numbering, SeqRow>(
  'C. the numbering pass',
  [
    {
      name: 'measured',
      note: 'same level, same scheme and same startAt; deeper levels skipped, shallower breaks',
      of: measuredNumbering,
    },
    {
      name: 'one-counter',
      note: 'one counter per text body, incremented by every numbered paragraph',
      of: (paras) => {
        let next = 0;
        return paras.map((para) => (para.scheme === null ? null : para.startAt + next++));
      },
    },
    {
      name: 'per-level-never-restart',
      note: 'a counter per level that never restarts, so a nested list resumes where it left off',
      of: (paras) => {
        const counters = new Map<number, number>();
        return paras.map((para) => {
          if (para.scheme === null) return null;
          const seen = counters.get(para.level) ?? 0;
          counters.set(para.level, seen + 1);
          return para.startAt + seen;
        });
      },
    },
    {
      name: 'startat-does-not-break',
      note: 'the run is broken by a change of scheme but not by a change of startAt',
      of: (paras) =>
        paras.map((para, index) => {
          if (para.scheme === null) return null;
          let count = 0;
          for (let back = index - 1; back >= 0; back--) {
            const other = paras[back];
            if (other === undefined) break;
            if (other.level > para.level) continue;
            if (other.level < para.level) break;
            if (other.scheme !== para.scheme) break;
            count++;
          }
          return para.startAt + count;
        }),
    },
    {
      name: 'deeper-levels-break',
      note: 'any intervening paragraph breaks the run, deeper ones included',
      of: (paras) =>
        paras.map((para, index) => {
          if (para.scheme === null) return null;
          let count = 0;
          for (let back = index - 1; back >= 0; back--) {
            const other = paras[back];
            if (other === undefined) break;
            if (other.level !== para.level) break;
            if (other.scheme !== para.scheme || other.startAt !== para.startAt) break;
            count++;
          }
          return para.startAt + count;
        }),
    },
    {
      name: 'unnumbered-consumes',
      note: 'an unnumbered paragraph takes a number without showing one',
      of: (paras) =>
        paras.map((para, index) => {
          if (para.scheme === null) return null;
          let count = 0;
          for (let back = index - 1; back >= 0; back--) {
            const other = paras[back];
            if (other === undefined) break;
            if (other.level > para.level) continue;
            if (other.level < para.level) break;
            if (
              other.scheme !== null &&
              (other.scheme !== para.scheme || other.startAt !== para.startAt)
            ) {
              break;
            }
            count++;
          }
          return para.startAt + count;
        }),
    },
  ],
  seqRows,
  (model, row) => JSON.stringify(model(row.paras)) === JSON.stringify(row.numbers),
);

/* -------------------------------------------------------------------------- */
/* section D - the private-use mapping                                        */
/* -------------------------------------------------------------------------- */

/**
 * The faces PowerPoint drew a low code point into the private-use area for.
 *
 * Measured, not a guess about what "symbol font" means: a `buChar` of U+00A7 in
 * Wingdings is drawn as U+F0A7, and the same character in Arial is drawn as
 * U+00A7. The list is what this experiment probed and no more; the general rule
 * is a property of the font's `cmap` - a (3,0) symbol subtable - and reading
 * that is 8.1's job.
 */
const SYMBOL_FACES = ['Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings', 'Symbol'];

interface CharRow {
  readonly id: string;
  readonly char: string;
  readonly font: string | null;
  readonly drawn: string;
  readonly face: string | null;
}

/**
 * What each probe asked for, by code point rather than by literal.
 *
 * Deliberately numeric. A private-use character has no visible shape in an
 * editor, so the same placeholder box in two files is not the same character -
 * and the first version of this table silently disagreed with the builder's
 * about which U+F1xx it meant, which scored as a model failure rather than as
 * the transcription error it was.
 */
const CHAR_PROBES: Readonly<Record<string, { code: number; font: string | null }>> = {
  'char-wd-pua-a7': { code: 0xf0a7, font: 'Wingdings' },
  'char-wd-low-a7': { code: 0x00a7, font: 'Wingdings' },
  'char-wd-pua-6c': { code: 0xf06c, font: 'Wingdings' },
  'char-wd-low-6c': { code: 0x006c, font: 'Wingdings' },
  'char-wd-pua-a8': { code: 0xf0a8, font: 'Wingdings' },
  'char-sym-pua-b7': { code: 0xf0b7, font: 'Symbol' },
  'char-sym-low-b7': { code: 0x00b7, font: 'Symbol' },
  'char-arial-bullet': { code: 0x2022, font: 'Arial' },
  'char-arial-pua-a7': { code: 0xf0a7, font: 'Arial' },
  'char-arial-low-a7': { code: 0x00a7, font: 'Arial' },
  'char-nofont-pua-a7': { code: 0xf0a7, font: null },
  'char-nofont-bullet': { code: 0x2022, font: null },
  'char-wd-dash': { code: 0x002d, font: 'Wingdings' },
  'char-wd2-pua-a2': { code: 0xf0a2, font: 'Wingdings 2' },
  'char-webdings-a4': { code: 0xf0a4, font: 'Webdings' },
  'char-wd-high-2022': { code: 0x2022, font: 'Wingdings' },
  'char-wd-latin-A': { code: 0x0041, font: 'Wingdings' },
  'char-wd-pua-hi': { code: 0xf13a, font: 'Wingdings' },
  'char-sym-alpha': { code: 0x0061, font: 'Symbol' },
  'char-digit': { code: 0x0037, font: 'Arial' },
  'char-missing-glyph': { code: 0x0e01, font: 'Courier New' },
};

const charRows: CharRow[] = [];
for (const [id, spec] of Object.entries(CHAR_PROBES)) {
  const chars = bulletChars(id, 0, 'X');
  if (chars === null) continue;
  charRows.push({
    id,
    char: String.fromCodePoint(spec.code),
    font: spec.font,
    drawn: asText(chars),
    face: chars[0]?.face ?? null,
  });
}

/**
 * The twenty-seven code points where Windows-1252 is not Latin-1.
 *
 * This table is the whole finding. A symbol bullet is not mapped by code point;
 * it is mapped by **ANSI codepage byte**, and U+2022 BULLET - by a distance the
 * most common bullet character there is - is byte 0x95 in Windows-1252 and
 * nothing at all in Latin-1. So `buChar char="•" buFont="Wingdings"` draws
 * U+F095, a filled circle, and an implementation that adds 0xF000 to the code
 * point draws U+F022, an envelope.
 */
const CP1252_HIGH: Readonly<Record<number, number>> = {
  0x20ac: 0x80,
  0x201a: 0x82,
  0x0192: 0x83,
  0x201e: 0x84,
  0x2026: 0x85,
  0x2020: 0x86,
  0x2021: 0x87,
  0x02c6: 0x88,
  0x2030: 0x89,
  0x0160: 0x8a,
  0x2039: 0x8b,
  0x0152: 0x8c,
  0x017d: 0x8e,
  0x2018: 0x91,
  0x2019: 0x92,
  0x201c: 0x93,
  0x201d: 0x94,
  0x2022: 0x95,
  0x2013: 0x96,
  0x2014: 0x97,
  0x02dc: 0x98,
  0x2122: 0x99,
  0x0161: 0x9a,
  0x203a: 0x9b,
  0x0153: 0x9c,
  0x017e: 0x9e,
  0x0178: 0x9f,
};

/** The Windows-1252 byte for a code point, or the low byte where there is none. */
function ansiByte(code: number): number {
  if (code >= 0x20 && code <= 0xff) return code;
  return CP1252_HIGH[code] ?? code & 0xff;
}

type PuaMap = (char: string, font: string | null) => string;

const measuredPua: PuaMap = (char, font) => {
  if (font === null || !SYMBOL_FACES.includes(font)) return char;
  return String.fromCodePoint(0xf000 + ansiByte(char.codePointAt(0) ?? 0));
};

chooseUnique<PuaMap, CharRow>(
  'D. the private-use mapping for buChar',
  [
    {
      name: 'measured',
      note: 'a symbol face maps the ANSI codepage byte into U+F000..U+F0FF; other faces do not',
      of: measuredPua,
    },
    { name: 'never', note: 'the character is drawn exactly as written', of: (char) => char },
    {
      name: 'always',
      note: 'every face maps the low range, so a section sign in Arial becomes U+F0A7',
      of: (char) => {
        const code = char.codePointAt(0) ?? 0;
        return code >= 0x20 && code <= 0xff ? String.fromCodePoint(0xf000 + code) : char;
      },
    },
    {
      name: 'latin-1-only',
      note: 'only U+0020..U+00FF is mapped, so U+2022 is drawn as itself',
      of: (char, font) => {
        if (font === null || !SYMBOL_FACES.includes(font)) return char;
        const code = char.codePointAt(0) ?? 0;
        return code >= 0x20 && code <= 0xff ? String.fromCodePoint(0xf000 + code) : char;
      },
    },
    {
      name: 'low-byte',
      note: 'the code point is masked to its low byte rather than mapped through the codepage',
      of: (char, font) => {
        if (font === null || !SYMBOL_FACES.includes(font)) return char;
        return String.fromCodePoint(0xf000 + ((char.codePointAt(0) ?? 0) & 0xff));
      },
    },
  ],
  charRows,
  (model, row) => model(row.char, row.font) === row.drawn,
);

/* -------------------------------------------------------------------------- */
/* section E - which face, which size, which colour                           */
/* -------------------------------------------------------------------------- */

interface FontRow {
  readonly id: string;
  readonly kind: 'autonum' | 'char';
  readonly buFont: string | null;
  readonly runFace: string;
  readonly face: string | null;
}

/** `+mj-lt` and `+mn-lt` in these packages, from the theme the builder writes. */
const THEME_FACES: Readonly<Record<string, string>> = { '+mj-lt': 'Georgia', '+mn-lt': 'Tahoma' };

const fontRows: FontRow[] = [];
for (const [label, buFont] of [
  ['Arial', 'Arial'],
  ['CourierNew', 'Courier New'],
  ['TimesNewRoman', 'Times New Roman'],
  ['Wingdings', 'Wingdings'],
  ['+mj-lt', '+mj-lt'],
  ['+mn-lt', '+mn-lt'],
  ['tx', 'tx'],
  ['absent', null],
] as const) {
  for (const kind of ['autonum', 'char'] as const) {
    const id = `font-${kind === 'autonum' ? 'num' : 'char'}-${label}`;
    const chars = bulletChars(id, 0, 'X');
    if (chars === null) continue;
    fontRows.push({ id, kind, buFont, runFace: 'Arial', face: chars[0]?.face ?? null });
  }
}
for (const [id, buFont, runFace] of [
  ['font-tx-run-CourierNew', 'tx', 'Courier New'],
  ['font-tx-run-TimesNewRoman', 'tx', 'Times New Roman'],
  ['font-num-override-CourierNew', 'Wingdings', 'Courier New'],
  ['font-num-override-TimesNewRoman', 'Wingdings', 'Times New Roman'],
] as const) {
  const chars = bulletChars(id, 0, 'X');
  if (chars !== null) {
    fontRows.push({ id, kind: 'autonum', buFont, runFace, face: chars[0]?.face ?? null });
  }
}

chooseUnique<(row: FontRow) => string, FontRow>(
  'E. which face draws the bullet',
  [
    {
      name: 'measured',
      note: 'buFont for a character bullet; the run face for an autonumber, always',
      of: (row) => {
        if (row.kind === 'autonum') return row.runFace;
        if (row.buFont === null || row.buFont === 'tx') return row.runFace;
        return THEME_FACES[row.buFont] ?? row.buFont;
      },
    },
    {
      name: 'buFont-for-both',
      note: 'buFont applies to an autonumber as well, which is what the schema implies',
      of: (row) => {
        if (row.buFont === null || row.buFont === 'tx') return row.runFace;
        return THEME_FACES[row.buFont] ?? row.buFont;
      },
    },
    {
      name: 'run-face-always',
      note: 'buFont is decorative and the run face draws everything',
      of: (row) => row.runFace,
    },
    {
      name: 'no-theme-resolution',
      note: 'a theme reference is used as a literal typeface name',
      of: (row) => {
        if (row.kind === 'autonum') return row.runFace;
        if (row.buFont === null || row.buFont === 'tx') return row.runFace;
        return row.buFont;
      },
    },
  ],
  fontRows,
  (model, row) => model(row) === row.face,
);

/* the size, read out of LOGFONTW.lfHeight */

interface SizeRow {
  readonly id: string;
  /** The first run's `@sz`, hundredths of a point. */
  readonly runSz: number;
  readonly szPct: number | null;
  readonly szPts: number | null;
  readonly szTx: boolean;
  /** `lfHeight`, negative. Proportional to the em size with an unknown constant. */
  readonly height: number;
}

const SIZE_PROBES: readonly (readonly [string, number, number | null, number | null, boolean])[] = [
  ['size-pct-25000', 2400, 25000, null, false],
  ['size-pct-50000', 2400, 50000, null, false],
  ['size-pct-75000', 2400, 75000, null, false],
  ['size-pct-100000', 2400, 100000, null, false],
  ['size-pct-150000', 2400, 150000, null, false],
  ['size-pct-200000', 2400, 200000, null, false],
  ['size-pct-400000', 2400, 400000, null, false],
  ['size-pts-600', 2400, null, 600, false],
  ['size-pts-1200', 2400, null, 1200, false],
  ['size-pts-2400', 2400, null, 2400, false],
  ['size-pts-4800', 2400, null, 4800, false],
  ['size-tx', 2400, null, null, true],
  ['size-pct-vs-defrpr', 2400, 50000, null, false],
  ['size-two-runs-small-first', 800, null, null, false],
  ['size-two-runs-large-first', 4000, null, null, false],
];

const sizeRows: SizeRow[] = [];
for (const [id, runSz, szPct, szPts, szTx] of SIZE_PROBES) {
  const chars = bulletChars(id, 0, id.startsWith('size-two-runs') ? 'aabb' : 'X');
  const height = chars?.[0]?.height;
  if (height == null) continue;
  sizeRows.push({ id, runSz, szPct, szPts, szTx, height });
}

// `lfHeight` is in logical units, and the constant relating it to points is
// whatever the export DC used. It is recovered from the 100% row rather than
// assumed, so the section measures a ratio and not a unit.
const unitRow = sizeRows.find((r) => r.id === 'size-pct-100000');
if (unitRow === undefined) throw new Error('no 100% row to calibrate lfHeight against');
const UNITS_PER_POINT = -unitRow.height / (unitRow.runSz / 100);

chooseUnique<(row: SizeRow) => number, SizeRow>(
  'F. what size the bullet is drawn at',
  [
    {
      name: 'measured',
      note: 'buSzPct is a percentage of the first run size; buSzPts is absolute; else the run size',
      of: (row) => {
        if (row.szPts !== null) return row.szPts / 100;
        if (row.szPct !== null) return ((row.runSz / 100) * row.szPct) / 100000;
        return row.runSz / 100;
      },
    },
    {
      name: 'pct-of-defrpr',
      note: 'buSzPct is a percentage of the paragraph defRPr size where one is stated',
      of: (row) => {
        const runPt = row.id === 'size-pct-vs-defrpr' ? 12 : row.runSz / 100;
        if (row.szPts !== null) return row.szPts / 100;
        if (row.szPct !== null) return (runPt * row.szPct) / 100000;
        return runPt;
      },
    },
    {
      name: 'pts-are-scaled',
      note: 'buSzPts is a percentage too, so 600 means six per cent',
      of: (row) => {
        const runPt = row.runSz / 100;
        if (row.szPts !== null) return (runPt * row.szPts) / 100000;
        if (row.szPct !== null) return (runPt * row.szPct) / 100000;
        return runPt;
      },
    },
    {
      name: 'largest-run',
      note: 'the bullet follows the largest run in the paragraph rather than the first',
      of: (row) => {
        const runPt = row.id.startsWith('size-two-runs') ? 40 : row.runSz / 100;
        if (row.szPts !== null) return row.szPts / 100;
        if (row.szPct !== null) return (runPt * row.szPct) / 100000;
        return runPt;
      },
    },
  ],
  sizeRows,
  (model, row) => Math.abs(model(row) * UNITS_PER_POINT + row.height) < 0.5,
);

/* the colour, read out of EMR_SETTEXTCOLOR */

interface ColourRow {
  readonly id: string;
  readonly buClr: string | null;
  readonly clrTx: boolean;
  readonly runClr: string;
  readonly drawnRgb: string;
}

/** `accent2` in the theme the builder writes. */
const SCHEME_ACCENT2 = 'ED7D31';

const colourRows: ColourRow[] = [];
for (const [id, buClr, clrTx] of [
  ['clr-srgb', 'C00000', false],
  ['clr-scheme', SCHEME_ACCENT2, false],
  ['clr-tx', null, true],
  ['clr-absent', null, false],
  ['clr-num-srgb', 'C00000', false],
  ['clr-num-scheme', SCHEME_ACCENT2, false],
  ['clr-num-absent', null, false],
] as const) {
  const chars = bulletChars(id, 0, 'X');
  const bgr = chars?.[0]?.colorBgr;
  if (bgr == null) continue;
  const rgb = ((bgr & 0xff) << 16) | (bgr & 0xff00) | ((bgr >> 16) & 0xff);
  colourRows.push({
    id,
    buClr,
    clrTx,
    runClr: '1F7A1F',
    drawnRgb: rgb.toString(16).toUpperCase().padStart(6, '0'),
  });
}

chooseUnique<(row: ColourRow) => string, ColourRow>(
  'G. what colour the bullet is drawn in',
  [
    {
      name: 'measured',
      note: 'buClr where stated, whatever the bullet kind; otherwise the run colour',
      of: (row) => (row.buClr !== null && !row.clrTx ? row.buClr : row.runClr),
    },
    {
      name: 'text-colour-always',
      note: 'buClr is decorative and the bullet follows the text, as buFont does for an autonumber',
      of: (row) => row.runClr,
    },
    {
      name: 'char-only',
      note: 'buClr reaches a character bullet but not an autonumber, matching buFont',
      of: (row) =>
        row.buClr !== null && !row.clrTx && !row.id.startsWith('clr-num') ? row.buClr : row.runClr,
    },
    {
      name: 'default-black',
      note: 'a bullet with no buClr is black rather than following the text',
      of: (row) => (row.buClr !== null && !row.clrTx ? row.buClr : '000000'),
    },
  ],
  colourRows,
  (model, row) => model(row) === row.drawnRgb,
);

/* -------------------------------------------------------------------------- */
/* section H - picture bullets                                                */
/* -------------------------------------------------------------------------- */

interface BlipRow {
  readonly id: string;
  readonly aspect: number;
  readonly sizePt: number;
  readonly szPct: number;
  readonly advancePt: number;
}

const blipRows: BlipRow[] = [];
for (const [id, aspect, sizePt, szPct] of [
  ['blip-size-1200', 1, 12, 100000],
  ['blip-size-4800', 1, 48, 100000],
  ['blip-bullet-square-plain', 1, 24, 100000],
  ['blip-bullet-square-50000', 1, 24, 50000],
  ['blip-bullet-square-200000', 1, 24, 200000],
  ['blip-bullet-wide-plain', 4, 24, 100000],
  ['blip-bullet-wide-50000', 4, 24, 50000],
  ['blip-bullet-wide-200000', 4, 24, 200000],
  ['blip-bullet-tall-plain', 0.25, 24, 100000],
  ['blip-bullet-tall-50000', 0.25, 24, 50000],
  ['blip-bullet-tall-200000', 0.25, 24, 200000],
] as const) {
  const value = advance(id);
  if (value !== null) blipRows.push({ id, aspect, sizePt, szPct, advancePt: value });
}

/**
 * The height a picture bullet is scaled to, as a fraction of the font size.
 *
 * One number, not three: a square source at 12, 24 and 48 points advances 8.4,
 * 16.8 and 33.6.
 */
const BLIP_FACTOR = 0.7;

chooseUnique<(row: BlipRow) => number, BlipRow>(
  'H. what box a picture bullet occupies',
  [
    {
      name: 'measured',
      note: `height is ${String(BLIP_FACTOR)} x the font size, width follows the aspect ratio`,
      of: (row) => (BLIP_FACTOR * row.sizePt * row.szPct * row.aspect) / 100000,
    },
    {
      name: 'square',
      note: 'a picture bullet is letterboxed into a square box, so the aspect ratio is ignored',
      of: (row) => (BLIP_FACTOR * row.sizePt * row.szPct) / 100000,
    },
    {
      name: 'full-em',
      note: 'the height is the font size itself rather than seven tenths of it',
      of: (row) => (row.sizePt * row.szPct * row.aspect) / 100000,
    },
    {
      name: 'fixed-height',
      note: 'the height is a constant 16.8pt whatever the font size',
      of: (row) => (16.8 * row.szPct * row.aspect) / 100000,
    },
  ],
  blipRows,
  (model, row) => Math.abs(model(row) - row.advancePt) < 0.05,
);

/* -------------------------------------------------------------------------- */
/* section I - where the bullet and the text go                               */
/* -------------------------------------------------------------------------- */

interface IndentRow {
  readonly id: string;
  readonly marL: number;
  readonly indent: number;
  /** The bullet's own advance, from the marL=indent=0 probe of the same bullet. */
  readonly bulletPt: number;
  readonly textLeftPt: number;
}

const indentRows: IndentRow[] = [];
{
  const baseline = new Map<number, number>();
  for (const startAt of [8, 888]) {
    const value = advance(`ind-${String(startAt)}-0-0`);
    if (value !== null) baseline.set(startAt, value);
  }
  for (const id of probeOf.keys()) {
    const match = /^ind-(8|888)-(\d+)-(-?\d+)$/.exec(id);
    if (match === null) continue;
    const bulletPt = baseline.get(Number(match[1]));
    const textLeftPt = advance(id);
    if (bulletPt === undefined || textLeftPt === null) continue;
    indentRows.push({ id, marL: Number(match[2]), indent: Number(match[3]), bulletPt, textLeftPt });
  }
}

/**
 * Where the bullet is drawn, which is not `marL + indent`.
 *
 * Two things a plausible implementation gets wrong, and each is worth one line:
 *
 * - **A positive `indent` does not move the bullet.** Only a hanging indent
 *   does, so the term is `min(0, indent)` and not `indent`. Eighteen probes at
 *   `indent="+12"` put the text at `marL + bullet`, which is where it would be
 *   at `indent="0"`.
 * - **The bullet is clamped to the frame.** A hanging indent deeper than `marL`
 *   would put it left of the text box; PowerPoint pins it at zero instead, and
 *   the text follows.
 */
const bulletLeft = (marL: number, indent: number): number =>
  Math.max(0, marL + Math.min(0, indent));

chooseUnique<(row: IndentRow) => number, IndentRow>(
  'I. where the text starts on a bulleted first line',
  [
    {
      name: 'measured',
      note: 'max(marL, -indent, bulletLeft + bullet), where a positive indent moves nothing',
      of: (row) => Math.max(row.marL, -row.indent, bulletLeft(row.marL, row.indent) + row.bulletPt),
    },
    {
      name: 'indent-moves-both-ways',
      note: 'a positive indent pushes the bullet right, as a first-line indent would',
      of: (row) =>
        Math.max(row.marL, -row.indent, Math.max(0, row.marL + row.indent) + row.bulletPt),
    },
    {
      name: 'no-clamp',
      note: 'the bullet may be drawn left of the frame, so a hanging indent can pull it negative',
      of: (row) => Math.max(row.marL, row.marL + Math.min(0, row.indent) + row.bulletPt),
    },
    {
      name: 'always-marL',
      note: 'the text always starts at marL and a wide bullet overlaps it',
      of: (row) => row.marL,
    },
    {
      name: 'bullet-abuts',
      note: 'the text always starts at the end of the bullet',
      of: (row) => bulletLeft(row.marL, row.indent) + row.bulletPt,
    },
    {
      name: 'no-hanging-floor',
      note: 'without the -indent term, a hanging indent wider than the bullet is ignored',
      of: (row) => Math.max(row.marL, bulletLeft(row.marL, row.indent) + row.bulletPt),
    },
  ],
  indentRows,
  (model, row) => Math.abs(model(row) - row.textLeftPt) < 0.05,
);

interface WrapRow {
  readonly id: string;
  readonly marL: number;
  readonly indent: number;
  readonly secondLinePt: number;
}

const wrapRows: WrapRow[] = [];
for (const id of probeOf.keys()) {
  const match = /^ind-wrap-(\d+)-(-?\d+)$/.exec(id);
  if (match === null) continue;
  const shape = shapeOf.get(id);
  const second = shape?.lines[1];
  if (shape?.left == null || second?.left == null) continue;
  wrapRows.push({
    id,
    marL: Number(match[1]),
    indent: Number(match[2]),
    secondLinePt: round(second.left - shape.left),
  });
}

chooseUnique<(row: WrapRow) => number, WrapRow>(
  'J. where a wrapped line starts',
  [
    { name: 'measured', note: 'marL exactly, with no bullet and no clamp', of: (row) => row.marL },
    {
      name: 'clamped',
      note: 'the clamp that moves the first line applies to every line',
      of: (row) => row.marL + Math.max(0, -(row.marL + row.indent)),
    },
    {
      name: 'first-line-position',
      note: 'every line lines up under the first line rather than under marL',
      of: (row) => Math.max(row.marL, -row.indent),
    },
  ],
  wrapRows,
  (model, row) => Math.abs(model(row) - row.secondLinePt) < 0.05,
);

/* -------------------------------------------------------------------------- */
/* section K - script runs                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Which of `a:latin`, `a:ea`, `a:cs` and `a:sym` draws each character.
 *
 * The EMF makes this a reading rather than an inference: a run whose characters
 * resolve to different faces comes back as several text records, each naming the
 * face GDI was asked for. Every probe declares all four slots with a different
 * face in each, so the record names the slot outright.
 *
 * Where the chosen slot's face has no glyph, Windows substitutes - Courier New
 * has no Thai, so the Thai probe comes back in Leelawadee UI - so the model is
 * scored against the *slot*, resolved through a substitution table read off the
 * probes rather than assumed.
 */
const SCRIPT_SAMPLES: readonly (readonly [string, string, string])[] = [
  ['latin-basic', 'Az', 'latin'],
  ['latin-1', 'Éñ', 'latin'],
  ['latin-ext-a', 'Łš', 'latin'],
  ['greek', 'Αβ', 'latin'],
  ['cyrillic', 'Дж', 'latin'],
  ['hebrew', 'את', 'cs'],
  ['arabic', 'اب', 'cs'],
  ['thai', 'กข', 'cs'],
  ['devanagari', 'कख', 'cs'],
  ['hiragana', 'あい', 'ea'],
  ['katakana', 'アイ', 'ea'],
  ['han', '一二', 'ea'],
  ['hangul', '가나', 'ea'],
  ['fullwidth', 'ＡＢ', 'ea'],
  ['cjk-punct', '、。', 'ea'],
  ['general-punct', '–—', 'latin'],
  ['currency', '€£', 'latin'],
  ['symbol-math', '∑∞', 'latin'],
  ['digits', '01', 'latin'],
  ['space-punct', '. ', 'latin'],
  ['pua', '', 'sym'],
];

/**
 * What Windows substituted where the declared face had no glyph.
 *
 * Read off the probes, not assumed: the `cs` slot is Courier New in every probe,
 * and Courier New covers neither Thai nor Devanagari, so two of the complex
 * script rows come back in a substitute. Recording the substitution separately
 * keeps "the cs slot was chosen" and "Courier New drew it" from being confused.
 */
const SUBSTITUTIONS: Readonly<Record<string, string>> = {
  'Leelawadee UI': 'Courier New',
  'Nirmala UI': 'Courier New',
  'Malgun Gothic': 'MS Gothic',
};

interface ScriptRow {
  readonly id: string;
  readonly sample: string;
  readonly text: string;
  readonly expectSlot: string;
  readonly faces: readonly string[];
  readonly slots: readonly string[];
}

const scriptRows: ScriptRow[] = [];
for (const [sample, text, slot] of SCRIPT_SAMPLES) {
  const id = `script-all-${sample}`;
  const chars = bulletChars(id, 0, '');
  if (chars === null || chars.length === 0) continue;
  const faces = [...new Set(chars.map((c) => c.face ?? '?'))];
  const slots = faces.map((face) => {
    const resolved = SUBSTITUTIONS[face] ?? face;
    const entry = Object.entries(SCRIPT_FACES).find(([, value]) => value === resolved);
    return entry?.[0] ?? `?${face}`;
  });
  scriptRows.push({ id, sample, text, expectSlot: slot, faces, slots });
}

const inRange = (code: number, from: number, to: number): boolean => code >= from && code <= to;

const measuredScript = (text: string): string => {
  const code = text.codePointAt(0) ?? 0;
  if (inRange(code, 0xe000, 0xf8ff)) return 'sym';
  if (
    inRange(code, 0x0590, 0x05ff) ||
    inRange(code, 0x0600, 0x06ff) ||
    inRange(code, 0x0900, 0x097f) ||
    inRange(code, 0x0e00, 0x0e7f)
  ) {
    return 'cs';
  }
  if (
    inRange(code, 0x1100, 0x11ff) ||
    inRange(code, 0x3000, 0x30ff) ||
    inRange(code, 0x3130, 0x318f) ||
    inRange(code, 0x4e00, 0x9fff) ||
    inRange(code, 0xac00, 0xd7af) ||
    inRange(code, 0xff00, 0xffef)
  ) {
    return 'ea';
  }
  return 'latin';
};

chooseUnique<(text: string) => string, ScriptRow>(
  'K. which script slot draws a character',
  [
    {
      name: 'measured',
      note: 'private use to sym; Hebrew, Arabic, Thai and Devanagari to cs; CJK to ea; else latin',
      of: measuredScript,
    },
    {
      name: 'no-sym',
      note: 'a:sym is never used and the private-use area falls to a:latin',
      of: (text) => (measuredScript(text) === 'sym' ? 'latin' : measuredScript(text)),
    },
    {
      name: 'non-ascii-is-cs',
      note: 'anything above Latin-1 goes to a:cs, which is how the name reads',
      of: (text) => ((text.codePointAt(0) ?? 0) > 0xff ? 'cs' : 'latin'),
    },
    {
      name: 'punct-is-ea',
      note: 'general punctuation and currency follow the East Asian slot',
      of: (text) =>
        inRange(text.codePointAt(0) ?? 0, 0x2000, 0x20ff) ? 'ea' : measuredScript(text),
    },
  ],
  scriptRows,
  (model, row) =>
    model(row.text) === row.expectSlot && row.slots.every((slot) => slot === model(row.text)),
);

/**
 * What a character falls back to when its own slot is not stated.
 *
 * A separate question from which slot claims it, and one the `script-all`
 * family cannot answer because every one of those probes declares all four. The
 * `script-no<slot>` decks drop one slot at a time, so "the ea slot was chosen"
 * and "there was no ea slot and something else drew it" become different
 * readings.
 */
interface FallbackRow {
  readonly id: string;
  readonly sample: string;
  readonly text: string;
  readonly dropped: string;
  readonly slot: string;
  readonly faces: readonly string[];
}

const fallbackRows: FallbackRow[] = [];
for (const dropped of ['ea', 'cs', 'sym'] as const) {
  for (const [sample, text, slot] of SCRIPT_SAMPLES) {
    const id = `script-no${dropped}-${sample}`;
    const chars = bulletChars(id, 0, '');
    if (chars === null || chars.length === 0) continue;
    fallbackRows.push({
      id,
      sample,
      text,
      dropped,
      slot,
      faces: [...new Set(chars.map((c) => c.face ?? '?'))],
    });
  }
}

/**
 * Score only the rows PowerPoint drew in a face the probe actually named.
 *
 * Where the chosen slot's face has no glyph, Windows substitutes, and the
 * substitute is a property of the *pair* rather than of the rule: Georgia
 * without Hebrew becomes Arial, Georgia without Thai becomes Angsana New,
 * Courier New without CJK becomes Yu Gothic. Scoring those rows would be
 * scoring Windows' font-linking table, and both candidate rules predict a
 * substitution for them, so they separate nothing.
 *
 * Twenty-two of the 63 need no substitute, and one of those is decisive on its
 * own: with `a:sym` dropped, the private-use character is drawn in **Georgia** -
 * the latin face, directly, no substitution - which is the whole answer.
 */
const directRows = fallbackRows.filter((row) => {
  const declared = new Set(Object.values(SCRIPT_FACES) as string[]);
  return row.faces.every((face) => declared.has(face));
});

chooseUnique<(row: FallbackRow) => string, FallbackRow>(
  'K2. which face draws a character whose slot is absent',
  [
    {
      name: 'measured',
      note: 'an absent slot falls straight to a:latin',
      of: (row) => (row.slot === row.dropped ? SCRIPT_FACES.latin : facesOfSlot(row.slot)),
    },
    {
      name: 'ea-then-cs',
      note: 'an absent slot tries the other non-latin slots before a:latin',
      of: (row) => {
        if (row.slot !== row.dropped) return facesOfSlot(row.slot);
        return row.dropped === 'ea' ? SCRIPT_FACES.cs : SCRIPT_FACES.ea;
      },
    },
    {
      name: 'no-fallback',
      note: 'a character whose slot is absent is not drawn at all',
      of: (row) => (row.slot === row.dropped ? '(none)' : facesOfSlot(row.slot)),
    },
  ],
  directRows,
  (model, row) => row.faces.every((face) => face === model(row)),
);
console.log(
  `   ${String(directRows.length)} of ${String(fallbackRows.length)} rows needed no font substitution`,
);

function facesOfSlot(slot: string): string {
  const found = Object.entries(SCRIPT_FACES).find(([name]) => name === slot);
  return found?.[1] ?? SCRIPT_FACES.latin;
}

/* -------------------------------------------------------------------------- */
/* section L - fields                                                         */
/* -------------------------------------------------------------------------- */

/**
 * What a field renders, read from COM rather than from the EMF.
 *
 * Four of the sixteen locales are complex scripts, which PowerPoint shapes
 * before handing to GDI, so their EMF records hold glyph ids. `TextRange.Text`
 * gives the characters for every locale, which is the whole point of keeping two
 * instruments.
 */
interface FieldRow {
  readonly id: string;
  readonly type: string;
  readonly lang: string;
  readonly cached: string;
  readonly rendered: string;
}

function fieldText(id: string): string | null {
  const shape = shapeOf.get(id);
  const marker = probeOf.get(id)?.markers[0];
  if (shape?.text == null || marker === undefined) return null;
  const suffix = `#${String(marker)}#`;
  return shape.text.endsWith(suffix) ? shape.text.slice(0, -suffix.length) : null;
}

const fieldRows: FieldRow[] = [];
for (const type of FIELD_TYPES) {
  const rendered = fieldText(`fld-${type}`);
  if (rendered !== null) {
    fieldRows.push({ id: `fld-${type}`, type, lang: 'en-US', cached: '#stale#', rendered });
  }
}
for (const lang of FIELD_LANGS) {
  for (const type of FIELD_TYPES) {
    if (type === 'slidenum') continue;
    const id = `fld-lang-${lang}-${type}`;
    const rendered = fieldText(id);
    if (rendered !== null) fieldRows.push({ id, type, lang, cached: '#stale#', rendered });
  }
}
const unknownField = fieldText('fld-unknown-type');
const slideNumbers = [1, 2, 3, 4, 5].map((slide) => ({
  slide,
  rendered: fieldText(`fld-slidenum-${String(slide)}`),
}));

chooseUnique<(row: FieldRow) => boolean, FieldRow>(
  'L. whether a field shows its cached text',
  [
    {
      name: 'measured',
      note: 'a reserved type is recomputed on open and the cache is never shown',
      of: (row) => row.rendered !== row.cached,
    },
    {
      name: 'cache-wins',
      note: 'the cached a:t is what the reader displays',
      of: (row) => row.rendered === row.cached,
    },
  ],
  fieldRows,
  (model, row) => model(row),
);

const langForms = new Set(fieldRows.filter((r) => r.type === 'datetime1').map((r) => r.rendered))
  .size;

console.log(
  `M. ${String(fieldRows.length)} field readings; datetime1 takes ${String(langForms)} distinct ` +
    `forms across ${String(FIELD_LANGS.length)} locales, so @lang drives the format`,
);
console.log(
  `   slidenum on slides 1..5 rendered ${slideNumbers.map((r) => r.rendered ?? '?').join(', ')}`,
);
console.log(`   an unreserved @type rendered ${JSON.stringify(unknownField)}`);

/* -------------------------------------------------------------------------- */
/* section N - date formats, as patterns                                      */
/* -------------------------------------------------------------------------- */

/**
 * Turn a rendered date back into the pattern that produced it.
 *
 * The plan for 3.5 was `Intl.DateTimeFormat` keyed on `@lang`, and the
 * measurement refutes it: `Intl` reproduces PowerPoint on 25 of 97 readings.
 * What PowerPoint does is format with **Windows' own per-locale patterns**,
 * which is why `datetime3` is "5 September 2026" in one locale and "05/09/26" in
 * the next, and why `datetime12` in German is "12:56 " - the pattern ends in an
 * AM/PM designator that German defines as the empty string.
 *
 * So the table is measured cell by cell. This derives each cell by substituting
 * the components of the instant the deck was read, longest token first, and then
 * **re-renders the pattern and asserts it reproduces the string**. A cell that
 * does not round-trip is not emitted: the Japanese era calendar, the Umm al-Qura
 * and Buddhist years and the Hebrew calendar are all outside a pattern language
 * of this shape, and a wrong pattern is worse than a missing one because nothing
 * would ever prompt anyone to look at it again.
 */
interface DatePattern {
  readonly lang: string;
  readonly type: string;
  readonly pattern: string;
  readonly rendered: string;
}

/**
 * A component's text in one locale, taken from a full date so it is in context.
 *
 * `toLocaleString(lang, { month: 'long' })` gives Russian "сентябрь", the
 * nominative standalone form, and a Russian date says "сентября". `formatToParts`
 * over a whole date gives the form the locale actually uses inside one, which is
 * the string PowerPoint's output holds. Getting this wrong does not fail loudly:
 * the derivation simply leaves the month name as literal text, the round trip
 * passes, and the pattern is silently frozen to one month of one year.
 */
function localePart(
  lang: string,
  at: Date,
  part: 'month' | 'weekday' | 'dayPeriod',
  width: 'long' | 'short',
): string {
  const options: Intl.DateTimeFormatOptions =
    part === 'dayPeriod'
      ? { hour: 'numeric', minute: '2-digit', hour12: true }
      : part === 'month'
        ? { day: 'numeric', month: width, year: 'numeric' }
        : { weekday: width, day: 'numeric', month: 'numeric', year: 'numeric' };
  try {
    const found = new Intl.DateTimeFormat(lang, options)
      .formatToParts(at)
      .find((piece) => piece.type === part);
    const value = found?.value ?? '';
    // A locale that writes its months as numbers - Japanese, Chinese, Korean -
    // has ICU return "9" for the long month name. Letting that stand as MMMM
    // produces a pattern that renders correctly and documents a lie, so a
    // numeric name is no name at all.
    return part === 'dayPeriod' || !/\p{Nd}/u.test(value) ? value : '';
  } catch {
    return '';
  }
}

/**
 * The pattern language, longest token first.
 *
 * `lang` is threaded through because the month and weekday names are the
 * locale's, and the whole point of the derivation is that PowerPoint's own
 * output is the only place those strings can be checked against.
 */
type PatternToken = readonly [string, (at: Date, lang: string) => string];

const PATTERN_TOKENS: readonly PatternToken[] = [
  ['yyyy', (at) => String(at.getFullYear())],
  ['MMMM', (at, lang) => localePart(lang, at, 'month', 'long')],
  ['dddd', (at, lang) => localePart(lang, at, 'weekday', 'long')],
  ['MMM', (at, lang) => localePart(lang, at, 'month', 'short')],
  ['ddd', (at, lang) => localePart(lang, at, 'weekday', 'short')],
  ['yy', (at) => String(at.getFullYear() % 100).padStart(2, '0')],
  ['MM', (at) => String(at.getMonth() + 1).padStart(2, '0')],
  ['dd', (at) => String(at.getDate()).padStart(2, '0')],
  ['HH', (at) => String(at.getHours()).padStart(2, '0')],
  ['hh', (at) => String(at.getHours() % 12 === 0 ? 12 : at.getHours() % 12).padStart(2, '0')],
  ['mm', (at) => String(at.getMinutes()).padStart(2, '0')],
  ['ss', (at) => String(at.getSeconds()).padStart(2, '0')],
  ['M', (at) => String(at.getMonth() + 1)],
  ['d', (at) => String(at.getDate())],
  ['H', (at) => String(at.getHours())],
  ['h', (at) => String(at.getHours() % 12 === 0 ? 12 : at.getHours() % 12)],
  ['m', (at) => String(at.getMinutes())],
  ['s', (at) => String(at.getSeconds())],
  ['tt', (at, lang) => localePart(lang, at, 'dayPeriod', 'long')],
];

/** Render a derived pattern, so the derivation can be checked against its input. */
function renderPattern(pattern: string, at: Date, lang: string): string {
  let out = '';
  for (let index = 0; index < pattern.length;) {
    const token = PATTERN_TOKENS.find(([name]) => pattern.startsWith(name, index));
    if (token === undefined) {
      out += pattern[index] ?? '';
      index++;
      continue;
    }
    out += token[1](at, lang);
    index += token[0].length;
  }
  return out;
}

const DIGIT = /\p{Nd}/u;

/**
 * Where `value` occurs in `text` as a whole run rather than inside a longer one.
 *
 * The round trip alone is not enough, because one instant cannot tell a real
 * component from a digit that happens to match. The Umm al-Qura year 1448
 * derived as `h448` - the hour was 1 o'clock, `h` renders "1", and re-rendering
 * put the 1 back - so the pattern round-tripped perfectly and was nonsense.
 * A numeric token has to line up with a maximal run of digits.
 */
function wholeRunIndex(text: string, value: string): number {
  const numeric = DIGIT.test(value);
  let from = 0;
  for (;;) {
    const at = text.indexOf(value, from);
    if (at < 0) return -1;
    const before = text[at - 1];
    const after = text[at + value.length];
    const bounded =
      !numeric ||
      ((before === undefined || !DIGIT.test(before)) &&
        (after === undefined || !DIGIT.test(after)));
    if (bounded) return at;
    from = at + 1;
  }
}

/**
 * Replace each component of `at` in `rendered` with its token, longest first.
 *
 * Two guards, both paid for by a wrong answer that looked right:
 *
 * - a numeric token must match a **whole** run of digits, so a year inside
 *   another calendar's year cannot be mistaken for an hour;
 * - the finished pattern must hold **no digits at all**, because a Windows date
 *   pattern never does. A leftover digit is a component frozen into the pattern,
 *   which is precisely what happens for the Hijri, Buddhist, Hebrew and Japanese
 *   era calendars - none of which this pattern language can express, and all of
 *   which are better absent than wrong.
 */
function derivePattern(rendered: string, at: Date, lang: string): string | null {
  let pattern = rendered;
  for (const [name, of] of PATTERN_TOKENS) {
    const value = of(at, lang);
    if (value === '') continue;
    const found = wholeRunIndex(pattern, value);
    if (found >= 0) pattern = pattern.slice(0, found) + name + pattern.slice(found + value.length);
  }
  if (DIGIT.test(pattern)) return null;
  if (freezesAComponent(pattern, at, lang)) return null;
  if (hasForeignLetters(pattern)) return null;
  return renderPattern(pattern, at, lang) === rendered ? pattern : null;
}

/** The letters a token is spelled with; any other letter is literal text. */
const TOKEN_LETTERS = new Set([...'yMdHhmst']);
const LETTER = /\p{L}/u;

/**
 * Whether the pattern holds a letter that is not part of a token.
 *
 * The conservative half of the derivation, and it drops cells that are probably
 * right. It is here because ICU and Windows do not agree on locale data and one
 * instant cannot tell the difference: ICU says Korean's afternoon marker is
 * "PM", Windows writes 오후, so the substitution never happens, the
 * marker survives as literal text, and the pattern round-trips perfectly while
 * saying "afternoon" at nine in the morning. `freezesAComponent` catches that
 * only when ICU knows the string, which here it does not.
 *
 * So a literal letter is refused outright. That loses the Chinese year, month
 * and day markers and the Russian abbreviation for "year", which are genuine
 * literals - they are recorded as gaps with this reason, and `renderField`
 * shows the cached text for them, which is a visibly stale date rather than a
 * confidently wrong one. Closing the gap needs either Windows' own NLS tables or
 * a second measurement at a different time of day, and both are outside 3.5.
 */
function hasForeignLetters(pattern: string): boolean {
  return [...pattern].some((ch) => LETTER.test(ch) && !TOKEN_LETTERS.has(ch));
}

/**
 * Whether the pattern still holds a component as literal text.
 *
 * The round trip cannot catch this on its own. Where ICU spells a component
 * differently from Windows - Korean's afternoon marker, a locale's month name -
 * the substitution simply does not happen, the string stays literal, and the
 * pattern re-renders perfectly for the one instant it was derived from and
 * wrongly for every other. Korean's `datetime12` derived as
 * `"오후 h시 mm분"`, which says "afternoon" at nine in the
 * morning.
 *
 * So each literal is checked against every name the locale could have put
 * there: twelve months, seven weekdays, both day periods, in both widths. That
 * keeps the genuine literals - Chinese year, month and day markers, the Russian
 * abbreviation for "year" - and drops the frozen components.
 */
function freezesAComponent(pattern: string, at: Date, lang: string): boolean {
  const names = new Set<string>();
  for (let month = 0; month < 12; month++) {
    const when = new Date(at.getFullYear(), month, 15, at.getHours());
    for (const width of ['long', 'short'] as const) {
      names.add(localePart(lang, when, 'month', width));
    }
  }
  for (let day = 0; day < 7; day++) {
    const when = new Date(at.getFullYear(), at.getMonth(), 1 + day, at.getHours());
    for (const width of ['long', 'short'] as const) {
      names.add(localePart(lang, when, 'weekday', width));
    }
  }
  for (const hour of [9, 21]) {
    const when = new Date(at.getFullYear(), at.getMonth(), at.getDate(), hour);
    names.add(localePart(lang, when, 'dayPeriod', 'long'));
  }
  names.delete('');
  return [...names].some((name) => pattern.includes(name));
}

const datePatterns: DatePattern[] = [];
const patternGaps: { lang: string; type: string; rendered: string }[] = [];
for (const row of fieldRows) {
  const shape = shapeOf.get(row.id);
  if (shape === undefined) continue;
  const from = new Date(shape.openedAt).getTime();
  const to = new Date(shape.closedAt ?? shape.openedAt).getTime();
  let derived: string | null = null;
  // The seconds move while a deck is being read, so every second in the window
  // is tried and the one that round-trips wins. A pattern with no seconds in it
  // matches on the first try.
  for (let at = from; at <= to && derived === null; at += 1000) {
    derived = derivePattern(row.rendered, new Date(at), row.lang);
  }
  if (derived === null)
    patternGaps.push({ lang: row.lang, type: row.type, rendered: row.rendered });
  else {
    datePatterns.push({ lang: row.lang, type: row.type, pattern: derived, rendered: row.rendered });
  }
}

console.log(
  `N. ${String(datePatterns.length)} of ${String(fieldRows.length)} date readings turned back ` +
    `into a pattern; ${String(patternGaps.length)} did not, in ` +
    `${[...new Set(patternGaps.map((g) => g.lang))].join(', ')}`,
);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Collapse a run of four or more identical characters to `char*count`.
 *
 * Mechanical and exactly reversible, and applied to the measurement rather than
 * derived from a model. `alphaLcPeriod` at 755 is thirty `a`s and a full stop;
 * written out and then escaped to `\u` form for the diff, the count sweep alone
 * is most of a megabyte, and the corpus caps a committed file at 512 KiB.
 */
function packRuns(text: string): string {
  if (text.includes('*')) throw new Error(`a reading holds an asterisk: ${JSON.stringify(text)}`);
  let out = '';
  for (let at = 0; at < text.length;) {
    let run = 1;
    while (text[at + run] === text[at]) run++;
    out += run >= 4 ? `${text[at] ?? ''}*${String(run)}` : (text[at] ?? '').repeat(run);
    at += run;
  }
  return out;
}

/** Group the scheme readings by scheme, then by probe source, then by start value. */
function byScheme(
  rows: readonly SchemeRow[],
): Record<string, Record<string, Record<string, string>>> {
  const out: Record<string, Record<string, Record<string, string>>> = {};
  for (const row of rows) {
    const source = row.id.startsWith('count-')
      ? 'c'
      : row.id.startsWith('scheme-mono-')
        ? 'm'
        : 'p';
    const scheme = (out[row.scheme] ??= {});
    const bucket = (scheme[source] ??= {});
    bucket[String(row.n)] = packRuns(row.drawn);
  }
  return out;
}

/** JSON with every non-ASCII character escaped, so a fixture diff is readable. */
function toJson(value: unknown): string {
  return JSON.stringify(value, null, 2).replace(/[-￿]/g, (ch) => {
    const code = ch.codePointAt(0) ?? 0;
    return `\\u${code.toString(16).padStart(4, '0')}`;
  });
}

const fixture = {
  $comment:
    'Measured against Microsoft PowerPoint by experiment T5. Do not hand-edit: regenerate with ' +
    'tools/ground-truth/text/bullets/analyse.ts. See docs/adr/phase-3-text/0031-bullets-fields-and-script-runs.md.',
  measuredOn: 'Microsoft PowerPoint 365, build 16.0.20326, Windows 11',
  method:
    'Slide.Export(path, "EMF") records PowerPoint own GDI calls, so a bullet comes back as the ' +
    'characters it drew with the face, size and colour it drew them in. Paragraphs(i).BoundLeft ' +
    'gives the position in points, which the EMF cannot, and TextRange.Text gives a field value ' +
    'for the locales PowerPoint shapes into glyph ids before drawing.',
  schemes: SCHEMES,
  alphabets: ALPHABETS,
  symbolFaces: SYMBOL_FACES,
  scriptFaces: SCRIPT_FACES,
  substitutions: SUBSTITUTIONS,
  blipHeightFactor: BLIP_FACTOR,
  packages,
  repaired: repairedDecks,
  refused: refusedDecks,
  scores: allScores,
  renderedNote:
    'Every scheme reading, keyed scheme -> source -> startAt, where the source is the face the ' +
    'probe used: p for Arial, m for Courier New, c for the one-at-a-time count sweep. A run of ' +
    'four or more identical characters is written char*count - a mechanical, lossless encoding, ' +
    'not a model: no scheme emits an asterisk, which the analysis asserts. Without it the ' +
    'forty-eight-letter readings alone push this file past the corpus size cap.',
  rendered: byScheme(schemeRows),
  glyphRunSchemes: [...new Set(glyphRows.map((r) => r.scheme))].sort(),
  glyphRunNote:
    'PowerPoint shapes the Arabic alphabetic schemes before drawing them, so their EMF records ' +
    'hold glyph ids rather than characters. The ids are kept - the repeat structure is still ' +
    'readable in them - but they are not compared as strings, because they are not text.',
  glyphRuns: byScheme(glyphRows),
  truncatedNote:
    'Readings where the decoration the scheme name promises never appeared, so PowerPoint stopped ' +
    'emitting the string part-way. Twelve of 6071, all hindiAlphaPeriod between 769 and 780 where ' +
    'the bullet is forty-eight Devanagari vowels. Excluded from the scoring and kept here.',
  truncated: truncatedRows.map((r) => ({ ...r, drawn: packRuns(r.drawn) })),
  sequences: seqRows,
  characters: charRows,
  fonts: fontRows,
  sizes: sizeRows.map((r) => ({ ...r, pointsDrawn: round(-r.height / UNITS_PER_POINT) })),
  colours: colourRows,
  blips: blipRows,
  indents: indentRows,
  wraps: wrapRows,
  scripts: scriptRows,
  scriptFallbacks: fallbackRows,
  fields: fieldRows,
  slideNumbers,
  unknownFieldType: unknownField,
  datePatternNote:
    'Windows NLS patterns, derived from the rendered strings and verified by re-rendering them. ' +
    'A locale whose calendar is not Gregorian - Japanese era, Umm al-Qura, Buddhist, Hebrew - has ' +
    'no cell here, because a pattern language of this shape cannot express one and a wrong ' +
    'pattern is worse than a missing one.',
  datePatterns,
  datePatternGaps: patternGaps,
};

if (fixturePath !== undefined) {
  writeFileSync(fixturePath, `${toJson(fixture)}\n`);
  console.log(`\nwrote ${fixturePath}`);
}

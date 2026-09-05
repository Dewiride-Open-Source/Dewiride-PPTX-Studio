/**
 * Experiment T3, step 2 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/text/line-breaks/build-deck.ts <work-dir>
 * ```
 *
 * Reads `break-widths.json` from step 1, sizes every box, and writes seven
 * packages plus `break-inputs.json` describing what each shape was asked.
 *
 * ## Why seven
 *
 * Most of the questions cannot be asked in one package, because they are
 * settings on `p:presentation` rather than on a shape:
 *
 * | deck            | `p:kinsoku`      | `@strictFirstAndLastChars` | asks                                                        |
 * |-----------------|------------------|----------------------------|-------------------------------------------------------------|
 * | `latin`         | absent           | absent                     | the Latin, number, punctuation and invisible-character rules |
 * | `ea-kinsoku`    | standard ja-JP   | absent                     | East Asian breaking with the file supplying the sets         |
 * | `ea-nokinsoku`  | **absent**       | absent                     | whether the sets have to be in the file at all               |
 * | `ea-strict0`    | standard ja-JP   | **0**                      | the toggle, with the custom sets equal to the standard ones  |
 * | `ea-custom`     | **novel ja-JP**  | **0**                      | the toggle, with sets no built-in list contains              |
 * | `latin-kinsoku` | **Latin chars**  | absent                     | whether the filter is East-Asian-only or character-general   |
 * | `latin-custom`  | **Latin chars**  | **0**                      | the same, with the toggle that is supposed to enable it      |
 *
 * `ea-nokinsoku` is the falsifier for the whole kinsoku design. If a package
 * that states no sets breaks exactly like one that states the standard sets,
 * then PowerPoint is using a built-in list and reading `p:kinsoku` from the file
 * is the wrong implementation - which is the reading a renderer would arrive at
 * by following ECMA and never notice was wrong.
 *
 * `ea-custom` and `latin-custom` exist because the first pass of T3 got this
 * wrong. It compared `strictFirstAndLastChars="0"` against the default while
 * the file's sets *were* the standard sets, so the toggle had nothing to
 * toggle and the deck could only ever report "no difference". A toggle has to
 * be given two different things to choose between before its answer means
 * anything.
 *
 * ## What is held still
 *
 * Everything T2 turned off is turned off again - zero insets, no autofit, top
 * anchor, no bullet, no indent, single spacing, zero paragraph spacing - and
 * `kern="0"` besides, which T2 measured as "never kern". Kerning cannot move a
 * break opportunity, only a width, so removing it removes the only way a box
 * could land on the wrong side of a boundary.
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildSheetPackage, SCHEME_ONE, shape, type SlideSpec } from '../../lib/sheet-pptx.ts';
import {
  FACES,
  PROBE_SIZE,
  allStrings,
  codePoints,
  makeProbes,
  type BreakFlags,
  type Category,
  type FaceKey,
  type PrefixWidths,
  type Probe,
} from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/text/line-breaks/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

interface Widths {
  readonly chromium: string;
  readonly sizePt: number;
  readonly installed: Readonly<Record<string, boolean>>;
  readonly hyphen: Readonly<Record<string, number>>;
  readonly space: Readonly<Record<string, number>>;
  readonly results: readonly { readonly id: string; readonly widths: readonly number[] }[];
}
const widths = JSON.parse(readFileSync(join(outDir, 'break-widths.json'), 'utf8')) as Widths;

const widthById = new Map<string, PrefixWidths>(widths.results.map((r) => [r.id, r.widths]));
const probes = makeProbes((id) => {
  const w = widthById.get(id);
  if (w === undefined) throw new Error(`no measured widths for ${id} - rerun step 1`);
  return w;
});

// ------------------------------------------------------------------ markup

/**
 * The language each face is tagged with.
 *
 * `p:kinsoku/@lang` is `ja-JP`, and whether a run tagged `en-US` is inside its
 * scope is one of the things being measured - so the runs say what they mean
 * rather than all claiming to be English.
 */
const LANGS: Readonly<Record<FaceKey, string>> = {
  latin: 'en-US',
  mono: 'en-US',
  ea: 'ja-JP',
  thai: 'th-TH',
  arabic: 'ar-SA',
};

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function flagAttrs(flags: BreakFlags): string {
  const bit = (name: string, value: boolean | undefined): string =>
    value === undefined ? '' : ` ${name}="${value ? '1' : '0'}"`;
  return (
    bit('eaLnBrk', flags.eaLnBrk) +
    bit('latinLnBrk', flags.latinLnBrk) +
    bit('hangingPunct', flags.hangingPunct)
  );
}

/**
 * `a:pPr`, with every term that could move a break set to a stated value.
 *
 * `CT_TextParagraphProperties` carries `eaLnBrk`, `latinLnBrk` and
 * `hangingPunct` as attributes rather than children, so they sit beside `marL`
 * and the children stay in schema order: `lnSpc`, `spcBef`, `spcAft`, bullet.
 */
function pPrXml(flags: BreakFlags): string {
  return (
    `<a:pPr marL="0" marR="0" indent="0" algn="l"${flagAttrs(flags)}>` +
    '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>' +
    '<a:spcBef><a:spcPts val="0"/></a:spcBef>' +
    '<a:spcAft><a:spcPts val="0"/></a:spcAft>' +
    '<a:buNone/>' +
    '</a:pPr>'
  );
}

function rPrXml(probe: Probe, tag: 'a:rPr' | 'a:endParaRPr'): string {
  const face = escapeXml(FACES[probe.face]);
  const lang = LANGS[probe.face];
  const altLang = probe.face === 'ea' ? ' altLang="ja-JP"' : '';
  return (
    `<${tag} lang="${lang}"${altLang} sz="${String(PROBE_SIZE)}" kern="0" dirty="0">` +
    `<a:latin typeface="${face}"/><a:ea typeface="${face}"/><a:cs typeface="${face}"/>` +
    `</${tag}>`
  );
}

function txBody(probe: Probe): { bodyPr: string; paragraphs: readonly string[] } {
  const wrap = probe.kind === 'natural' ? 'none' : 'square';
  const bodyPr =
    `<a:bodyPr wrap="${wrap}" lIns="0" tIns="0" rIns="0" bIns="0" ` +
    'anchor="t" anchorCtr="0" rtlCol="0"><a:noAutofit/></a:bodyPr>';
  // Several probes lead or trail with a space, and that is the measurement.
  // Without `xml:space` the question never reaches PowerPoint.
  const preserve = probe.text !== probe.text.trim() ? ' xml:space="preserve"' : '';
  const run =
    '<a:r>' + rPrXml(probe, 'a:rPr') + `<a:t${preserve}>${escapeXml(probe.text)}</a:t>` + '</a:r>';
  return { bodyPr, paragraphs: ['<a:p>' + pPrXml(probe.flags) + run + '</a:p>'] };
}

// ------------------------------------------------------------------- decks

/**
 * A `p:kinsoku` as its two sets, rather than as markup.
 *
 * The analysis has to know exactly what each deck stated in order to score the
 * decks that override the built-in list, and re-parsing the element it just
 * wrote would be one more place for the two to drift apart.
 */
interface KinsokuSpec {
  /** `p:kinsoku/@lang`. */
  readonly lang: string;
  /** `@invalStChars`: characters that may not begin a line. */
  readonly st: string;
  /** `@invalEndChars`: characters that may not end a line. */
  readonly en: string;
}

function kinsokuXml(k: KinsokuSpec): string {
  return (
    `<p:kinsoku lang="${k.lang}" ` +
    `invalStChars="${escapeXml(k.st)}" invalEndChars="${escapeXml(k.en)}"/>`
  );
}

/**
 * The standard Japanese kinsoku sets, as `a10-rtl-cjk` states them.
 *
 * Stated here so the decks that override the built-in list have something real
 * to override it with. Note that this list is *not* the measured one: the sweep
 * in `tools/ground-truth/text/line-breaks/analyse.ts` found four characters it gets wrong, which is exactly
 * why the built-in table is measured rather than copied from here.
 *
 * The ASCII apostrophe and double quote are in PowerPoint's own sets and are
 * left out for the same reason they are left out there: carrying them means
 * writing an escaped quote inside an attribute, and `conventions.test.ts`
 * asserts this repository emits only the three escapes PowerPoint was measured
 * using. Neither character appears in any probe string.
 */
const KINSOKU_JA: KinsokuSpec = {
  lang: 'ja-JP',
  st:
    '!%),.:;?]}¢’”‰′″℃、。々' +
    '〉》」』】〕ぁぃぅぇぉっゃゅ' +
    'ょゎ゛゜ゝゞァィゥェォッャュ' +
    'ョヮヵヶ・ーヽヾ！％），．：' +
    '；？］｝｡｣､･ﾞﾟ',
  en: '$([\\{£¥‘“〈《「『【〔＃＄（［｛￡￥',
};

/**
 * A kinsoku naming two Latin letters.
 *
 * `b` may not begin a line and `a` may not end one. Every string in the
 * `latin-kinsoku` deck breaks at a position that one of those two forbids, so
 * if the filter is character-general the opportunity set visibly shrinks, and if
 * it is gated on script nothing moves at all.
 */
const KINSOKU_LATIN: KinsokuSpec = { lang: 'ja-JP', st: 'b', en: 'a' };

/**
 * A kinsoku naming two *ordinary* Japanese characters.
 *
 * The first pass of T3 stated the standard sets in the `strictFirstAndLastChars`
 * deck, which was useless: the custom sets and the built-in ones were the same
 * sets, so the toggle had nothing to toggle. This one names characters no
 * kinsoku list contains - the katakana SU, which occurs mid-word in the test
 * string, and the kanji NICHI, which occurs at its start.
 *
 * That makes two predictions instead of one. If the file's sets are honoured
 * when strict is off, `ea-plain` loses the opportunity before SU and the one
 * after NICHI. And if the file's sets *replace* the built-in list rather than
 * adding to it, `ea-kinsoku-start` gains back the positions before the
 * ideographic comma and full stop, which these sets do not mention.
 */
const KINSOKU_EA_CUSTOM: KinsokuSpec = { lang: 'ja-JP', st: 'ス', en: '日' };

/**
 * A kinsoku naming nothing at all.
 *
 * `ea-custom` showed the file's sets *replacing* the built-in strict list for
 * the corner brackets, the small kana and the prolonged sound mark - but not
 * for the ideographic comma and full stop, which stayed forbidden at the start
 * of a line even though the custom set never mentions them. Two readings fit
 * that: either the replacement leaves a hard core, or those two characters are
 * governed by something other than the kinsoku list.
 *
 * Empty sets separate them. If a line may still not begin with an ideographic
 * comma when the file forbids nothing whatsoever, the rule is not the list.
 */
const KINSOKU_EMPTY: KinsokuSpec = { lang: 'ja-JP', st: '', en: '' };

/**
 * The Latin sets, tagged as English rather than Japanese.
 *
 * `latin-custom` states `lang="ja-JP"` and has no effect on English runs, which
 * leaves two explanations standing: kinsoku is scoped by language, or it is
 * gated on script and never applies to Latin at all. This deck changes only the
 * `@lang`, so whichever way it comes out names one of them.
 */
const KINSOKU_LATIN_EN: KinsokuSpec = { lang: 'en-US', st: 'b', en: 'a' };

interface DeckSpec {
  readonly deck: string;
  /** Which probe strings go in: a category filter, or an explicit list. */
  readonly categories?: readonly Category[] | undefined;
  readonly stringIds?: readonly string[] | undefined;
  /** Keep only the first flag variant, for decks asking a presentation-level question. */
  readonly firstVariantOnly?: boolean | undefined;
  readonly kinsoku?: KinsokuSpec | undefined;
  readonly strictFirstAndLastChars?: boolean | undefined;
  /**
   * Paragraph flags forced on every probe in the deck, replacing whatever the
   * string states.
   *
   * The kinsoku sweep needs this. With hanging punctuation on - which is the
   * default - a character that hangs is *preferred* by the greedy rule whether
   * or not the position before it is forbidden, because the longer prefix fits
   * either way. So the sweep reads the same for "may not begin a line" and for
   * "hangs", on exactly the characters where it matters. Running it again with
   * `hangingPunct="0"` removes the preference, and the difference between the
   * two runs is the hanging set.
   */
  readonly flagOverride?: BreakFlags | undefined;
}

/**
 * The strings the Latin kinsoku decks use.
 *
 * Every one of them breaks at a position that `invalStChars="b"` or
 * `invalEndChars="a"` forbids, so if the filter is character-general the
 * opportunity set visibly shrinks - and `sp-trail` loses its only opportunity
 * altogether, which would be unmissable.
 */
const LATIN_KINSOKU_STRINGS: readonly string[] = ['sp-words', 'sp-trail', 'sp-double', 'pu-comma'];

const LATIN_CATEGORIES: readonly Category[] = [
  'space',
  'hyphen',
  'number',
  'punct',
  'invisible',
  'uri',
  'unbreakable',
];

const DECKS: readonly DeckSpec[] = [
  { deck: 'latin', categories: LATIN_CATEGORIES },
  // No `p:kinsoku` and no `@strictFirstAndLastChars`: the sweep is measuring
  // PowerPoint's *built-in* list, so the file must state nothing that could
  // stand in for it.
  { deck: 'kinsoku', categories: ['kinsoku'] },
  // The same sweep with hanging punctuation off. A character that hangs is
  // preferred by the greedy rule either way, so only this run can tell
  // "may not begin a line" apart from "hangs" - and the difference between
  // the two runs is the hanging set itself.
  { deck: 'kinsoku-nohang', categories: ['kinsoku'], flagOverride: { hangingPunct: false } },
  { deck: 'ea-kinsoku', categories: ['ea', 'script'], kinsoku: KINSOKU_JA },
  { deck: 'ea-nokinsoku', categories: ['ea'], firstVariantOnly: true },
  {
    deck: 'ea-strict0',
    categories: ['ea'],
    firstVariantOnly: true,
    kinsoku: KINSOKU_JA,
    strictFirstAndLastChars: false,
  },
  {
    deck: 'ea-custom',
    categories: ['ea'],
    firstVariantOnly: true,
    kinsoku: KINSOKU_EA_CUSTOM,
    strictFirstAndLastChars: false,
  },
  {
    deck: 'latin-kinsoku',
    stringIds: LATIN_KINSOKU_STRINGS,
    firstVariantOnly: true,
    kinsoku: KINSOKU_LATIN,
  },
  {
    deck: 'ea-empty',
    categories: ['ea'],
    firstVariantOnly: true,
    kinsoku: KINSOKU_EMPTY,
    strictFirstAndLastChars: false,
  },
  // Empty sets *and* no hanging, so a character that still refuses to begin a
  // line is refusing for a reason the file cannot reach.
  {
    deck: 'ea-empty-nohang',
    categories: ['ea'],
    firstVariantOnly: true,
    kinsoku: KINSOKU_EMPTY,
    strictFirstAndLastChars: false,
    flagOverride: { hangingPunct: false },
  },
  {
    deck: 'latin-custom',
    stringIds: LATIN_KINSOKU_STRINGS,
    firstVariantOnly: true,
    kinsoku: KINSOKU_LATIN,
    strictFirstAndLastChars: false,
  },
  {
    deck: 'latin-custom-lang',
    stringIds: LATIN_KINSOKU_STRINGS,
    firstVariantOnly: true,
    kinsoku: KINSOKU_LATIN_EN,
    strictFirstAndLastChars: false,
  },
];

const categoryOf = new Map<string, Category>(allStrings().map((s) => [s.id, s.category]));

function probesFor(spec: DeckSpec): readonly Probe[] {
  return probes.filter((p) => {
    if (spec.firstVariantOnly === true && p.variant !== 0) return false;
    if (spec.stringIds !== undefined) return spec.stringIds.includes(p.stringId);
    const category = categoryOf.get(p.stringId);
    if (category === undefined) throw new Error(`no category for ${p.stringId}`);
    return spec.categories?.includes(category) ?? false;
  });
}

/** Ten shapes to a slide. They overlap - each is measured alone, never rendered together. */
const PER_SLIDE = 10;
const NATURAL_WIDTH = 940;

function slidesFor(list: readonly Probe[], override: BreakFlags | undefined): SlideSpec[] {
  const slides: SlideSpec[] = [];
  for (let i = 0; i < list.length; i += PER_SLIDE) {
    const chunk = list.slice(i, i + PER_SLIDE);
    slides.push({
      layout: 0,
      shapes: chunk.map((probe, j) =>
        shape({
          id: j + 2,
          name: probe.id,
          rect: {
            x: 10,
            y: 10 + j * 52,
            w: probe.kind === 'natural' ? NATURAL_WIDTH : (probe.widthPt ?? 0),
            h: 48,
          },
          fill: '<a:noFill/>',
          line: '<a:ln><a:solidFill><a:srgbClr val="C8C8C8"/></a:solidFill></a:ln>',
          ...txBody(override === undefined ? probe : { ...probe, flags: override }),
        }),
      ),
    });
  }
  return slides;
}

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
  readonly probes: number;
  readonly kinsoku: boolean;
  /** The sets the deck states, so the analysis never re-parses the markup. */
  readonly invalStChars: string | null;
  readonly invalEndChars: string | null;
  /** Paragraph flags forced on every probe in this deck, if any. */
  readonly flagOverride: BreakFlags | null;
  readonly strictFirstAndLastChars: boolean | null;
}

const entries: DeckEntry[] = [];
const emitted = new Set<string>();
const perDeck = new Map<string, readonly Probe[]>();

for (const spec of DECKS) {
  const list = probesFor(spec);
  if (list.length === 0) throw new Error(`deck ${spec.deck} has no probes`);
  for (const p of list) {
    // A probe id must name one shape in one deck: the reader matches by name,
    // and a duplicate would silently overwrite one reading with another deck's
    // answer to a different question.
    const key = `${spec.deck}/${p.id}`;
    if (emitted.has(key)) throw new Error(`duplicate probe ${key}`);
    emitted.add(key);
  }
  perDeck.set(spec.deck, list);
  const slides = slidesFor(list, spec.flagOverride);
  const file = `breaks-${spec.deck}.pptx`;
  writeFileSync(
    join(outDir, file),
    buildSheetPackage({
      themes: [{ scheme: SCHEME_ONE, majorLatin: 'Arial', minorLatin: 'Arial' }],
      masters: [{ theme: 0 }],
      layouts: [{ master: 0 }],
      slides,
      ...(spec.kinsoku === undefined ? {} : { kinsoku: kinsokuXml(spec.kinsoku) }),
      ...(spec.strictFirstAndLastChars === undefined
        ? {}
        : { strictFirstAndLastChars: spec.strictFirstAndLastChars }),
    }),
  );
  entries.push({
    deck: spec.deck,
    file,
    slides: slides.length,
    probes: list.length,
    kinsoku: spec.kinsoku !== undefined,
    invalStChars: spec.kinsoku?.st ?? null,
    invalEndChars: spec.kinsoku?.en ?? null,
    flagOverride: spec.flagOverride ?? null,
    strictFirstAndLastChars: spec.strictFirstAndLastChars ?? null,
  });
}

writeFileSync(
  join(outDir, 'break-inputs.json'),
  JSON.stringify(
    {
      sizePt: PROBE_SIZE / 100,
      chromium: widths.chromium,
      faces: FACES,
      fontsCheck: widths.installed,
      hyphenWidth: widths.hyphen,
      spaceWidth: widths.space,
      strings: allStrings().map((s) => ({
        ...s,
        points: codePoints(s.text),
        widths: widthById.get(s.id) ?? [],
      })),
      decks: entries.map((e) => ({
        ...e,
        probeIds: (perDeck.get(e.deck) ?? []).map((p) => p.id),
      })),
      probes,
    },
    null,
    2,
  ),
);

const total = entries.reduce((n, e) => n + e.probes, 0);
console.log(`${String(total)} probes across ${String(entries.length)} decks`);
for (const e of entries) {
  const k = e.kinsoku ? 'kinsoku' : 'none   ';
  const s = e.strictFirstAndLastChars === null ? 'default' : String(e.strictFirstAndLastChars);
  console.log(
    `  ${e.deck.padEnd(14)} ${String(e.probes).padStart(4)} probes  ` +
      `${String(e.slides).padStart(3)} slides  ${k}  strict=${s}`,
  );
}

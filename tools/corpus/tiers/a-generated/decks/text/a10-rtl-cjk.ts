import { grid, scheme, shape, solidFill } from '../../markup/shapes.ts';
import { bodyPr, para, run, textLine, txBody, type RunProps } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Right-to-left, East Asian, and the four script slots of a single run.
 *
 * This is where `@lang` comes back into the generator, threaded properly. It
 * was added and removed inside one sitting during the recipe refactor because
 * nothing declared it and nothing honoured it; here it is load-bearing, because
 * line breaking, font selection and field formatting all key on it.
 *
 * ## One run, four typefaces
 *
 * `a:rPr` carries `a:latin`, `a:ea`, `a:cs` and `a:sym` at once, and which one
 * applies is decided **per character** by the script that character belongs to,
 * not per run. A run holding "Hello 世界 مرحبا" draws from three of the four
 * simultaneously. Picking one face for the whole run - which is what a naive
 * implementation does - renders two of the three in a fallback and puts the
 * advance widths out by enough to break every line break after it.
 *
 * Slide 1 is that case, deliberately mixed, so a per-character script-run
 * splitter has something that fails visibly when it is not there.
 *
 * ## `a:rtl` on a run is not `@rtl` on a paragraph
 *
 * The paragraph attribute sets base direction and therefore where the first
 * character goes and which way the bullet hangs. The run element marks the run
 * itself as right-to-left for the bidi algorithm. They are different
 * mechanisms, both present here, and `a:rtl` is the **last** child of `a:rPr`
 * before `a:extLst` - experiment E11's question, closed by the schema order
 * table rather than by guessing.
 *
 * ## `p:kinsoku`
 *
 * Japanese line breaking forbids certain characters at the start of a line and
 * others at the end. `p:kinsoku` on `p:presentation` carries the two sets, and
 * the correct implementation is a **backward-walking post-filter** over an
 * already-chosen break, not a rule inside the break search. The deck states the
 * standard Japanese sets so the filter has real data rather than a placeholder.
 *
 * ## Vertical text
 *
 * `vert`, `vert270` and `eaVert` are in scope and are on slide 3.
 * `mongolianVert` and the two WordArt directions are not: they render
 * approximately and are marked non-editable, which is a scope decision worth
 * having a fixture for either way, so they are here too and expected to look
 * wrong rather than to be absent.
 */

const BOX_LINE = '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>';

/**
 * The standard Japanese kinsoku sets.
 *
 * The ASCII apostrophe and double quote are in PowerPoint's own sets and are
 * left out of these, because carrying them would mean writing `&quot;` inside
 * an attribute and `conventions.test.ts` asserts that this repository emits
 * only `&amp;`, `&lt;` and `&gt;` - the three PowerPoint was measured using.
 * Escaping conventions are `a40-unicode`'s probe, not this deck's, and the
 * omission changes nothing about what the backward-walking filter has to do.
 */
const KINSOKU =
  '<p:kinsoku lang="ja-JP"' +
  ' invalStChars="!%),.:;?]}¢’”‰′″℃、。々' +
  '〉》」』】〕ぁぃぅぇぉっゃゅ' +
  'ょゎ゛゜ゝゞァィゥェォッャュ' +
  'ョヮヵヶ・ーヽヾ！％），．：' +
  '；？］｝｡｣､･ﾞﾟ"' +
  ' invalEndChars="$([\\{£¥‘“〈《「『【〔' +
  '＃＄（［｛￡￥"/>';

// ------------------------------------------------ slide 1: four script slots

/** Every run on slide 1 names all four faces, so the splitter has to choose. */
const FOUR_FACES: RunProps = {
  latin: 'Times New Roman',
  ea: 'Yu Gothic',
  cs: 'Arial',
  sym: 'Wingdings',
};

const SCRIPT_ROWS: readonly { readonly label: string; readonly props: RunProps }[] = [
  {
    label: 'Hello 世界 مرحبا — one run, three scripts',
    props: { ...FOUR_FACES, lang: 'en-GB', altLang: 'ja-JP' },
  },
  {
    label: '日本語のテキストです。 lang="ja-JP"',
    props: { ...FOUR_FACES, lang: 'ja-JP' },
  },
  {
    label: '简体中文的文本。 lang="zh-CN"',
    props: { ...FOUR_FACES, lang: 'zh-CN' },
  },
  {
    label: '한국어 텍스트입니다. lang="ko-KR"',
    props: { ...FOUR_FACES, lang: 'ko-KR' },
  },
  {
    label: 'مرحبا بالعالم — a:rtl val="1"',
    props: { ...FOUR_FACES, lang: 'ar-SA', rtl: true },
  },
  {
    label: 'שלום עולם — a:rtl val="1"',
    props: { ...FOUR_FACES, lang: 'he-IL', rtl: true },
  },
  {
    label: 'सत्यमेव जयते — a:cs decides this one',
    props: { ...FOUR_FACES, lang: 'hi-IN' },
  },
  {
    label: 'สวัสดีครับ — Thai, which needs ' + 'dictionary segmentation no JS library provides',
    props: { ...FOUR_FACES, lang: 'th-TH' },
  },
];

function scriptSlots(): string {
  const cell = grid(1, 1);
  const paras = SCRIPT_ROWS.map((row) =>
    para({
      content: run(row.label, {
        ...row.props,
        sz: 1600,
        fill: solidFill(scheme('tx1')),
      }),
    }),
  ).join('');
  return shape({
    id: 10,
    name: 'One run, four typeface slots',
    ...cell(0),
    line: BOX_LINE,
    textBody: txBody({ bodyPr: bodyPr({ wrap: 'square', autofit: '<a:normAutofit/>' }), paras }),
  });
}

// ------------------------------------- slide 2: paragraph direction and kinsoku

function paragraphDirection(): string {
  const cell = grid(2, 2);

  const rtlParagraph = txBody({
    bodyPr: bodyPr({ wrap: 'square', autofit: '<a:normAutofit/>' }),
    paras:
      textLine('pPr/@rtl="1" — base direction, so the text starts on the right', {
        lang: 'en-GB',
        sz: 1200,
        fill: solidFill(scheme('accent3')),
      }) +
      para({
        props: { rtl: true, marL: 342900, indent: -342900 },
        content: run('هذا نص عربي طويل ' + 'بما يكفي ليلتف ' + 'على أكثر من سطر ' + 'واحد.', {
          lang: 'ar-SA',
          cs: 'Arial',
          sz: 1600,
          rtl: true,
          fill: solidFill(scheme('tx1')),
        }),
      }),
  });

  const ltrParagraph = txBody({
    bodyPr: bodyPr({ wrap: 'square', autofit: '<a:normAutofit/>' }),
    paras:
      textLine('pPr/@rtl="0" with an rtl run inside — bidi, not base direction', {
        lang: 'en-GB',
        sz: 1200,
        fill: solidFill(scheme('accent3')),
      }) +
      para({
        props: { rtl: false },
        content:
          run('The Arabic for "hello world" is ', {
            lang: 'en-GB',
            sz: 1600,
            fill: solidFill(scheme('tx1')),
          }) +
          run('مرحبا بالعالم', {
            lang: 'ar-SA',
            cs: 'Arial',
            sz: 1600,
            rtl: true,
            fill: solidFill(scheme('accent1')),
          }) +
          run(', which reads right to left inside a left-to-right sentence.', {
            lang: 'en-GB',
            sz: 1600,
            fill: solidFill(scheme('tx1')),
          }),
      }),
  });

  const kinsokuBox = txBody({
    bodyPr: bodyPr({ wrap: 'square', autofit: '<a:normAutofit/>' }),
    paras:
      textLine('p:kinsoku — no line may begin with 、。) or end with ( 「', {
        lang: 'en-GB',
        sz: 1200,
        fill: solidFill(scheme('accent3')),
      }) +
      para({
        props: { eaLnBrk: true, hangingPunct: true, latinLnBrk: false },
        content: run(
          '日本語の改行は、行頭に句点' +
            'や閉じ括弧が来ないように、' +
            '選ばれた位置を後ろ向きに' +
            '修正します。',
          { lang: 'ja-JP', ea: 'Yu Gothic', sz: 1600, fill: solidFill(scheme('tx1')) },
        ),
      }),
  });

  const latinLnBrkBox = txBody({
    bodyPr: bodyPr({ wrap: 'square', autofit: '<a:normAutofit/>' }),
    paras:
      textLine('latinLnBrk — ECMA says the default is true, Office behaves as false', {
        lang: 'en-GB',
        sz: 1200,
        fill: solidFill(scheme('accent3')),
      }) +
      para({
        props: { latinLnBrk: true },
        content: run(
          'latinLnBrk="1" Supercalifragilisticexpialidocious antidisestablishmentarianism',
          {
            lang: 'en-GB',
            sz: 1600,
            fill: solidFill(scheme('tx1')),
          },
        ),
      }) +
      para({
        props: { latinLnBrk: false },
        content: run(
          'latinLnBrk="0" Supercalifragilisticexpialidocious antidisestablishmentarianism',
          {
            lang: 'en-GB',
            sz: 1600,
            fill: solidFill(scheme('tx1')),
          },
        ),
      }),
  });

  const bodies = [rtlParagraph, ltrParagraph, kinsokuBox, latinLnBrkBox];
  return bodies
    .map((body, index) =>
      shape({
        id: 10 + index,
        name: 'Direction and breaking ' + String(index + 1),
        ...cell(index),
        line: BOX_LINE,
        textBody: body,
      }),
    )
    .join('');
}

// ----------------------------------------------------- slide 3: vertical text

const VERTICAL_MODES: readonly { readonly vert: string; readonly note: string }[] = [
  { vert: 'horz', note: 'the default' },
  { vert: 'vert', note: 'rotate 90 clockwise' },
  { vert: 'vert270', note: 'rotate 90 anticlockwise' },
  { vert: 'eaVert', note: 'East Asian, glyphs upright' },
  { vert: 'mongolianVert', note: 'out of scope; approximate' },
  { vert: 'wordArtVertRtl', note: 'out of scope; approximate' },
];

function verticalText(): string {
  const cell = grid(6, 1);
  return VERTICAL_MODES.map((mode, index) =>
    shape({
      id: 10 + index,
      name: 'vert ' + mode.vert,
      ...cell(index),
      line: BOX_LINE,
      textBody: txBody({
        bodyPr: bodyPr({ vert: mode.vert, wrap: 'square', anchor: 't' }),
        paras:
          textLine(mode.vert, {
            lang: 'en-GB',
            sz: 1200,
            b: true,
            fill: solidFill(scheme('accent1')),
          }) +
          textLine('縦書き Vertical', {
            lang: 'ja-JP',
            ea: 'Yu Gothic',
            sz: 1400,
            fill: solidFill(scheme('tx1')),
          }) +
          textLine(mode.note, { lang: 'en-GB', sz: 900, i: true, fill: solidFill(scheme('tx1')) }),
      }),
    }),
  ).join('');
}

export const a10RtlCjk: ProbeDeck = {
  id: 'a10-rtl-cjk',
  title: 'PPTX Studio corpus: a10 RTL and CJK',
  description:
    'Runs naming a:latin, a:ea, a:cs and a:sym at once so the per-character script splitter has ' +
    'to choose, a:rtl on runs beside pPr/@rtl on paragraphs, p:kinsoku carrying the standard ' +
    'Japanese sets, latinLnBrk stated both ways because ECMA and Office disagree on its default, ' +
    'and all six ST_TextVerticalType values including the two that are out of scope.',
  features: {
    shape: 17,
    placeholder: 6,
    presetGeom: 11,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a10 RTL and CJK',
    presentationTail: KINSOKU,
    slides: [
      { title: 'a10 — one run, four typeface slots', body: scriptSlots() },
      { title: 'a10 — direction, kinsoku, latinLnBrk', body: paragraphDirection() },
      { title: 'a10 — the six vertical modes', body: verticalText() },
    ],
  }),
};

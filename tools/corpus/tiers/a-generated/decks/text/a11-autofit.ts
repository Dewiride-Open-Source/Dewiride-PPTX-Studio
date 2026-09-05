import { grid, scheme, shape, solidFill, srgb } from '../../markup/shapes.ts';
import { bodyPr, normAutofit, para, run, spcPts, textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Autofit, overflow, and the two things that decide a box's height when it
 * looks empty.
 *
 * ## Two modes, and the fixture serves the first
 *
 * A viewer applies the stored `@fontScale` and `@lnSpcReduction` **verbatim**,
 * so a PowerPoint-authored file pixel-matches. An editor discards them and
 * re-runs the ladder, rounding `pt x scale` to whole points before measuring.
 * This deck is the first case: twelve boxes carrying the same text at twelve
 * different stored scales, so a viewer that recomputes instead of obeying is
 * visibly wrong in eleven of them.
 *
 * The scales here span the range PowerPoint's own ladder covers, from 92.5%
 * down to 25%. They are **not** the authoritative step list - sub-phase 3.4
 * pins that against a measurement of what PowerPoint writes for given content,
 * and a fixture that guessed the steps would turn a measurement into a
 * tautology. What this deck is authoritative about is that stored values are
 * obeyed, whatever they are.
 *
 * ## `spAutoFit` resizes the shape, `normAutofit` resizes the text
 *
 * They are opposites and are confused constantly. `spAutoFit` means the height
 * in `a:ext` is a **cached result**: PowerPoint recomputed it when the text
 * last changed and will recompute it again. `normAutofit` leaves the shape
 * alone and shrinks the text into it. `noAutofit` does neither and lets the
 * text spill, which is legal and common.
 *
 * ## `spcFirstLastPara` defaults false, and that is not a rounding error
 *
 * With it unset, the first paragraph's `spcBef` and the last paragraph's
 * `spcAft` are **discarded**. Honouring them anyway pushes every centred box
 * down and every bottom-anchored box up by the amount of the spacing, on every
 * slide in the deck. Slide 3 has the two side by side with 36pt of spacing, so
 * the difference is a centimetre rather than a hairline.
 *
 * ## `a:endParaRPr` is what gives an empty paragraph its height
 *
 * An `a:p` with no runs still occupies a line, and the size of that line comes
 * from `a:endParaRPr` and nowhere else. Drop it and a deliberate blank line
 * between two paragraphs collapses; read it as a run and an empty paragraph
 * grows text that was never there. Slide 3 stacks four empty paragraphs at
 * four sizes between two visible ones.
 */

const BOX_LINE = '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>';

/** Enough words that every scale on slide 1 wraps differently. */
const SPECIMEN =
  'Autofit changes what the text measures to, never what the text says. ' +
  'The same sentence at a different scale wraps in a different place.';

// ------------------------------------------------ slide 1: stored scale values

const STORED_SCALES: readonly { readonly fontScale: number; readonly lnSpcReduction?: number }[] = [
  { fontScale: 100000 },
  { fontScale: 92500 },
  { fontScale: 85000 },
  { fontScale: 77500 },
  { fontScale: 70000 },
  { fontScale: 65000 },
  { fontScale: 60000, lnSpcReduction: 10000 },
  { fontScale: 55000, lnSpcReduction: 10000 },
  { fontScale: 50000, lnSpcReduction: 10000 },
  { fontScale: 40000, lnSpcReduction: 20000 },
  { fontScale: 32000, lnSpcReduction: 20000 },
  { fontScale: 25000, lnSpcReduction: 20000 },
];

function storedScales(): string {
  const cell = grid(4, 3);
  return STORED_SCALES.map((step, index) => {
    const label =
      String(step.fontScale / 1000) +
      '%' +
      (step.lnSpcReduction === undefined
        ? ''
        : ' / -' + String(step.lnSpcReduction / 1000) + '% lines');
    return shape({
      id: 10 + index,
      name: label,
      ...cell(index),
      line: BOX_LINE,
      textBody: txBody({
        bodyPr: bodyPr({
          wrap: 'square',
          lIns: 45720,
          tIns: 27432,
          rIns: 45720,
          bIns: 27432,
          autofit: normAutofit(step.fontScale, step.lnSpcReduction),
        }),
        paras:
          textLine(label, {
            lang: 'en-GB',
            sz: 1000,
            b: true,
            fill: solidFill(scheme('accent1')),
          }) + textLine(SPECIMEN, { lang: 'en-GB', sz: 1800, fill: solidFill(scheme('tx1')) }),
      }),
    });
  }).join('');
}

// ------------------------------------------- slide 2: the three autofit modes

const OVERFLOW_CASES: readonly {
  readonly name: string;
  readonly note: string;
  readonly autofit: string;
  readonly vertOverflow?: string;
}[] = [
  {
    name: 'noAutofit',
    note: 'the text spills past the bottom edge, and that is legal',
    autofit: '<a:noAutofit/>',
  },
  {
    name: 'normAutofit, no attributes',
    note: 'shrink-to-fit is on, but nothing is stored, so an editor computes it',
    autofit: '<a:normAutofit/>',
  },
  {
    name: 'spAutoFit',
    note: 'a:ext/@cy here is a cached result, not an instruction',
    autofit: '<a:spAutoFit/>',
  },
  {
    name: 'vertOverflow overflow',
    note: 'the default: draw it anyway',
    autofit: '<a:noAutofit/>',
    vertOverflow: 'overflow',
  },
  {
    name: 'vertOverflow ellipsis',
    note: 'truncate the last visible line',
    autofit: '<a:noAutofit/>',
    vertOverflow: 'ellipsis',
  },
  {
    name: 'vertOverflow clip',
    note: 'cut at the box edge, mid-glyph if need be',
    autofit: '<a:noAutofit/>',
    vertOverflow: 'clip',
  },
];

function overflowModes(): string {
  const cell = grid(3, 2);
  return OVERFLOW_CASES.map((entry, index) =>
    shape({
      id: 10 + index,
      name: entry.name,
      ...cell(index),
      line: BOX_LINE,
      textBody: txBody({
        bodyPr: bodyPr({
          wrap: 'square',
          ...(entry.vertOverflow === undefined ? {} : { vertOverflow: entry.vertOverflow }),
          autofit: entry.autofit,
        }),
        paras:
          textLine(entry.name, {
            lang: 'en-GB',
            sz: 1100,
            b: true,
            fill: solidFill(scheme('accent1')),
          }) +
          textLine(entry.note, {
            lang: 'en-GB',
            sz: 900,
            i: true,
            fill: solidFill(scheme('accent3')),
          }) +
          textLine(SPECIMEN + ' ' + SPECIMEN, {
            lang: 'en-GB',
            sz: 1600,
            fill: solidFill(scheme('tx1')),
          }),
      }),
    }),
  ).join('');
}

// ------------------------- slide 3: spcFirstLastPara, and empty-paragraph height

const SPACED_PARAGRAPHS =
  para({
    props: { spcBef: spcPts(3600) },
    content: run('First paragraph, 36pt spcBef', {
      lang: 'en-GB',
      sz: 1400,
      fill: solidFill(scheme('tx1')),
    }),
  }) +
  textLine('Middle paragraph, no spacing of its own', {
    lang: 'en-GB',
    sz: 1400,
    fill: solidFill(scheme('tx1')),
  }) +
  para({
    props: { spcAft: spcPts(3600) },
    content: run('Last paragraph, 36pt spcAft', {
      lang: 'en-GB',
      sz: 1400,
      fill: solidFill(scheme('tx1')),
    }),
  });

/** Four empty paragraphs whose only height source is `a:endParaRPr`. */
const EMPTY_PARAGRAPH_STACK =
  textLine('Above four empty paragraphs at 8, 16, 32 and 54pt:', {
    lang: 'en-GB',
    sz: 1200,
    fill: solidFill(scheme('accent3')),
  }) +
  [800, 1600, 3200, 5400].map((sz) => para({ endProps: { lang: 'en-GB', sz } })).join('') +
  textLine('Below them. If the gap looks like one line, endParaRPr was ignored.', {
    lang: 'en-GB',
    sz: 1200,
    fill: solidFill(scheme('accent3')),
  }) +
  // An empty paragraph carrying formatting nothing will ever draw: the size is
  // real and the bold, the colour and the typeface are inert until somebody
  // types into it. Dropping them on export loses the user's caret formatting.
  para({
    endProps: {
      lang: 'en-GB',
      sz: 2400,
      b: true,
      latin: 'Georgia',
      fill: solidFill(srgb('C00000')),
    },
  });

function spacingAndEmptyParagraphs(): string {
  const cell = grid(2, 2);
  const boxes = [
    {
      name: 'spcFirstLastPara unset — first spcBef and last spcAft discarded',
      bodyPr: bodyPr({ wrap: 'square', anchor: 'ctr' }),
      paras: SPACED_PARAGRAPHS,
    },
    {
      name: 'spcFirstLastPara 1 — both honoured',
      bodyPr: bodyPr({ wrap: 'square', anchor: 'ctr', spcFirstLastPara: true }),
      paras: SPACED_PARAGRAPHS,
    },
    {
      name: 'a:endParaRPr sets the height of an empty paragraph',
      bodyPr: bodyPr({ wrap: 'square', anchor: 't' }),
      paras: EMPTY_PARAGRAPH_STACK,
    },
    {
      name: 'anchor b with the same paragraphs — the shift is the same size',
      bodyPr: bodyPr({ wrap: 'square', anchor: 'b' }),
      paras: SPACED_PARAGRAPHS,
    },
  ];
  return boxes
    .map((box, index) =>
      shape({
        id: 10 + index,
        name: box.name,
        ...cell(index),
        line: BOX_LINE,
        textBody: txBody({ bodyPr: box.bodyPr, paras: box.paras }),
      }),
    )
    .join('');
}

export const a11Autofit: ProbeDeck = {
  id: 'a11-autofit',
  title: 'PPTX Studio corpus: a11 autofit',
  description:
    'Twelve boxes carrying the same sentence at twelve stored normAutofit scales from 100% to ' +
    '25%, so a viewer that recomputes rather than obeying is wrong in eleven of them; noAutofit, ' +
    'spAutoFit and the three vertOverflow modes; spcFirstLastPara set and unset with 36pt of ' +
    'spacing so the default-false behaviour is a centimetre apart; and four empty paragraphs ' +
    'whose only height source is a:endParaRPr.',
  features: {
    shape: 28,
    placeholder: 6,
    presetGeom: 22,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a11 autofit',
    slides: [
      { title: 'a11 — twelve stored autofit scales', body: storedScales() },
      { title: 'a11 — the three autofit modes, and overflow', body: overflowModes() },
      {
        title: 'a11 — paragraph spacing and empty-paragraph height',
        body: spacingAndEmptyParagraphs(),
      },
    ],
  }),
};

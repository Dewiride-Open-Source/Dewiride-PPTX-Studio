import { placeholderXml, TITLE_BOX, type ProbeLayout } from '../package.ts';
import { grid, scheme, shape, solidFill } from '../shapes.ts';
import { lstStyle, para, pPr, rPr, run, spcPct, spcPts, textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The ten-source text cascade, one visible property per source.
 *
 * A resolved character property in DrawingML can come from ten different
 * places, merged **per property and per level** rather than per element:
 *
 * ```
 *  1 run a:rPr                     6 master p:txStyles[bucket][lvl]
 *  2 paragraph a:pPr/a:defRPr      7 p:defaultTextStyle  (non-placeholders only)
 *  3 shape a:lstStyle[lvl]         8 theme a:objectDefaults/a:txDef
 *  4 layout placeholder            9 theme a:fontScheme
 *  5 master placeholder           10 schema defaults
 * ```
 *
 * Slide 1 sets nine paragraphs at nine levels so that each level's size and
 * colour can only have come from one of sources 1 to 6, and slide 2 does the
 * same for 7 to 10 with plain text boxes, which are the only shapes that see
 * `p:defaultTextStyle` at all. A resolver test can then assert `origin` per
 * paragraph and fail loudly when a tier is skipped, rather than "looking about
 * right" because two tiers happened to agree.
 *
 * The bucket a placeholder reads from is not its type verbatim:
 * `title`/`ctrTitle` read `titleStyle`; `dt`/`ftr`/`sldNum` and anything with
 * no `p:ph` at all read `otherStyle`; everything else reads `bodyStyle`.
 *
 * ## Slide 3: `ST_Percentage` has two lexical forms and both are legal
 *
 * `<a:spcPct val="150000"/>` and `<a:spcPct val="150%"/>` mean the same thing.
 * `parseInt("150%")` yields 150, which is 0.15% line spacing, and the slide
 * collapses into a single overlapping line - a failure that looks like a layout
 * bug rather than a parsing one. The string form lives in exactly **one** shape
 * on slide 3, so if PowerPoint turns out to refuse it the bisection is two
 * steps rather than a hunt.
 *
 * The same slide carries the `@marL`/`@indent` trap. Their schema defaults are
 * 347663 and -342900, and they apply only when nothing is inheritable.
 * Materializing them before the cascade runs indents every paragraph in the
 * deck by a third of an inch, so one box here states them and the next states
 * nothing, and the two must not render alike.
 */

// ------------------------------------------------------------ the master tier

const BODY_SIZES = [2000, 1800, 1700, 1600, 1500, 1400, 1300, 1200, 1100] as const;

/** `p:txStyles`: sources 6, and the bucket a shape reads is decided by its `p:ph`. */
const TEXT_STYLES =
  '<p:txStyles>' +
  '<p:titleStyle>' +
  pPr('lvl1pPr', {
    algn: 'l',
    defRPr: rPr('defRPr', { sz: 2400, fill: solidFill(scheme('tx1')), latin: '+mj-lt' }),
  }) +
  '</p:titleStyle>' +
  '<p:bodyStyle>' +
  BODY_SIZES.map((sz, index) =>
    pPr(`lvl${String(index + 1)}pPr`, {
      marL: index === 0 ? 0 : 342900 * index,
      indent: index === 0 ? 0 : -342900,
      defRPr: rPr('defRPr', { sz, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
    }),
  ).join('') +
  '</p:bodyStyle>' +
  '<p:otherStyle>' +
  [1, 2, 3, 4, 5, 6, 7, 8, 9]
    .map((n) =>
      pPr(`lvl${String(n)}pPr`, {
        defRPr: rPr('defRPr', { sz: 1800, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
      }),
    )
    .join('') +
  '</p:otherStyle>' +
  '</p:txStyles>';

/** Source 5: the master's own body placeholder, which overrides level 7 alone. */
const MASTER_PLACEHOLDERS =
  placeholderXml({ id: 2, name: 'Title Placeholder 1', type: 'title', ...TITLE_BOX }) +
  placeholderXml({
    id: 3,
    name: 'Text Placeholder 2',
    type: 'body',
    idx: 1,
    x: 457200,
    y: 1097280,
    cx: 11277600,
    cy: 5303520,
    lstStyle: lstStyle({
      7: { defRPr: rPr('defRPr', { sz: 1900, fill: solidFill(scheme('accent3')) }) },
    }),
  });

/** Source 8: the theme's default text shape, which nothing else can supply. */
const THEME_OBJECT_DEFAULTS =
  '<a:objectDefaults><a:txDef><a:spPr/><a:bodyPr/>' +
  lstStyle({
    3: { defRPr: rPr('defRPr', { sz: 2000, fill: solidFill(scheme('accent6')) }) },
  }) +
  '</a:txDef></a:objectDefaults>';

/** Source 7: read by non-placeholders only, and it stops at level 2 on purpose. */
const DEFAULT_TEXT_STYLE =
  '<p:defaultTextStyle>' +
  pPr('defPPr', { defRPr: rPr('defRPr', { latin: '+mn-lt' }) }) +
  pPr('lvl1pPr', {
    defRPr: rPr('defRPr', { sz: 2200, fill: solidFill(scheme('accent5')) }),
  }) +
  pPr('lvl2pPr', {
    marL: 457200,
    defRPr: rPr('defRPr', { sz: 1600, fill: solidFill(scheme('accent5')) }),
  }) +
  '</p:defaultTextStyle>';

// ------------------------------------------------------------------- layouts

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'cust',
    name: 'Cascade',
    hasTitle: true,
    shapes:
      placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }) +
      // Source 4: the layout placeholder, which overrides level 2 alone.
      placeholderXml({
        id: 3,
        name: 'Content Placeholder 2',
        type: 'body',
        idx: 1,
        x: 457200,
        y: 1188720,
        cx: 11277600,
        cy: 5181600,
        lstStyle: lstStyle({
          2: { defRPr: rPr('defRPr', { sz: 2100, fill: solidFill(scheme('accent1')) }) },
        }),
      }),
  },
  {
    type: 'cust',
    name: 'Plain Boxes',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
];

// -------------------------------------------------------------------- slide 1

/** What each of the nine paragraphs is the probe for, and how it says so. */
const CASCADE_ROWS: readonly {
  readonly label: string;
  readonly paragraph: string;
}[] = [
  {
    label: '6 · master p:bodyStyle/lvl1pPr — 20pt tx1',
    paragraph: '',
  },
  {
    label: '4 · layout placeholder a:lstStyle/lvl2pPr — 21pt accent1',
    paragraph: '',
  },
  {
    label: '3 · shape a:lstStyle/lvl3pPr — 20pt accent2',
    paragraph: '',
  },
  {
    label: '2 · paragraph a:pPr/a:defRPr — 18pt accent4',
    paragraph: 'defRPr',
  },
  {
    label: '1 · run a:rPr — 17pt accent5',
    paragraph: 'rPr',
  },
  {
    label: '6 · master p:bodyStyle/lvl6pPr — 14pt tx1',
    paragraph: '',
  },
  {
    label: '5 · master placeholder a:lstStyle/lvl7pPr — 19pt accent3',
    paragraph: '',
  },
  {
    label: '6 · master p:bodyStyle/lvl8pPr — 12pt tx1',
    paragraph: '',
  },
  {
    label: '9 · theme a:fontScheme via +mn-lt, size from lvl9pPr — 11pt',
    paragraph: '',
  },
];

function cascadeParagraphs(): string {
  return CASCADE_ROWS.map((row, index) => {
    const level = index === 0 ? {} : { lvl: index };
    if (row.paragraph === 'defRPr') {
      return para({
        props: {
          ...level,
          defRPr: rPr('defRPr', { sz: 1800, fill: solidFill(scheme('accent4')) }),
        },
        content: run(row.label, { lang: 'en-GB', dirty: false }),
      });
    }
    if (row.paragraph === 'rPr') {
      return para({
        props: level,
        content: run(row.label, {
          lang: 'en-GB',
          sz: 1700,
          fill: solidFill(scheme('accent5')),
          dirty: false,
        }),
      });
    }
    return para({ props: level, content: run(row.label, { lang: 'en-GB', dirty: false }) });
  }).join('');
}

// -------------------------------------------------------------------- slide 3

const PERCENT_BOXES: readonly { readonly name: string; readonly paras: string }[] = [
  {
    name: 'lnSpc as an integer',
    paras:
      textLine('lnSpc val="150000" — line 1 of three') +
      para({
        props: { lnSpc: spcPct(150000) },
        content: run('line 2, spaced 150% the integer way', { lang: 'en-GB' }),
      }) +
      para({
        props: { lnSpc: spcPct(150000) },
        content: run('line 3', { lang: 'en-GB' }),
      }),
  },
  {
    // The only shape in the corpus written with the percent-string form. If
    // PowerPoint ever refuses it, the bisection is slide 3, then this shape.
    name: 'lnSpc as a percent string',
    paras:
      textLine('lnSpc val="150%" — line 1 of three') +
      para({
        props: { lnSpc: spcPct('150%') },
        content: run('line 2, spaced 150% the string way', { lang: 'en-GB' }),
      }) +
      para({
        props: { lnSpc: spcPct('150%'), spcBef: spcPct('50%') },
        content: run('line 3, with spcBef val="50%"', { lang: 'en-GB' }),
      }),
  },
  {
    name: 'spcBef and spcAft in points',
    paras:
      textLine('spcPts is hundredths of a point, not EMU') +
      para({
        props: { spcBef: spcPts(1200), spcAft: spcPts(600) },
        content: run('12pt before, 6pt after', { lang: 'en-GB' }),
      }) +
      textLine('and the paragraph after it'),
  },
  {
    name: 'marL and indent stated',
    paras:
      textLine('marL 347663, indent -342900 — stated') +
      para({
        props: { marL: 347663, indent: -342900 },
        content: run('a hanging indent, written down', { lang: 'en-GB' }),
      }),
  },
  {
    name: 'marL and indent absent',
    paras:
      textLine('no marL, no indent — must not look like the box to its left') +
      para({ content: run('flush left, because nothing set a margin', { lang: 'en-GB' }) }),
  },
  {
    name: 'spcFirstLastPara defaults false',
    paras:
      para({
        props: { spcBef: spcPts(2400) },
        content: run('24pt spcBef on the first paragraph — discarded', { lang: 'en-GB' }),
      }) +
      textLine('so this box starts flush with its top inset') +
      para({
        props: { spcAft: spcPts(2400) },
        content: run('24pt spcAft on the last — discarded too', { lang: 'en-GB' }),
      }),
  },
];

export const a07TextCascade: ProbeDeck = {
  id: 'a07-text-cascade',
  title: 'PPTX Studio corpus: a07 text cascade',
  description:
    'One visible property per source of the ten-source text cascade: nine paragraphs at nine ' +
    'levels whose size and colour can each only have come from one tier, plain text boxes for ' +
    'p:defaultTextStyle and the theme a:txDef that placeholders never see, and both lexical ' +
    'forms of ST_Percentage alongside the marL/indent eager-default trap.',
  features: {
    shape: 18,
    placeholder: 9,
    presetGeom: 9,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a07 text cascade',
    layouts: LAYOUTS,
    masterPlaceholders: MASTER_PLACEHOLDERS,
    textStyles: TEXT_STYLES,
    themeObjectDefaults: THEME_OBJECT_DEFAULTS,
    presentationTail: DEFAULT_TEXT_STYLE,
    slides: [
      {
        title: 'a07 — six sources, nine levels',
        layout: 0,
        body: placeholderXml({
          id: 10,
          name: 'Content Placeholder 2',
          type: 'body',
          idx: 1,
          // Source 3: the shape's own list style, overriding level 3 alone.
          lstStyle: lstStyle({
            3: { defRPr: rPr('defRPr', { sz: 2000, fill: solidFill(scheme('accent2')) }) },
          }),
          body: cascadeParagraphs(),
        }),
      },
      {
        title: 'a07 — the sources placeholders never see',
        layout: 1,
        body: (() => {
          const cell = grid(3, 1);
          const boxes = [
            {
              name: '7 · p:defaultTextStyle lvl1pPr',
              paras:
                textLine('A plain text box at level 1.') +
                textLine(
                  '22pt accent5 comes from p:defaultTextStyle, which a placeholder ignores.',
                ),
            },
            {
              name: '8 · theme a:txDef lvl3pPr',
              paras:
                textLine('A plain text box at level 3.') +
                para({
                  props: { lvl: 2 },
                  content: run(
                    '20pt accent6 comes from the theme objectDefaults, because ' +
                      'p:defaultTextStyle stops at level 2.',
                    { lang: 'en-GB' },
                  ),
                }),
            },
            {
              name: '9 and 10 · fontScheme, then the schema',
              paras:
                textLine('Nothing here states a typeface, a weight or an alignment.') +
                textLine('The face is +mn-lt from the theme; b=0 and algn=l are the schema.'),
            },
          ];
          return boxes
            .map((box, index) =>
              shape({
                id: 10 + index,
                name: box.name,
                ...cell(index),
                line: '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>',
                textBody: txBody({
                  bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
                  paras: box.paras,
                }),
              }),
            )
            .join('');
        })(),
      },
      {
        title: 'a07 — ST_Percentage, both forms',
        layout: 1,
        body: (() => {
          const cell = grid(3, 2);
          return PERCENT_BOXES.map((box, index) =>
            shape({
              id: 10 + index,
              name: box.name,
              ...cell(index),
              line: '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>',
              textBody: txBody({
                bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
                paras: box.paras,
              }),
            }),
          ).join('');
        })(),
      },
    ],
  }),
};

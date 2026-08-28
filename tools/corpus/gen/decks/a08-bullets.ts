import { png } from '../png.ts';
import { grid, scheme, shape, solidFill, srgb } from '../shapes.ts';
import {
  buAutoNum,
  buBlip,
  buChar,
  buFont,
  buNone,
  buSzPct,
  lstStyle,
  para,
  run,
  textLine,
  txBody,
  type ParaProps,
} from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Bullets: every scheme, every source of the glyph, and the PUA trap.
 *
 * ## Symbol bullets are Private Use Area code points
 *
 * `<a:buChar char="§"/>` with `<a:buFont typeface="Wingdings"/>` is **not** a
 * section sign. Symbol-encoded TrueType faces map their glyphs into U+F000 to
 * U+F0FF, and Office writes the low code point while meaning the PUA one; a
 * renderer that takes the character literally and falls back to a text font
 * shows a section sign, a spade, or a box, on every bulleted list in the deck.
 *
 * Slide 1 writes the pair both ways round - the low character and its PUA
 * equivalent, same font, side by side - so the two must render identically.
 *
 * ## The forty-one autonumber schemes
 *
 * `ST_TextAutonumberScheme` has exactly 41 values, and slide 2 has one
 * paragraph per value, labelled with the value. They are not decoration: eleven
 * of them are East Asian, Thai, Hindi or Hebrew numbering that no
 * `Intl.NumberFormat` call produces, and having them written down is what turns
 * "implement autonumbering" into a finite job with a visible score.
 *
 * `@startAt` is exercised on slide 3 rather than here, because a scheme's
 * appearance and its counter are separate things to get wrong.
 *
 * ## `buBlip`
 *
 * An image bullet resolves its `r:embed` against the **slide's** relationships,
 * like any other blip. Slide 3 carries one, and the census is unmoved by it:
 * `a:buBlip` holds an `a:blip`, not an `a:blipFill`, so nothing here counts as
 * a picture fill. That distinction is the same one that made `blipFill` read
 * zero across every benchmark deck full of pictures.
 */

const BULLET_RID = 'rId2';

/** A four-pixel diamond, small enough that a bullet-sized blip is honest about it. */
function bulletPng(): Uint8Array {
  return png(8, 8, (x, y) => {
    const dx = Math.abs(x - 3.5);
    const dy = Math.abs(y - 3.5);
    return dx + dy <= 3.5 ? 0x4472c4 : 0xffffff;
  });
}

const BOX_LINE = '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>';
const BOX_BODY_PR = '<a:bodyPr wrap="square" lIns="91440" tIns="45720"><a:normAutofit/></a:bodyPr>';

// ------------------------------------------------- slide 1: where a glyph comes from

const WINGDINGS = buFont('Wingdings', 2, 2);
const ARIAL = buFont('Arial', 34, 0);

const CHAR_ROWS: readonly { readonly label: string; readonly props: ParaProps }[] = [
  {
    label: 'buChar "•" with buFont Arial — an ordinary text bullet',
    props: { buFont: ARIAL, bullet: buChar('•') },
  },
  {
    label: 'buChar "§" with buFont Wingdings — the low code point Office writes',
    props: { buFont: WINGDINGS, bullet: buChar('§') },
  },
  {
    label: 'buChar U+F0A7 with buFont Wingdings — the PUA code point it means',
    props: { buFont: WINGDINGS, bullet: buChar('\uF0A7') },
  },
  {
    label: 'buChar "ü" with buFont Wingdings — a tick, not a u-umlaut',
    props: { buFont: WINGDINGS, bullet: buChar('ü') },
  },
  {
    label: 'buClr accent2, buSzPct 150% — the bullet coloured and sized apart from its run',
    props: {
      buClr: '<a:buClr>' + scheme('accent2') + '</a:buClr>',
      buSz: buSzPct(150000),
      buFont: ARIAL,
      bullet: buChar('▪'),
    },
  },
  {
    label: 'buClrTx, buSzTx, buFontTx — the bullet follows the run in all three',
    props: {
      buClr: '<a:buClrTx/>',
      buSz: '<a:buSzTx/>',
      buFont: '<a:buFontTx/>',
      bullet: buChar('–'),
    },
  },
  {
    label: 'buNone — no bullet at all, and no reserved indent either',
    props: { marL: 0, indent: 0, bullet: buNone },
  },
];

function glyphSources(): string {
  const cell = grid(1, 1);
  const paras = CHAR_ROWS.map((row) =>
    para({
      props: { marL: 457200, indent: -457200, ...row.props },
      content: run(row.label, { lang: 'en-GB', sz: 1400, fill: solidFill(scheme('tx1')) }),
    }),
  ).join('');
  return shape({
    id: 10,
    name: 'Where the glyph comes from',
    ...cell(0),
    line: BOX_LINE,
    textBody: txBody({ bodyPr: BOX_BODY_PR, paras }),
  });
}

// ------------------------------------------- slide 2: all 41 autonumber schemes

/**
 * Every value of `ST_TextAutonumberScheme`.
 *
 * Taken from the enumeration, in its own declared order, not grouped by script
 * - so that a value added or dropped shows up as a diff in one place.
 */
const AUTONUMBER_SCHEMES: readonly string[] = [
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaLcPeriod',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'alphaUcPeriod',
  'arabic1Minus',
  'arabic2Minus',
  'arabicDbPeriod',
  'arabicDbPlain',
  'arabicParenBoth',
  'arabicParenR',
  'arabicPeriod',
  'arabicPlain',
  'circleNumDbPlain',
  'circleNumWdBlackPlain',
  'circleNumWdWhitePlain',
  'ea1ChsPeriod',
  'ea1ChsPlain',
  'ea1ChtPeriod',
  'ea1ChtPlain',
  'ea1JpnChsDbPeriod',
  'ea1JpnKorPeriod',
  'ea1JpnKorPlain',
  'hebrew2Minus',
  'hindiAlpha1Period',
  'hindiAlphaPeriod',
  'hindiNumParenR',
  'hindiNumPeriod',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcParenBoth',
  'romanUcParenR',
  'romanUcPeriod',
  'thaiAlphaParenBoth',
  'thaiAlphaParenR',
  'thaiAlphaPeriod',
  'thaiNumParenBoth',
  'thaiNumParenR',
  'thaiNumPeriod',
];

function autonumberSchemes(): string {
  const cell = grid(3, 1);
  const perColumn = 14;
  const boxes: string[] = [];
  for (let column = 0; column < 3; column++) {
    const slice = AUTONUMBER_SCHEMES.slice(column * perColumn, (column + 1) * perColumn);
    const paras = slice
      .map((name) =>
        para({
          props: {
            marL: 742950,
            indent: -742950,
            buFont: buFont('+mj-lt'),
            bullet: buAutoNum(name),
          },
          content: run(name, { lang: 'en-GB', sz: 1100, fill: solidFill(scheme('tx1')) }),
        }),
      )
      .join('');
    const first = column * perColumn + 1;
    boxes.push(
      shape({
        id: 10 + column,
        name: 'Schemes ' + String(first) + ' to ' + String(first + slice.length - 1),
        ...cell(column),
        line: BOX_LINE,
        textBody: txBody({ bodyPr: BOX_BODY_PR, paras }),
      }),
    );
  }
  return boxes.join('');
}

// ------------------------------- slide 3: blips, startAt, and nine inherited levels

/** Nine levels of bullet, declared once on the shape and inherited by paragraph. */
const NINE_LEVELS = lstStyle({
  1: { marL: 285750, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('•') },
  2: { marL: 571500, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('–') },
  3: { marL: 857250, indent: -285750, buFont: buFont('Wingdings', 2, 2), bullet: buChar('§') },
  4: { marL: 1143000, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('–') },
  5: { marL: 1428750, indent: -285750, buFont: buFont('Wingdings', 2, 2), bullet: buChar('Ø') },
  6: { marL: 1714500, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('»') },
  7: { marL: 2000250, indent: -285750, buFont: buFont('Wingdings', 2, 2), bullet: buChar('Ø') },
  8: { marL: 2286000, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('•') },
  9: { marL: 2571750, indent: -285750, buFont: buFont('Arial', 34, 0), bullet: buChar('•') },
});

function blipsAndCounters(): string {
  const cell = grid(2, 1);

  const blipParas =
    textLine('buBlip resolves r:embed against the slide, like any other blip:', {
      lang: 'en-GB',
      sz: 1200,
    }) +
    [1, 2, 3]
      .map((n) =>
        para({
          props: {
            marL: 457200,
            indent: -457200,
            buSz: buSzPct(100000),
            bullet: buBlip(BULLET_RID),
          },
          content: run('image bullet ' + String(n), {
            lang: 'en-GB',
            sz: 1400,
            fill: solidFill(scheme('tx1')),
          }),
        }),
      )
      .join('') +
    textLine('startAt="7" restarts the counter mid-list:', { lang: 'en-GB', sz: 1200 }) +
    [7, 8, 9]
      .map((n, index) =>
        para({
          props: {
            marL: 457200,
            indent: -457200,
            buFont: buFont('+mj-lt'),
            bullet: index === 0 ? buAutoNum('arabicPeriod', 7) : buAutoNum('arabicPeriod'),
          },
          content: run('numbered ' + String(n), {
            lang: 'en-GB',
            sz: 1400,
            fill: solidFill(scheme('tx1')),
          }),
        }),
      )
      .join('');

  const levelParas = [0, 1, 2, 3, 4, 5, 6, 7, 8]
    .map((lvl) =>
      para({
        ...(lvl === 0 ? {} : { props: { lvl } }),
        content: run('level ' + String(lvl + 1) + ', bullet from a:lstStyle', {
          lang: 'en-GB',
          sz: 1200,
          fill: solidFill(scheme('tx1')),
        }),
      }),
    )
    .join('');

  return (
    shape({
      id: 10,
      name: 'buBlip and startAt',
      ...cell(0),
      line: BOX_LINE,
      textBody: txBody({ bodyPr: BOX_BODY_PR, paras: blipParas }),
    }) +
    shape({
      id: 11,
      name: 'Nine levels from one a:lstStyle',
      ...cell(1),
      fill: solidFill(srgb('F2F2F2')),
      line: BOX_LINE,
      textBody: txBody({ bodyPr: BOX_BODY_PR, lstStyle: NINE_LEVELS, paras: levelParas }),
    })
  );
}

export const a08Bullets: ProbeDeck = {
  id: 'a08-bullets',
  title: 'PPTX Studio corpus: a08 bullets',
  description:
    'All 41 ST_TextAutonumberScheme values one paragraph each, buChar with buFont/buClr/buSzPct ' +
    'and their three follow-the-text forms, the Wingdings PUA pair written both ways so they ' +
    'must render alike, an image bullet resolving against the slide rels, startAt mid-list, and ' +
    'nine bullet levels declared once in a shape a:lstStyle.',
  features: {
    shape: 12,
    placeholder: 6,
    presetGeom: 6,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a08 bullets',
    slides: [
      { title: 'a08 — where a bullet glyph comes from', body: glyphSources() },
      { title: 'a08 — all 41 autonumber schemes', body: autonumberSchemes() },
      {
        title: 'a08 — image bullets, counters, nine levels',
        body: blipsAndCounters(),
        rels: [
          {
            id: BULLET_RID,
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            target: '../media/image1.png',
          },
        ],
      },
    ],
    parts: [
      {
        name: 'ppt/media/image1.png',
        bytes: bulletPng(),
        contentType: { kind: 'default', extension: 'png', type: 'image/png' },
      },
    ],
  }),
};

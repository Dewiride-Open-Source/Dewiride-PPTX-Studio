import { graphicFrame, GRAPHIC_URI, scheme, solidFill, srgb } from '../shapes.ts';
import { aTxBody, textLine } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Tables: the occupancy grid, `a:tcPr`, and both sources of a table style.
 *
 * A table is not a shape. It is an `a:tbl` inside a `p:graphicFrame`, which
 * means three things a renderer has to get right before a single cell is
 * painted: the frame's transform is `p:xfrm` and not `a:xfrm`, the frame says
 * nothing about what it holds until `a:graphicData/@uri` is read, and the whole
 * table lives in the **slide part** rather than in a part of its own - unlike
 * every other thing a `p:graphicFrame` can carry.
 *
 * ## The occupancy grid, and why a merged-away cell is still there
 *
 * `a:tblGrid` fixes the column count once. Every `a:tr` then holds exactly that
 * many `a:tc` elements **whatever the merges say** - a cell swallowed by the one
 * to its left is written out as `<a:tc hMerge="1">` with an empty body, and it
 * still occupies its column. Drop it, or skip it while walking, and every cell
 * after it in that row shifts one column left.
 *
 * Four attributes, and they come in two pairs that are easy to confuse:
 *
 * | on the anchor cell | on the swallowed cell | direction  |
 * | ------------------ | --------------------- | ---------- |
 * | `gridSpan="n"`     | `hMerge="1"`          | horizontal |
 * | `rowSpan="n"`      | `vMerge="1"`          | vertical   |
 *
 * A 2x2 block merge carries `gridSpan` **and** `rowSpan` on its anchor, and the
 * three cells it covers carry `hMerge`, `vMerge`, and both at once. Slide 1
 * writes exactly that, so an implementation that handles only one axis produces
 * a visibly wrong grid rather than a subtly wrong one.
 *
 * ## `a:tcPr`, whose children are in an order nobody would guess
 *
 * `lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr, cell3D, <fill>, headers, extLst` -
 * left, right, top, bottom, and only then the two diagonals. Not clockwise, not
 * the CSS order, and the six are all `CT_LineProperties`, so a cell border is
 * an `a:ln` with a different tag name and everything sub-phase 2.8 does to
 * lines applies unchanged.
 *
 * ## Two places a table style can come from, and one of them is usually empty
 *
 * `a:tblPr` holds a choice: `a:tableStyleId`, a GUID resolved against
 * `ppt/tableStyles.xml`, or an inline `a:tableStyle` carrying the definition
 * itself. Slide 3 writes one of each, side by side.
 *
 * The GUID is the case that matters and the case that fails silently.
 * `ppt/tableStyles.xml` in every stock PowerPoint template - including this
 * corpus's - is an **empty** `<a:tblStyleLst def="{5C22544A-...}"/>`: a default
 * attribute and no definitions at all. So a renderer that resolves the id
 * against that part, finds nothing and paints nothing renders every real table
 * white and borderless while behaving exactly as the file told it to. Getting
 * that right is sub-phase 4.2's job, and it needs the 74 built-in definitions
 * re-derived from PowerPoint rather than read out of the package.
 *
 * The inline form is where the 13-layer cascade can actually be probed, and it
 * carries a trap of its own: **the schema sequence is not the application
 * order.** `CT_TableStyle` is
 *
 * ```
 * tblBg, wholeTbl, band1H, band2H, band1V, band2V, lastCol, firstCol,
 * lastRow, seCell, swCell, firstRow, neCell, nwCell, extLst
 * ```
 *
 * while sub-phase 4.3 applies wholeTbl, firstRow, lastRow, firstCol, lastCol,
 * the horizontal bands, the four corners, then the vertical bands, then the
 * cell's own `a:tcPr`. A parser that reads the children positionally and
 * applies them in the order it read them puts `lastCol` ahead of `firstRow` and
 * loses every bold header on a table that also styles its last column.
 *
 * `a:tcTxStyle/@b` and `@i` are the other trap, and they are not booleans:
 * `ST_OnOffStyleType` is `on`, `off` or `def`. A `getBool()`-shaped parser
 * reads `"on"` as false and drops the bold from every header row in the corpus.
 * The inline style writes all three values so that failure has somewhere to
 * show up.
 */

// ------------------------------------------------------------------ helpers

const CELL_TEXT = { sz: 1100, lang: 'en-GB' } as const;

/** One cell. `a:tc` is `txBody, tcPr, extLst`, and the body is never omitted. */
function tc(
  text: string,
  options: {
    readonly gridSpan?: number;
    readonly rowSpan?: number;
    readonly hMerge?: boolean;
    readonly vMerge?: boolean;
    readonly tcPr?: string;
    readonly bold?: boolean;
  } = {},
): string {
  const attributes =
    (options.rowSpan === undefined ? '' : ` rowSpan="${String(options.rowSpan)}"`) +
    (options.gridSpan === undefined ? '' : ` gridSpan="${String(options.gridSpan)}"`) +
    (options.hMerge === true ? ' hMerge="1"' : '') +
    (options.vMerge === true ? ' vMerge="1"' : '');
  const paras =
    text === ''
      ? '<a:p><a:endParaRPr lang="en-GB" sz="1100"/></a:p>'
      : textLine(text, { ...CELL_TEXT, ...(options.bold === true ? { b: true } : {}) });
  return (
    `<a:tc${attributes}>` +
    aTxBody({ bodyPr: '<a:bodyPr/>', paras }) +
    (options.tcPr ?? '<a:tcPr/>') +
    '</a:tc>'
  );
}

/** One row. `a:tr/@h` is required, and is a minimum rather than a fixed height. */
const tr = (height: number, cells: readonly string[]): string =>
  `<a:tr h="${String(height)}">${cells.join('')}</a:tr>`;

function tbl(spec: {
  readonly tblPr: string;
  readonly columns: readonly number[];
  readonly rows: string;
}): string {
  return (
    '<a:tbl>' +
    spec.tblPr +
    '<a:tblGrid>' +
    spec.columns.map((w) => `<a:gridCol w="${String(w)}"/>`).join('') +
    '</a:tblGrid>' +
    spec.rows +
    '</a:tbl>'
  );
}

/** The GUID `ppt/tableStyles.xml` names as its `def` - Medium Style 2 Accent 1. */
const MEDIUM_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

const frame = (
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  table: string,
): string =>
  graphicFrame({
    id,
    name,
    x,
    y,
    cx,
    cy,
    locks: 'noGrp="1"',
    uri: GRAPHIC_URI.table,
    content: table,
  });

// ------------------------------------------------------------------ slide 1

const COL = 2540000;
const ROW = 700000;

/**
 * Four columns, four rows, three merges and one of them on both axes.
 *
 * Row 1: a `gridSpan="2"` header over columns 1-2, then two plain cells.
 * Rows 2-3, columns 3-4: a 2x2 block, so its anchor spans and stacks at once.
 * Rows 2-3, column 1: a `rowSpan="2"`.
 */
const MERGES = tbl({
  tblPr:
    '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>' +
    MEDIUM_2_ACCENT_1 +
    '</a:tableStyleId></a:tblPr>',
  columns: [COL, COL, COL, COL],
  rows:
    tr(ROW, [
      tc('gridSpan="2" - one cell, two columns', { gridSpan: 2, bold: true }),
      tc('', { hMerge: true }),
      tc('plain', { bold: true }),
      tc('plain', { bold: true }),
    ]) +
    tr(ROW, [
      tc('rowSpan="2" starts here', { rowSpan: 2 }),
      tc('plain'),
      // The block: two columns wide and two rows tall, from one anchor.
      tc('gridSpan="2" rowSpan="2" - a 2x2 block', { gridSpan: 2, rowSpan: 2 }),
      tc('', { hMerge: true }),
    ]) +
    tr(ROW, [
      // Swallowed vertically by the rowSpan above it.
      tc('', { vMerge: true }),
      tc('plain'),
      // Swallowed vertically by the block; still one `a:tc` per column.
      tc('', { vMerge: true }),
      tc('', { hMerge: true, vMerge: true }),
    ]) +
    tr(ROW, [tc('plain'), tc('plain'), tc('plain'), tc('plain')]),
});

// ------------------------------------------------------------------ slide 2

/** An `a:ln`-shaped border under one of the six `a:tcPr` border tags. */
const border = (tag: string, width: number, colour: string, dash?: string): string =>
  `<a:${tag} w="${String(width)}" cap="flat" cmpd="sng" algn="ctr">` +
  `<a:solidFill>${colour}</a:solidFill>` +
  (dash === undefined ? '' : `<a:prstDash val="${dash}"/>`) +
  `</a:${tag}>`;

const BORDERS = tbl({
  tblPr: '<a:tblPr/>',
  columns: [COL, COL, COL, COL],
  rows:
    tr(ROW, [
      tc('lnL + lnR', {
        tcPr:
          '<a:tcPr>' +
          border('lnL', 57150, srgb('C00000')) +
          border('lnR', 57150, srgb('C00000')) +
          '</a:tcPr>',
      }),
      tc('lnT + lnB', {
        tcPr:
          '<a:tcPr>' +
          border('lnT', 57150, srgb('0070C0')) +
          border('lnB', 57150, srgb('0070C0')) +
          '</a:tcPr>',
      }),
      tc('both diagonals', {
        tcPr:
          '<a:tcPr>' +
          border('lnTlToBr', 28575, srgb('7030A0')) +
          border('lnBlToTr', 28575, srgb('7030A0'), 'dash') +
          '</a:tcPr>',
      }),
      tc('all six at once', {
        tcPr:
          '<a:tcPr>' +
          border('lnL', 19050, scheme('accent1')) +
          border('lnR', 19050, scheme('accent2')) +
          border('lnT', 19050, scheme('accent3')) +
          border('lnB', 19050, scheme('accent4')) +
          border('lnTlToBr', 12700, scheme('accent5'), 'sysDot') +
          border('lnBlToTr', 12700, scheme('accent6'), 'sysDot') +
          '</a:tcPr>',
      }),
    ]) +
    tr(ROW, [
      // The four margins, all four different, so a symmetric default shows.
      tc('marL 0 marR 457200', {
        tcPr: '<a:tcPr marL="0" marR="457200" marT="0" marB="0"/>',
      }),
      tc('anchor="b"', { tcPr: '<a:tcPr anchor="b"/>' }),
      tc('anchor="ctr" anchorCtr="1"', { tcPr: '<a:tcPr anchor="ctr" anchorCtr="1"/>' }),
      tc('horzOverflow="clip"', { tcPr: '<a:tcPr horzOverflow="clip"/>' }),
    ]) +
    tr(ROW, [
      tc('vert="vert270"', { tcPr: '<a:tcPr vert="vert270"/>' }),
      tc('vert="vert"', { tcPr: '<a:tcPr vert="vert"/>' }),
      tc('vert="eaVert"', { tcPr: '<a:tcPr vert="eaVert"/>' }),
      tc('vert="wordArtVert"', { tcPr: '<a:tcPr vert="wordArtVert"/>' }),
    ]) +
    tr(ROW, [
      tc('solidFill', {
        tcPr: '<a:tcPr>' + solidFill(scheme('accent1', '<a:alpha val="40000"/>')) + '</a:tcPr>',
      }),
      tc('gradFill', {
        tcPr:
          '<a:tcPr><a:gradFill rotWithShape="1">' +
          '<a:gsLst><a:gs pos="0">' +
          scheme('accent2') +
          '</a:gs><a:gs pos="100000">' +
          scheme('accent6') +
          '</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>' +
          '</a:gradFill></a:tcPr>',
      }),
      tc('pattFill', {
        tcPr:
          '<a:tcPr><a:pattFill prst="ltUpDiag"><a:fgClr>' +
          scheme('accent4') +
          '</a:fgClr><a:bgClr>' +
          scheme('bg1') +
          '</a:bgClr></a:pattFill></a:tcPr>',
      }),
      tc('cell3D', {
        // `a:cell3D` sits between the diagonals and the fill, and its
        // `@prstMaterial` is the enumeration a shape's `a:sp3d` uses.
        tcPr:
          '<a:tcPr><a:cell3D prstMaterial="matte">' +
          '<a:bevel w="76200" h="76200" prst="circle"/>' +
          '<a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig>' +
          '</a:cell3D>' +
          solidFill(scheme('accent5', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')) +
          '</a:tcPr>',
      }),
    ]),
});

// ------------------------------------------------------------------ slide 3

/** An `a:tcStyle` with a fill and a full border set, for the inline style. */
const tcStyle = (fill: string, borderColour: string): string =>
  '<a:tcStyle><a:tcBdr>' +
  ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']
    .map(
      (tag) =>
        `<a:${tag}><a:ln w="12700" cmpd="sng"><a:solidFill>${borderColour}</a:solidFill></a:ln></a:${tag}>`,
    )
    .join('') +
  '</a:tcBdr>' +
  // The wrapper matters and is easy to miss: `a:tblPr` takes a fill element
  // directly, while `a:tcStyle` and `a:tblBg` take a choice of `a:fill` or
  // `a:fillRef` - so the fill goes one level deeper here than it does two
  // elements away. Sub-phase 1.2's `V012` found this deck writing it flat.
  '<a:fill>' +
  fill +
  '</a:fill>' +
  '</a:tcStyle>';

/**
 * An inline `a:tableStyle`, written in **schema order** rather than in the
 * order sub-phase 4.3 applies its parts. All three `ST_OnOffStyleType` values
 * are here: `on` on the first row, `off` on the last, `def` on the whole table.
 */
const INLINE_STYLE =
  '<a:tableStyle styleId="{1B1D64F0-0000-4000-8000-0000000A200A}"' +
  ' styleName="PPTX Studio Corpus Probe">' +
  // tblBg first, and it is the table's own background, not a cell fill.
  '<a:tblBg><a:fill>' +
  solidFill(scheme('bg1')) +
  '</a:fill></a:tblBg>' +
  '<a:wholeTbl>' +
  '<a:tcTxStyle b="def" i="def">' +
  '<a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef>' +
  scheme('tx1') +
  '</a:tcTxStyle>' +
  tcStyle(solidFill(scheme('bg1')), scheme('tx1', '<a:alpha val="40000"/>')) +
  '</a:wholeTbl>' +
  '<a:band1H>' +
  tcStyle(solidFill(scheme('accent1', '<a:alpha val="20000"/>')), scheme('accent1')) +
  '</a:band1H>' +
  '<a:band2H>' +
  tcStyle(solidFill(scheme('bg1')), scheme('accent1')) +
  '</a:band2H>' +
  '<a:band1V>' +
  tcStyle(solidFill(scheme('accent3', '<a:alpha val="15000"/>')), scheme('accent3')) +
  '</a:band1V>' +
  // lastCol before firstCol, and lastRow before firstRow. Schema order.
  '<a:lastCol>' +
  `<a:tcTxStyle b="on" i="on">${scheme('tx1')}</a:tcTxStyle>` +
  tcStyle(solidFill(scheme('accent6', '<a:alpha val="30000"/>')), scheme('accent6')) +
  '</a:lastCol>' +
  '<a:firstCol>' +
  `<a:tcTxStyle b="on" i="def">${scheme('tx1')}</a:tcTxStyle>` +
  tcStyle(solidFill(scheme('accent2', '<a:alpha val="30000"/>')), scheme('accent2')) +
  '</a:firstCol>' +
  '<a:lastRow>' +
  `<a:tcTxStyle b="off" i="def">${scheme('tx1')}</a:tcTxStyle>` +
  tcStyle(solidFill(scheme('accent4', '<a:alpha val="25000"/>')), scheme('accent4')) +
  '</a:lastRow>' +
  '<a:seCell>' +
  tcStyle(solidFill(srgb('FFE699')), srgb('BF8F00')) +
  '</a:seCell>' +
  '<a:swCell>' +
  tcStyle(solidFill(srgb('C5E0B4')), srgb('538135')) +
  '</a:swCell>' +
  '<a:firstRow>' +
  `<a:tcTxStyle b="on" i="def">${scheme('bg1')}</a:tcTxStyle>` +
  tcStyle(solidFill(scheme('accent1')), scheme('accent1')) +
  '</a:firstRow>' +
  '<a:neCell>' +
  tcStyle(solidFill(srgb('F8CBAD')), srgb('C55A11')) +
  '</a:neCell>' +
  '<a:nwCell>' +
  tcStyle(solidFill(srgb('B4C7E7')), srgb('2E5395')) +
  '</a:nwCell>' +
  '</a:tableStyle>';

const styledRows = (label: string): string =>
  tr(ROW, [tc(label, { bold: true }), tc('col 2', { bold: true }), tc('col 3', { bold: true })]) +
  tr(ROW, [tc('first col'), tc('body'), tc('last col')]) +
  tr(ROW, [tc('first col'), tc('body'), tc('last col')]) +
  tr(ROW, [tc('last row'), tc('last row'), tc('last row')]);

const BY_ID = tbl({
  tblPr:
    '<a:tblPr firstRow="1" lastRow="1" firstCol="1" lastCol="1" bandRow="1" bandCol="1">' +
    '<a:tableStyleId>' +
    MEDIUM_2_ACCENT_1 +
    '</a:tableStyleId></a:tblPr>',
  columns: [1700000, 1700000, 1700000],
  rows: styledRows('tableStyleId'),
});

const BY_INLINE = tbl({
  tblPr:
    '<a:tblPr rtl="0" firstRow="1" lastRow="1" firstCol="1" lastCol="1" bandRow="1" bandCol="1">' +
    INLINE_STYLE +
    '</a:tblPr>',
  columns: [1700000, 1700000, 1700000],
  rows: styledRows('inline a:tableStyle'),
});

export const a20Tables: ProbeDeck = {
  id: 'a20-tables',
  title: 'PPTX Studio corpus: a20 tables',
  description:
    'Tables inside p:graphicFrame: a four-by-four occupancy grid with gridSpan, hMerge, rowSpan, ' +
    'vMerge and one 2x2 block that carries both axes at once; every a:tcPr child in schema order - ' +
    'the six borders, the four margins, anchor, vert, horzOverflow, four fills and a:cell3D; and ' +
    'both sources of a table style side by side, a:tableStyleId against the empty stock ' +
    'tableStyles.xml and a full inline a:tableStyle carrying all three ST_OnOffStyleType values.',
  features: {
    shape: 6,
    placeholder: 6,
    graphicFrame: 4,
    table: 4,
    gradientFill: 3,
    patternFill: 1,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a20 tables',
    slides: [
      {
        title: 'a20 — the occupancy grid: gridSpan, hMerge, rowSpan, vMerge',
        body: frame(10, 'Merged table', 1016000, 1600200, COL * 4, ROW * 4, MERGES),
      },
      {
        title: 'a20 — a:tcPr: borders, margins, anchoring, vertical text, fills',
        body: frame(10, 'Cell properties table', 1016000, 1600200, COL * 4, ROW * 4, BORDERS),
      },
      {
        title: 'a20 — a table style by id, and the same table style inline',
        body:
          frame(10, 'Styled by id', 914400, 1600200, 5100000, ROW * 4, BY_ID) +
          frame(11, 'Styled inline', 6172200, 1600200, 5100000, ROW * 4, BY_INLINE),
      },
    ],
  }),
};

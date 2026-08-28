import {
  chartColorStyle,
  chartStyle,
  CT_CHART_COLORS,
  CT_CHART_STYLE,
  REL_CHART_COLORS,
  REL_CHART_STYLE,
} from '../chart-style.ts';
import { DECLARATION, NS_A, NS_R, relsXml, REL, type ProbePart } from '../package.ts';
import { graphicFrame, GRAPHIC_URI } from '../shapes.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Charts: four of them, and not one embedded workbook between them.
 *
 * Every element here was checked against a chart PowerPoint 16.0.20326 wrote on
 * 2026-08-27 through `Shapes.AddChart2`, so the shapes are its shapes rather
 * than a reading of the schema. What that deck does and this one deliberately
 * does not is carry `ppt/embeddings/Microsoft_Excel_Worksheet.xlsx` and the
 * `c:externalData` that points at it.
 *
 * ## Why the workbook is absent on purpose
 *
 * A chart is rendered entirely from the `c:numCache` and `c:strCache` copies
 * sitting inside each series. `c:f` - `Sheet1!$B$2:$B$5` - is a provenance
 * note, not a data source, and opening the workbook to render is both slow and
 * wrong. Leaving the workbook out makes that structural: nothing here *can* be
 * rendered by reaching for the spreadsheet, because there is no spreadsheet.
 * The `c:f` formulas are still written, pointing at a sheet that does not
 * exist, which is exactly the state a renderer must cope with.
 *
 * ## `c:pt/@idx` is sparse and authoritative
 *
 * The first chart's third series has `ptCount="6"` and four `c:pt` children, at
 * idx 0, 1, 3 and 5. Categories 3 and 5 are blank cells. Read the children
 * positionally and every value after the first gap lands under the wrong
 * category - the whole series shifts left, silently, and looks like data. The
 * companion is `c:dispBlanksAs`, which says what a gap means: `gap`, `zero` or
 * `span`. This deck writes `gap` on one chart and `zero` on another.
 *
 * ## Where a series colour comes from, in the order it is consulted
 *
 * theme -> `colors1.xml` -> `style1.xml` -> `c:ser/c:spPr`. The trap is that
 * the last of those is usually **absent**, so an implementation that reads only
 * `c:ser/c:spPr` and falls back to a palette of its own gets every chart wrong
 * while looking correct on the one chart it was tested against.
 *
 * Chart 1 carries both parts, from `tools/corpus/gen/chart-style.ts`. Charts 2,
 * 3 and 4 carry neither, which is just as real a case - it is what every
 * pre-2011 producer emits - so the two configurations cover both branches.
 *
 * That split is also how this deck found a rule nobody writes down. The first
 * `style1.xml` here held **four** of `cs:chartStyle`'s thirty-one entries -
 * chartArea, dataPoint, legend, plotArea, which is everything the corpus
 * probes - and PowerPoint refused the package outright. Bisection: chart 1 with
 * no `.rels`, opens; with a `.rels` naming only the colour style, opens; with
 * the four-entry chart style, refused; with all thirty-one entries present,
 * opens. **A partial `cs:chartStyle` is a whole-package refusal**, and the
 * message names no part.
 *
 * ## Two things that are choices, and one of them is not `c:strLit`
 *
 * `c:cat` accepts `multiLvlStrRef`, `numRef`, `numLit`, `strRef` or `strLit`,
 * and `c:val` accepts `numRef` or `numLit`. `c:tx` looks like it should join
 * them and does not: `CT_SerTx` is `c:strRef` **or** `c:v`, nothing else.
 * Chart 3 first wrote `<c:tx><c:strLit>…</c:strLit></c:tx>` and that too was a
 * whole-package refusal, found by the same loop.
 *
 * ## `c:plotArea` may hold more than one chart group
 *
 * Chart 2 is a bar and a line in one plot area against a **secondary** value
 * axis: four `c:axId` values, two axis pairs, and the second category axis
 * carrying `c:delete val="1"` because it is never drawn. Code that takes the
 * plot area's first chart-group child silently loses the line overlay, which is
 * the most common way a combo chart renders as half of itself.
 *
 * ## Literal data, which has no formula at all
 *
 * Chart 3 uses `c:strLit` and `c:numLit` rather than `c:strRef`/`c:numRef`.
 * They are the same cache shape with no `c:f` - the case where there genuinely
 * is no spreadsheet behind the chart - and a parser that reaches for `c:f`
 * before looking at what it has throws on a chart that is perfectly legal.
 */

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

// ------------------------------------------------------------------ builders

const val = (tag: string, value: string | number): string => `<c:${tag} val="${String(value)}"/>`;

/** `c:pt` children from a sparse index map, in ascending index order. */
const points = (entries: ReadonlyArray<readonly [number, string]>): string =>
  entries.map(([idx, v]) => `<c:pt idx="${String(idx)}"><c:v>${v}</c:v></c:pt>`).join('');

/** A cached string reference: the formula, then the copy that is actually read. */
const strRef = (
  formula: string,
  count: number,
  pts: ReadonlyArray<readonly [number, string]>,
): string =>
  '<c:strRef>' +
  `<c:f>${formula}</c:f>` +
  '<c:strCache>' +
  val('ptCount', count) +
  points(pts) +
  '</c:strCache></c:strRef>';

const numRef = (
  formula: string,
  count: number,
  pts: ReadonlyArray<readonly [number, string]>,
  formatCode = 'General',
): string =>
  '<c:numRef>' +
  `<c:f>${formula}</c:f>` +
  '<c:numCache>' +
  `<c:formatCode>${formatCode}</c:formatCode>` +
  val('ptCount', count) +
  points(pts) +
  '</c:numCache></c:numRef>';

/** `c:strLit` - the same cache with no formula behind it. */
const strLit = (count: number, pts: ReadonlyArray<readonly [number, string]>): string =>
  '<c:strLit>' + val('ptCount', count) + points(pts) + '</c:strLit>';

const numLit = (count: number, pts: ReadonlyArray<readonly [number, string]>): string =>
  '<c:numLit>' +
  '<c:formatCode>General</c:formatCode>' +
  val('ptCount', count) +
  points(pts) +
  '</c:numLit>';

const seriesFill = (accent: string): string =>
  `<c:spPr><a:solidFill><a:schemeClr val="${accent}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`;

const seriesLine = (accent: string): string =>
  `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="${accent}"/></a:solidFill>` +
  '<a:round/></a:ln><a:effectLst/></c:spPr>';

/** A category axis. `CT_CatAx` order is fixed and `c:crossAx` comes late. */
function catAx(spec: {
  readonly id: number;
  readonly crossAx: number;
  readonly deleted?: boolean;
  readonly pos?: string;
}): string {
  return (
    '<c:catAx>' +
    val('axId', spec.id) +
    '<c:scaling>' +
    val('orientation', 'minMax') +
    '</c:scaling>' +
    val('delete', spec.deleted === true ? 1 : 0) +
    val('axPos', spec.pos ?? 'b') +
    '<c:numFmt formatCode="General" sourceLinked="1"/>' +
    val('majorTickMark', 'none') +
    val('minorTickMark', 'none') +
    val('tickLblPos', 'nextTo') +
    val('crossAx', spec.crossAx) +
    val('crosses', 'autoZero') +
    val('auto', 1) +
    val('lblAlgn', 'ctr') +
    val('lblOffset', 100) +
    val('noMultiLvlLbl', 0) +
    '</c:catAx>'
  );
}

function valAx(spec: {
  readonly id: number;
  readonly crossAx: number;
  readonly pos: string;
  readonly crosses: string;
  readonly gridlines?: boolean;
  readonly logBase?: number;
  readonly orientation?: string;
}): string {
  return (
    '<c:valAx>' +
    val('axId', spec.id) +
    '<c:scaling>' +
    (spec.logBase === undefined ? '' : val('logBase', spec.logBase)) +
    val('orientation', spec.orientation ?? 'minMax') +
    '</c:scaling>' +
    val('delete', 0) +
    val('axPos', spec.pos) +
    (spec.gridlines === true ? '<c:majorGridlines/>' : '') +
    '<c:numFmt formatCode="General" sourceLinked="1"/>' +
    val('majorTickMark', 'out') +
    val('minorTickMark', 'none') +
    val('tickLblPos', 'nextTo') +
    val('crossAx', spec.crossAx) +
    val('crosses', spec.crosses) +
    val('crossBetween', 'between') +
    '</c:valAx>'
  );
}

/**
 * A whole `c:chartSpace`.
 *
 * `CT_ChartSpace` is date1904, lang, roundedCorners, style, clrMapOvr,
 * pivotSource, protection, chart, spPr, txPr, externalData, printSettings,
 * userShapes, extLst - and this writes the first three, the chart, and nothing
 * after it, because `c:externalData` is the workbook this deck does not have.
 */
function chartSpace(spec: {
  readonly plotArea: string;
  readonly legendPos?: string;
  readonly dispBlanksAs?: string;
  readonly title?: string;
}): string {
  return (
    DECLARATION +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    val('date1904', 0) +
    val('lang', 'en-GB') +
    val('roundedCorners', 0) +
    '<c:chart>' +
    (spec.title === undefined
      ? val('autoTitleDeleted', 1)
      : '<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis"' +
        ' vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>' +
        `<a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" baseline="0"/></a:pPr>` +
        `<a:r><a:rPr lang="en-GB"/><a:t>${spec.title}</a:t></a:r></a:p>` +
        '</c:rich></c:tx>' +
        val('overlay', 0) +
        '</c:title>' +
        val('autoTitleDeleted', 0)) +
    `<c:plotArea><c:layout/>${spec.plotArea}</c:plotArea>` +
    (spec.legendPos === undefined
      ? ''
      : `<c:legend>${val('legendPos', spec.legendPos)}${val('overlay', 0)}</c:legend>`) +
    val('plotVisOnly', 1) +
    val('dispBlanksAs', spec.dispBlanksAs ?? 'gap') +
    '</c:chart>' +
    '</c:chartSpace>'
  );
}

// -------------------------------------------------------- chart 1: clustered

const CATEGORIES: ReadonlyArray<readonly [number, string]> = [
  [0, 'Jan'],
  [1, 'Feb'],
  [2, 'Mar'],
  [3, 'Apr'],
  [4, 'May'],
  [5, 'Jun'],
];

const barSeries = (
  index: number,
  name: string,
  accent: string,
  values: ReadonlyArray<readonly [number, string]>,
): string =>
  '<c:ser>' +
  val('idx', index) +
  val('order', index) +
  `<c:tx>${strRef('Sheet1!$' + String.fromCharCode(66 + index) + '$1', 1, [[0, name]])}</c:tx>` +
  seriesFill(accent) +
  val('invertIfNegative', 0) +
  `<c:cat>${strRef('Sheet1!$A$2:$A$7', 6, CATEGORIES)}</c:cat>` +
  `<c:val>${numRef('Sheet1!$' + String.fromCharCode(66 + index) + '$2:$' + String.fromCharCode(66 + index) + '$7', 6, values)}</c:val>` +
  '</c:ser>';

const CHART_1 = chartSpace({
  title: 'Clustered bar, and a series with holes in it',
  legendPos: 'b',
  dispBlanksAs: 'gap',
  plotArea:
    '<c:barChart>' +
    val('barDir', 'col') +
    val('grouping', 'clustered') +
    val('varyColors', 0) +
    barSeries(0, 'Series 1', 'accent1', [
      [0, '4.3'],
      [1, '2.5'],
      [2, '3.5'],
      [3, '4.5'],
      [4, '2.2'],
      [5, '5.1'],
    ]) +
    barSeries(1, 'Series 2', 'accent2', [
      [0, '2.4'],
      [1, '4.4'],
      [2, '1.8'],
      [3, '2.8'],
      [4, '3.9'],
      [5, '1.2'],
    ]) +
    // Six categories, four points: idx 2 and 4 are blank cells. A positional
    // read puts May's value under April and June's under May.
    barSeries(2, 'Series 3, sparse', 'accent3', [
      [0, '2.0'],
      [1, '2.0'],
      [3, '3.0'],
      [5, '5.0'],
    ]) +
    val('gapWidth', 219) +
    val('overlap', -27) +
    val('axId', 111111111) +
    val('axId', 222222222) +
    '</c:barChart>' +
    catAx({ id: 111111111, crossAx: 222222222 }) +
    valAx({ id: 222222222, crossAx: 111111111, pos: 'l', crosses: 'autoZero', gridlines: true }),
});

// ------------------------------------------------------------ chart 2: combo

const CHART_2 = chartSpace({
  title: 'A bar and a line, on two value axes',
  legendPos: 'b',
  dispBlanksAs: 'zero',
  plotArea:
    '<c:barChart>' +
    val('barDir', 'col') +
    val('grouping', 'clustered') +
    val('varyColors', 0) +
    '<c:ser>' +
    val('idx', 0) +
    val('order', 0) +
    `<c:tx>${strRef('Sheet1!$B$1', 1, [[0, 'Volume']])}</c:tx>` +
    seriesFill('accent1') +
    val('invertIfNegative', 0) +
    `<c:cat>${strRef('Sheet1!$A$2:$A$7', 6, CATEGORIES)}</c:cat>` +
    `<c:val>${numRef('Sheet1!$B$2:$B$7', 6, [
      [0, '120'],
      [1, '145'],
      [2, '132'],
      [3, '168'],
      [4, '151'],
      [5, '190'],
    ])}</c:val>` +
    '</c:ser>' +
    val('gapWidth', 150) +
    val('axId', 111111111) +
    val('axId', 222222222) +
    '</c:barChart>' +
    // Second chart group, second axis pair. Both are children of one plot area.
    '<c:lineChart>' +
    val('grouping', 'standard') +
    val('varyColors', 0) +
    '<c:ser>' +
    val('idx', 1) +
    val('order', 1) +
    `<c:tx>${strRef('Sheet1!$C$1', 1, [[0, 'Rate %']])}</c:tx>` +
    seriesLine('accent4') +
    '<c:marker>' +
    val('symbol', 'circle') +
    val('size', 6) +
    '</c:marker>' +
    `<c:cat>${strRef('Sheet1!$A$2:$A$7', 6, CATEGORIES)}</c:cat>` +
    `<c:val>${numRef(
      'Sheet1!$C$2:$C$7',
      6,
      [
        [0, '3.1'],
        [1, '3.4'],
        [2, '2.9'],
        [3, '4.1'],
        [4, '3.8'],
        [5, '4.6'],
      ],
      '0.0%',
    )}</c:val>` +
    val('smooth', 0) +
    '</c:ser>' +
    val('marker', 1) +
    val('axId', 333333333) +
    val('axId', 444444444) +
    '</c:lineChart>' +
    catAx({ id: 111111111, crossAx: 222222222 }) +
    valAx({ id: 222222222, crossAx: 111111111, pos: 'l', crosses: 'autoZero', gridlines: true }) +
    // The secondary category axis is never drawn - `c:delete val="1"` - but it
    // must exist, because a chart group names two axes and both must resolve.
    catAx({ id: 333333333, crossAx: 444444444, deleted: true, pos: 'b' }) +
    valAx({ id: 444444444, crossAx: 333333333, pos: 'r', crosses: 'max' }),
});

// ------------------------------------------------- chart 3: pie, literal data

const CHART_3 = chartSpace({
  title: 'Pie, with literal data and an exploded slice',
  legendPos: 'r',
  plotArea:
    '<c:pieChart>' +
    val('varyColors', 1) +
    '<c:ser>' +
    val('idx', 0) +
    val('order', 0) +
    // `CT_SerTx` is a choice of `c:strRef` or `c:v`, and **not** `c:strLit`.
    // Writing a literal here was a whole-package refusal: see the file comment.
    '<c:tx><c:v>Share</c:v></c:tx>' +
    // One data point pulled out of the pie, and one recoloured.
    '<c:dPt>' +
    val('idx', 1) +
    val('bubble3D', 0) +
    val('explosion', 20) +
    seriesFill('accent5') +
    '</c:dPt>' +
    `<c:cat>${strLit(4, [
      [0, 'North'],
      [1, 'South'],
      [2, 'East'],
      [3, 'West'],
    ])}</c:cat>` +
    `<c:val>${numLit(4, [
      [0, '41'],
      [1, '27'],
      [2, '19'],
      [3, '13'],
    ])}</c:val>` +
    '</c:ser>' +
    // 90 degrees clockwise from twelve o'clock, so a renderer that starts at
    // zero puts every slice in the wrong quadrant.
    val('firstSliceAng', 90) +
    '</c:pieChart>',
});

// ------------------------------------------- chart 4: scatter, trend, error

const CHART_4 = chartSpace({
  title: 'Scatter, with a trendline and error bars',
  plotArea:
    '<c:scatterChart>' +
    val('scatterStyle', 'lineMarker') +
    val('varyColors', 0) +
    '<c:ser>' +
    val('idx', 0) +
    val('order', 0) +
    `<c:tx>${strRef('Sheet1!$B$1', 1, [[0, 'Observed']])}</c:tx>` +
    '<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr>' +
    '<c:marker>' +
    val('symbol', 'circle') +
    val('size', 5) +
    seriesFill('accent2') +
    '</c:marker>' +
    // `CT_ScatterSer` puts trendline and errBars before xVal and yVal.
    '<c:trendline>' +
    '<c:name>Linear fit</c:name>' +
    '<c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill>' +
    '<a:prstDash val="sysDot"/></a:ln></c:spPr>' +
    val('trendlineType', 'linear') +
    val('dispRSqr', 1) +
    val('dispEq', 1) +
    '</c:trendline>' +
    '<c:errBars>' +
    val('errDir', 'y') +
    val('errBarType', 'both') +
    val('errValType', 'fixedVal') +
    val('noEndCap', 0) +
    val('val', 0.4) +
    '</c:errBars>' +
    `<c:xVal>${numRef('Sheet1!$A$2:$A$7', 6, [
      [0, '1'],
      [1, '2'],
      [2, '3'],
      [3, '4'],
      [4, '5'],
      [5, '6'],
    ])}</c:xVal>` +
    `<c:yVal>${numRef('Sheet1!$B$2:$B$7', 6, [
      [0, '2.1'],
      [1, '3.9'],
      [2, '6.4'],
      [3, '7.8'],
      [4, '10.5'],
      [5, '11.9'],
    ])}</c:yVal>` +
    val('smooth', 0) +
    '</c:ser>' +
    val('axId', 555555555) +
    val('axId', 666666666) +
    '</c:scatterChart>' +
    valAx({ id: 555555555, crossAx: 666666666, pos: 'b', crosses: 'autoZero' }) +
    valAx({ id: 666666666, crossAx: 555555555, pos: 'l', crosses: 'autoZero', gridlines: true }),
});

// ------------------------------------------------------------------ assembly

const CHARTS = [CHART_1, CHART_2, CHART_3, CHART_4];

const PARTS: readonly ProbePart[] = [
  ...CHARTS.map((bytes, index): ProbePart => ({
    name: `ppt/charts/chart${String(index + 1)}.xml`,
    bytes,
    contentType: { kind: 'override', type: CT_CHART },
  })),
  {
    name: 'ppt/charts/colors1.xml',
    bytes: chartColorStyle(),
    contentType: { kind: 'override', type: CT_CHART_COLORS },
  },
  {
    name: 'ppt/charts/style1.xml',
    bytes: chartStyle(),
    contentType: { kind: 'override', type: CT_CHART_STYLE },
  },
  // Only chart 1 has relationships, so only chart 1 has a `.rels` part. The
  // rId order is PowerPoint's: style first, then colours.
  {
    name: 'ppt/charts/_rels/chart1.xml.rels',
    bytes: relsXml([
      { id: 'rId1', type: REL_CHART_STYLE, target: 'style1.xml' },
      { id: 'rId2', type: REL_CHART_COLORS, target: 'colors1.xml' },
    ]),
  },
];

/** The slide-side half: a frame whose `a:graphicData` is one `c:chart` stub. */
const chartFrame = (
  id: number,
  name: string,
  relId: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
): string =>
  graphicFrame({
    id,
    name,
    x,
    y,
    cx,
    cy,
    uri: GRAPHIC_URI.chart,
    // Everything about the chart is in the part. This element is the edge.
    content: `<c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${relId}"/>`,
  });

const chartRel = (relId: string, number: number): { id: string; type: string; target: string } => ({
  id: relId,
  type: REL + 'chart',
  target: `../charts/chart${String(number)}.xml`,
});

export const a21Charts: ProbeDeck = {
  id: 'a21-charts',
  title: 'PPTX Studio corpus: a21 charts',
  description:
    'Four c:chartSpace parts and no embedded workbook: a clustered bar whose third series has a ' +
    'sparse c:pt/@idx with two blank categories, a bar-plus-line combo on a secondary axis with ' +
    'four c:axId values and a deleted category axis, a pie built from c:strLit and c:numLit with ' +
    'an exploded slice and c:firstSliceAng, and a scatter with a trendline and error bars. Chart 1 ' +
    'carries the measured colors1.xml and a four-entry style1.xml; the other three carry neither, ' +
    'which is the case every pre-2011 producer emits.',
  features: {
    shape: 6,
    placeholder: 6,
    gradientFill: 2,
    graphicFrame: 4,
    chart: 4,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a21 charts',
    parts: PARTS,
    slides: [
      {
        title: 'a21 — a clustered bar, and a series with holes in it',
        rels: [chartRel('rId2', 1)],
        body: chartFrame(10, 'Chart 1', 'rId2', 914400, 1600200, 10363200, 4572000),
      },
      {
        title: 'a21 — a combo chart on two value axes',
        rels: [chartRel('rId2', 2)],
        body: chartFrame(10, 'Chart 2', 'rId2', 914400, 1600200, 10363200, 4572000),
      },
      {
        title: 'a21 — literal data, and a scatter with a fit',
        rels: [chartRel('rId2', 3), chartRel('rId3', 4)],
        body:
          chartFrame(10, 'Chart 3', 'rId2', 685800, 1600200, 5029200, 4572000) +
          chartFrame(11, 'Chart 4', 'rId3', 6477000, 1600200, 5029200, 4572000),
      },
    ],
  }),
};

import {
  chartColorStyle,
  chartStyle,
  CT_CHART_COLORS,
  CT_CHART_STYLE,
  REL_CHART_COLORS,
  REL_CHART_STYLE,
} from '../../markup/chart-style.ts';
import { DECLARATION, NS_A, NS_R, relsXml, type ProbePart } from '../../markup/chassis.ts';
import { quadrantPng } from '../../assets/png.ts';
import { graphicFrame, GRAPHIC_URI, picture } from '../../markup/shapes.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * ChartEx: seven of the eight `layoutId` values, and a different data model.
 *
 * Every one of these was authored through `Shapes.AddChart2` on 2026-08-27 and
 * read back, so the shapes are PowerPoint 16.0.20326's own. The eighth,
 * `regionMap`, is a declared gap - a region map's `cx:geography` is resolved
 * against an online service, so authoring one is a download.
 *
 * ## It is not a chart with extra layouts, it is a different model
 *
 * `c:chartSpace` hangs the data off each series: every `c:ser` carries its own
 * `c:cat` and `c:val`. `cx:chartSpace` inverts that. The data is a **list of
 * dimensions at the top of the part**, in `cx:chartData`, and a series points
 * into it by number with `<cx:dataId val="0"/>`. So a series can carry no data
 * of its own at all - `paretoLine` in chart 6 has no `cx:dataId`, only
 * `@ownerIdx="0"`, and reads the series it is drawn over.
 *
 * The dimension's `@type` is not fixed either:
 *
 * | layout                  | dimensions                            |
 * | ----------------------- | ------------------------------------- |
 * | waterfall, funnel, box  | `strDim type="cat"`, `numDim type="val"`  |
 * | treemap, sunburst       | `strDim type="cat"`, `numDim type="size"` |
 * | histogram, pareto       | `numDim type="val"` alone             |
 *
 * `size` rather than `val` on a treemap is the sort of thing that reads as a
 * typo and is not; a renderer keyed on `type="val"` draws an empty treemap.
 *
 * A hierarchy is **several `cx:lvl` children of one `cx:strDim`, deepest
 * first** - leaf, then stem, then branch - each with the same `ptCount` and one
 * entry per leaf. Not a tree. Chart 1 has three levels, which is what a real
 * treemap looks like.
 *
 * ## `cx:layoutPr` is where each layout keeps its own vocabulary
 *
 * | layout           | `cx:layoutPr` holds                                    |
 * | ---------------- | ------------------------------------------------------ |
 * | waterfall        | `cx:subtotals` with one `cx:idx` per subtotal column   |
 * | treemap          | `cx:parentLabelLayout val="overlapping"`               |
 * | histogram        | `cx:binning intervalClosed="r"`                        |
 * | box and whisker  | `cx:visibility` and `cx:statistics quartileMethod=…`   |
 * | pareto           | `cx:aggregation` on the column series                  |
 *
 * ## Two frames, and the second is not a frame at all
 *
 * PowerPoint writes a ChartEx inside `mc:AlternateContent`: the `mc:Choice`
 * requires `cx1` - **`http://schemas.microsoft.com/office/drawing/2015/9/8/chartex`**,
 * which is not the `cx` namespace the chart itself is in and exists only to be
 * named in `@Requires` - and the `mc:Fallback` holds a `p:pic` of a PNG
 * PowerPoint rasterised at save time. So every ChartEx in a real deck ships
 * with a picture of itself, and a reader that resolves MCE wrongly renders a
 * flat image that looks almost right.
 *
 * Charts 1 and 2 here are written that way. Charts 3 to 7 are bare
 * `p:graphicFrame`s with no switch and no fallback, which is equally legal and
 * is what any producer that is not PowerPoint emits.
 *
 * ## A ChartEx part with no relationships is a whole-package refusal
 *
 * Every one of these first shipped without a `.rels` of its own, the way a
 * classic `c:chartSpace` may, and PowerPoint refused all seven identically.
 * The bisection that settled it is the strongest kind available: `chartEx1.xml`
 * was lifted **byte for byte** out of the deck PowerPoint had just written,
 * dropped into a package with nothing else changed, and refused. Adding back
 * its `chartStyle` and `chartColorStyle` relationships made the same package
 * open, and so did adding them to ours.
 *
 * So `cx:chartSpace` requires the pair that `c:chartSpace` merely likes -
 * `a21-charts`'s charts 2, 3 and 4 carry no relationships and open. Both parts
 * come from `tools/corpus/tiers/a-generated/markup/chart-style.ts`, shared across all seven charts
 * rather than duplicated per chart the way PowerPoint writes them, because what
 * is being probed is the edge and not the palette.
 *
 * ## One finding that belongs to another deck
 *
 * In the measured package **`chartEx1.xml` wrote 12 of its 12 self-closing tags
 * spaced** - `<cx:title pos="t" align="ctr" overlay="0" />` - while the other
 * 3,046 self-closing tags across the same 60 parts had no space at all. There
 * is more than one XML serializer inside PowerPoint and the ChartEx one is the
 * odd member. ADR 0005 measured 0 spaced out of 1,305 and concluded no producer
 * available to this project emits the form; that conclusion is now wrong, and
 * `a36-spaced-tags` says so. This deck writes the unspaced form, because the
 * corpus generator has one serializer and `conventions.test.ts` holds it to it.
 */

const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
/** The MCE-only namespace. Named in `@Requires`, never used for an element. */
const NS_CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const REL_CHART_EX = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const CT_CHART_EX = 'application/vnd.ms-office.chartex+xml';

// ------------------------------------------------------------------ builders

/** `cx:pt` children. The value is the element's text, not a `@val`. */
const pts = (values: readonly string[]): string =>
  values.map((v, i) => `<cx:pt idx="${String(i)}">${v}</cx:pt>`).join('');

const lvl = (values: readonly string[], formatCode?: string): string =>
  `<cx:lvl ptCount="${String(values.length)}"` +
  (formatCode === undefined ? '' : ` formatCode="${formatCode}"`) +
  `>${pts(values)}</cx:lvl>`;

/** A category dimension. Several levels means a hierarchy, deepest **first**. */
const strDim = (levels: ReadonlyArray<readonly string[]>): string =>
  `<cx:strDim type="cat">${levels.map((values) => lvl(values)).join('')}</cx:strDim>`;

const numDim = (type: 'val' | 'size', values: readonly string[]): string =>
  `<cx:numDim type="${type}">${lvl(values, 'General')}</cx:numDim>`;

const data = (id: number, body: string): string => `<cx:data id="${String(id)}">${body}</cx:data>`;

const catAxis = (id: number, gapWidth: string): string =>
  `<cx:axis id="${String(id)}"><cx:catScaling gapWidth="${gapWidth}"/><cx:tickLabels/></cx:axis>`;

const valAxis = (id: number, extra = ''): string =>
  `<cx:axis id="${String(id)}"><cx:valScaling${extra === '' ? '/>' : `${extra}`}` +
  '<cx:majorGridlines/><cx:tickLabels/></cx:axis>';

function chartExSpace(spec: {
  readonly data: string;
  readonly series: string;
  readonly axes?: string;
  readonly legend?: boolean;
}): string {
  return (
    DECLARATION +
    `<cx:chartSpace xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:cx="${NS_CX}">` +
    `<cx:chartData>${spec.data}</cx:chartData>` +
    '<cx:chart>' +
    '<cx:title pos="t" align="ctr" overlay="0"/>' +
    '<cx:plotArea>' +
    `<cx:plotAreaRegion>${spec.series}</cx:plotAreaRegion>` +
    (spec.axes ?? '') +
    '</cx:plotArea>' +
    (spec.legend === true ? '<cx:legend pos="t" align="ctr" overlay="0"/>' : '') +
    '</cx:chart>' +
    '</cx:chartSpace>'
  );
}

/**
 * One `cx:series`.
 *
 * `@uniqueId` is a `ST_Guid` PowerPoint allocates per series. These are fixed
 * literals rather than generated, because a corpus deck whose bytes move on
 * every build is not a fixture.
 */
function series(spec: {
  readonly layoutId: string;
  readonly uniqueId: string;
  readonly dataId?: number;
  readonly ownerIdx?: number;
  readonly dataLabels?: string;
  readonly layoutPr?: string;
  readonly axisIds?: readonly number[];
}): string {
  return (
    `<cx:series layoutId="${spec.layoutId}"` +
    (spec.ownerIdx === undefined ? '' : ` ownerIdx="${String(spec.ownerIdx)}"`) +
    ` uniqueId="${spec.uniqueId}">` +
    (spec.dataLabels ?? '') +
    (spec.dataId === undefined ? '' : `<cx:dataId val="${String(spec.dataId)}"/>`) +
    (spec.layoutPr === undefined ? '' : `<cx:layoutPr>${spec.layoutPr}</cx:layoutPr>`) +
    (spec.axisIds ?? []).map((id) => `<cx:axisId val="${String(id)}"/>`).join('') +
    '</cx:series>'
  );
}

const labels = (attributes: string, pos?: string): string =>
  `<cx:dataLabels${pos === undefined ? '' : ` pos="${pos}"`}>` +
  `<cx:visibility ${attributes}/></cx:dataLabels>`;

// ------------------------------------------------------------- the seven

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug'];

/** 1. Treemap: three category levels and a `size` dimension. */
const TREEMAP = chartExSpace({
  data: data(
    0,
    strDim([
      ['Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6'],
      ['Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', 'Stem 3', 'Stem 3'],
      ['Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2'],
    ]) + numDim('size', ['22', '12', '18', '87', '25', '17']),
  ),
  series: series({
    layoutId: 'treemap',
    uniqueId: '{3F6FE1FC-70DC-47AA-B8D9-DEC1D7214DF0}',
    dataId: 0,
    dataLabels: labels('seriesName="0" categoryName="1" value="0"', 'inEnd'),
    layoutPr: '<cx:parentLabelLayout val="overlapping"/>',
  }),
  legend: true,
});

/** 2. Sunburst: the same data, a different `layoutId`, and no legend. */
const SUNBURST = chartExSpace({
  data: data(
    0,
    strDim([
      ['Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6'],
      ['Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', 'Stem 3', 'Stem 3'],
      ['Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2'],
    ]) + numDim('size', ['22', '12', '18', '87', '25', '17']),
  ),
  series: series({
    layoutId: 'sunburst',
    uniqueId: '{5C1BEA33-8A2C-4E1F-9A6A-1D8DA9B1B0C4}',
    dataId: 0,
    dataLabels: labels('seriesName="0" categoryName="1" value="0"'),
  }),
});

/** 3. Waterfall: `cx:subtotals` naming the columns that are running totals. */
const WATERFALL = chartExSpace({
  data: data(
    0,
    strDim([MONTHS]) + numDim('val', ['100', '20', '50', '-40', '130', '-60', '70', '140']),
  ),
  series: series({
    layoutId: 'waterfall',
    uniqueId: '{65EBC1B9-E930-4B01-8A4D-4F46E1A88413}',
    dataId: 0,
    dataLabels: labels('seriesName="0" categoryName="0" value="1"', 'outEnd'),
    layoutPr: '<cx:subtotals><cx:idx val="0"/><cx:idx val="4"/><cx:idx val="7"/></cx:subtotals>',
  }),
  axes: catAxis(0, '0.5') + valAxis(1),
  legend: true,
});

/** 4. Funnel: one category axis and no value axis at all. */
const FUNNEL = chartExSpace({
  data: data(
    0,
    strDim([['Awareness', 'Interest', 'Trial', 'Purchase', 'Loyalty']]) +
      numDim('val', ['5200', '3100', '1800', '900', '420']),
  ),
  series: series({
    layoutId: 'funnel',
    uniqueId: '{6164DCB1-9E7A-4C8F-A427-B9E2DCB079FA}',
    dataId: 0,
    dataLabels: labels('seriesName="0" categoryName="0" value="1"'),
  }),
  axes: `<cx:axis id="1"><cx:catScaling gapWidth="0.0599999987"/><cx:tickLabels/></cx:axis>`,
});

/** 5. Box and whisker: three series over three dimensions, one box each. */
const BOX_WHISKER = chartExSpace({
  data:
    data(0, strDim([['A', 'A', 'A', 'A', 'A']]) + numDim('val', ['12', '19', '14', '31', '17'])) +
    data(1, strDim([['B', 'B', 'B', 'B', 'B']]) + numDim('val', ['22', '25', '41', '28', '24'])) +
    data(2, strDim([['C', 'C', 'C', 'C', 'C']]) + numDim('val', ['8', '15', '11', '13', '35'])),
  series:
    series({
      layoutId: 'boxWhisker',
      uniqueId: '{46BB7A51-7ACE-4BFA-B8E1-030F00016C21}',
      dataId: 0,
      layoutPr:
        '<cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>' +
        '<cx:statistics quartileMethod="exclusive"/>',
    }) +
    series({
      layoutId: 'boxWhisker',
      uniqueId: '{DBC18700-12DB-4BA2-B86E-2EFB70394D4E}',
      dataId: 1,
      layoutPr:
        '<cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="1"/>' +
        '<cx:statistics quartileMethod="inclusive"/>',
    }) +
    series({
      layoutId: 'boxWhisker',
      uniqueId: '{8D4233EB-1C77-4871-AF86-2C7198302D47}',
      dataId: 2,
      layoutPr:
        '<cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="0"/>' +
        '<cx:statistics quartileMethod="exclusive"/>',
    }),
  axes: catAxis(0, '1') + valAxis(1),
});

/** 6. Histogram: no category dimension at all, and `cx:binning`. */
const HISTOGRAM = chartExSpace({
  data: data(
    0,
    numDim('val', ['12', '19', '14', '31', '17', '22', '25', '41', '28', '24', '8', '15']),
  ),
  series: series({
    layoutId: 'clusteredColumn',
    uniqueId: '{36DD84EA-09C6-4FAF-ABEB-F401A774CEF4}',
    dataId: 0,
    layoutPr: '<cx:binning intervalClosed="r"/>',
  }),
  axes: catAxis(0, '0') + valAxis(1),
});

/**
 * 7. Pareto: two series in one region, and the second borrows the first's data.
 *
 * `paretoLine` has no `cx:dataId`. `@ownerIdx="0"` says it is drawn over series
 * 0, and its `cx:axisId` puts it on axis 2 - a percentage axis with `max="1"`
 * and `cx:units unit="percentage"`, which no classic chart has an equivalent of.
 */
const PARETO = chartExSpace({
  data: data(0, numDim('val', ['41', '27', '19', '13', '9', '6', '4', '2'])),
  series:
    series({
      layoutId: 'clusteredColumn',
      uniqueId: '{C42CF01E-F374-4B7A-BAC9-8971DE638C61}',
      dataId: 0,
      layoutPr: '<cx:aggregation/>',
      axisIds: [1],
    }) +
    series({
      layoutId: 'paretoLine',
      ownerIdx: 0,
      uniqueId: '{A6B8E39A-8D24-42B9-85AA-FBD2BA8F1DFF}',
      axisIds: [2],
    }),
  axes:
    catAxis(0, '0') +
    valAxis(1) +
    '<cx:axis id="2"><cx:valScaling max="1" min="0"/>' +
    '<cx:units unit="percentage"/><cx:tickLabels/></cx:axis>',
});

const CHARTS = [TREEMAP, SUNBURST, WATERFALL, FUNNEL, BOX_WHISKER, HISTOGRAM, PARETO];

const PARTS: readonly ProbePart[] = [
  ...CHARTS.map((bytes, index): ProbePart => ({
    name: `ppt/charts/chartEx${String(index + 1)}.xml`,
    bytes,
    contentType: { kind: 'override', type: CT_CHART_EX },
  })),
  // Every ChartEx part needs relationships of its own. See below.
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
  ...CHARTS.map((_chart, index): ProbePart => ({
    name: `ppt/charts/_rels/chartEx${String(index + 1)}.xml.rels`,
    bytes: relsXml([
      { id: 'rId1', type: REL_CHART_STYLE, target: 'style1.xml' },
      { id: 'rId2', type: REL_CHART_COLORS, target: 'colors1.xml' },
    ]),
  })),
  {
    // One raster for both fallbacks. PowerPoint writes one per chart; sharing
    // it is legal, and what a corpus deck needs is the edge, not the pixels.
    name: 'ppt/media/image1.png',
    bytes: quadrantPng(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
];

// ------------------------------------------------------------------ slides

interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

const BOX: Box = { x: 685800, y: 1600200, cx: 5029200, cy: 4572000 };
const RIGHT: Box = { ...BOX, x: 6477000 };

const bareFrame = (id: number, name: string, relId: string, box: Box): string =>
  graphicFrame({
    id,
    name,
    ...box,
    uri: GRAPHIC_URI.chartEx,
    content: `<cx:chart xmlns:cx="${NS_CX}" xmlns:r="${NS_R}" r:id="${relId}"/>`,
  });

/** The frame PowerPoint writes: a switch, and a picture of the chart under it. */
const switchedFrame = (
  id: number,
  name: string,
  relId: string,
  imageRelId: string,
  box: Box,
): string =>
  `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
  `<mc:Choice xmlns:cx1="${NS_CX1}" Requires="cx1">` +
  bareFrame(id, name, relId, box) +
  '</mc:Choice>' +
  '<mc:Fallback>' +
  picture({ id, name, relId: imageRelId, ...box, description: '' }) +
  '</mc:Fallback>' +
  '</mc:AlternateContent>';

const chartExRel = (
  relId: string,
  number: number,
): { id: string; type: string; target: string } => ({
  id: relId,
  type: REL_CHART_EX,
  target: `../charts/chartEx${String(number)}.xml`,
});

const IMAGE_REL = {
  id: 'rId9',
  type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  target: '../media/image1.png',
};

export const a22ChartEx: ProbeDeck = {
  id: 'a22-chartex',
  title: 'PPTX Studio corpus: a22 chartex',
  description:
    'Seven of the eight ChartEx layoutId values, every one read back from a chart PowerPoint ' +
    '16.0.20326 authored: treemap and sunburst with a three-level cx:strDim and a numDim of type ' +
    'size, waterfall with cx:subtotals, funnel with a category axis and no value axis, three ' +
    'boxWhisker series over three cx:data entries, a histogram with cx:binning and no category ' +
    'dimension at all, and a Pareto whose paretoLine series carries @ownerIdx instead of a ' +
    'cx:dataId. Two are wrapped in the mc:AlternateContent PowerPoint writes, with a rasterised ' +
    'p:pic fallback; five are bare frames. regionMap is a declared gap: its cx:geography is ' +
    'resolved online.',
  features: {
    shape: 6,
    placeholder: 6,
    gradientFill: 2,
    graphicFrame: 7,
    chartEx: 7,
    alternateContent: 2,
    picture: 2,
    presetGeom: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a22 chartex',
    parts: PARTS,
    slides: [
      {
        title: 'a22 — treemap and sunburst, inside the switch PowerPoint writes',
        rels: [chartExRel('rId2', 1), chartExRel('rId3', 2), IMAGE_REL],
        body:
          switchedFrame(10, 'Treemap 1', 'rId2', 'rId9', BOX) +
          switchedFrame(11, 'Sunburst 2', 'rId3', 'rId9', RIGHT),
      },
      {
        title: 'a22 — waterfall and funnel, as bare frames',
        rels: [chartExRel('rId2', 3), chartExRel('rId3', 4)],
        body: bareFrame(10, 'Waterfall 3', 'rId2', BOX) + bareFrame(11, 'Funnel 4', 'rId3', RIGHT),
      },
      {
        title: 'a22 — box and whisker, histogram, and a Pareto',
        rels: [chartExRel('rId2', 5), chartExRel('rId3', 6), chartExRel('rId4', 7)],
        body:
          bareFrame(10, 'Box and whisker 5', 'rId2', { ...BOX, x: 457200, cx: 3505200 }) +
          bareFrame(11, 'Histogram 6', 'rId3', { ...BOX, x: 4343400, cx: 3505200 }) +
          bareFrame(12, 'Pareto 7', 'rId4', { ...BOX, x: 8229600, cx: 3505200 }),
      },
    ],
  }),
};

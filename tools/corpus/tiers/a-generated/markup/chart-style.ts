import { DECLARATION, NS_A } from './package.ts';

/**
 * The two parts every chart in this corpus needs, and one of them has to be
 * complete.
 *
 * Both `c:chartSpace` and `cx:chartSpace` are styled by a pair of Microsoft
 * extension parts related **from the chart part**, not from the slide:
 *
 * ```
 * ppt/charts/_rels/chartN.xml.rels
 *   rId1 -> style1.xml   …/2011/relationships/chartStyle
 *   rId2 -> colors1.xml  …/2011/relationships/chartColorStyle
 * ```
 *
 * ## Two rules, both found by bisection, neither written down anywhere
 *
 * **A `cs:chartStyle` with a subset of its entries is a whole-package
 * refusal.** `a21-charts` first shipped four of the thirty-one - chartArea,
 * dataPoint, legend, plotArea, which is everything the corpus actually probes -
 * and PowerPoint 16.0.20326 refused the file: "The file or directory is
 * corrupted and unreadable." With all thirty-one present it opens. Nothing in
 * any schema says the sequence is required rather than optional, and the
 * failure names no part.
 *
 * **A ChartEx part with no `.rels` at all is a whole-package refusal**, even
 * when the part is one PowerPoint wrote itself. That was measured the strong
 * way: `chartEx1.xml` was lifted byte for byte out of a deck PowerPoint saved,
 * put into a package with nothing else changed, and refused; adding its
 * `colors` and `style` relationships back made the same package open. A classic
 * `c:chartSpace` has no such requirement - `a21`'s charts 2, 3 and 4 carry no
 * relationships and open - so the rule is ChartEx's alone.
 *
 * Which is why this module exists rather than a constant in one deck: two decks
 * need the same two parts for two different reasons, and the reasons are worth
 * writing down once.
 *
 * ## The colour style is measured; the chart style is not
 *
 * `chartColorStyle()` is copied field for field from the `colors1.xml` of a
 * chart PowerPoint authored on 2026-08-27 - `meth="cycle"`, the six accents,
 * and nine `cs:variation` entries whose `lumMod`/`lumOff` pairs are the series
 * palette every Office chart cycles through.
 *
 * `chartStyle()` is **not** a copy. PowerPoint's is 9.7 KB of per-element
 * typography and the corpus has no reason to carry a facsimile of it; this is
 * the same thirty-one entries with uniform, minimal content. What it is
 * faithful about is the one link that matters to sub-phase 9.1:
 * `cs:fillRef/cs:styleClr val="auto"` on `cs:dataPoint`, which is what defers a
 * series fill to `colors1.xml` rather than to `c:ser/c:spPr`.
 */

const NS_CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';

export const REL_CHART_STYLE = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
export const REL_CHART_COLORS =
  'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';

export const CT_CHART_STYLE = 'application/vnd.ms-office.chartstyle+xml';
export const CT_CHART_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml';

/**
 * The thirty-one children of `cs:chartStyle`, in schema order.
 *
 * All thirty-one, because thirty is a refusal. See the file comment.
 */
const ENTRIES = [
  'axisTitle',
  'categoryAxis',
  'chartArea',
  'dataLabel',
  'dataLabelCallout',
  'dataPoint',
  'dataPoint3D',
  'dataPointLine',
  'dataPointMarker',
  'dataPointMarkerLayout',
  'dataPointWireframe',
  'dataTable',
  'downBar',
  'dropLine',
  'errorBar',
  'floor',
  'gridlineMajor',
  'gridlineMinor',
  'hiLoLine',
  'leaderLine',
  'legend',
  'plotArea',
  'plotArea3D',
  'seriesAxis',
  'seriesLine',
  'title',
  'trendline',
  'trendlineLabel',
  'upBar',
  'valueAxis',
  'wall',
] as const;

/** `CT_StyleEntry`: lnRef, fillRef, effectRef, fontRef, spPr, defRPr, bodyPr. */
function styleEntry(name: string): string {
  if (name === 'dataPointMarkerLayout') {
    // The one entry that is not a CT_StyleEntry at all - it is a marker
    // description, and PowerPoint writes it empty.
    return '<cs:dataPointMarkerLayout symbol="circle" size="5"/>';
  }
  const fillRef =
    name === 'dataPoint'
      ? // The link sub-phase 9.1 is about: `auto` defers to the colour style.
        '<cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>'
      : '<cs:fillRef idx="0"/>';
  return (
    `<cs:${name}>` +
    '<cs:lnRef idx="0"/>' +
    fillRef +
    '<cs:effectRef idx="0"/>' +
    '<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>' +
    `</cs:${name}>`
  );
}

/** `styleN.xml`. `@id` 201 is the built-in style PowerPoint defaults to. */
export function chartStyle(): string {
  return (
    DECLARATION +
    `<cs:chartStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" id="201">` +
    ENTRIES.map(styleEntry).join('') +
    '</cs:chartStyle>'
  );
}

/** `colorsN.xml`, copied from the measurement. */
export function chartColorStyle(): string {
  return (
    DECLARATION +
    `<cs:colorStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" meth="cycle" id="10">` +
    ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((name) => `<a:schemeClr val="${name}"/>`)
      .join('') +
    '<cs:variation/>' +
    '<cs:variation><a:lumMod val="60000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="80000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="50000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="70000"/></cs:variation>' +
    '<cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation>' +
    '</cs:colorStyle>'
  );
}

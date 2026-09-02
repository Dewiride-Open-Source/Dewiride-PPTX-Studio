import { CONTENT_TYPE } from '@pptx-studio/opc';
import {
  attribute,
  attributeValue,
  childElements,
  descendantElements,
  namespaceOf,
  NS,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from './context.js';
import { attributeLocation, elementLocation } from './location.js';
import { forEachElement } from './required-rules.js';

/**
 * `V022` … `V026`: schema-legal markup PowerPoint refuses.
 *
 * This is the category that justifies the package. Nothing here can be derived
 * from ECMA-376; every rule is the record of a package built with one change in
 * it, opened, and declined - "PowerPoint could not open the file", or
 * `0x80070570`, "the file or directory is corrupted and unreadable". No log, no
 * part named, no line number. Sub-phase 1.5's `cli bisect` exists because that
 * is the entire diagnostic channel, and these five rules are what it found
 * before it existed, by hand.
 *
 * Which means each of them is falsifiable in a way the schema rules are not: if
 * a later PowerPoint build opens one of these, the rule is wrong and should be
 * deleted. The `why` in `rules.ts` records what was tried and what opened, so
 * that the person who deletes it knows what they are contradicting.
 */

/** Parts where `hdr` and `sldImg` placeholders are legal - they belong to these families. */
const NOTES_FAMILY = new Set<string>([
  CONTENT_TYPE.notesSlide,
  CONTENT_TYPE.notesMaster,
  CONTENT_TYPE.handoutMaster,
]);

/** `p:ph/@type` is not `hdr` or `sldImg` on a slide, layout or master. */
export function v022PlaceholderType(ctx: Context): void {
  for (const part of ctx.parts()) {
    const contentType = ctx.contentType(part);
    if (contentType === undefined || NOTES_FAMILY.has(contentType)) continue;
    if (
      contentType !== CONTENT_TYPE.slide &&
      contentType !== CONTENT_TYPE.slideLayout &&
      contentType !== CONTENT_TYPE.slideMaster
    ) {
      continue;
    }
    const document = ctx.document(part);
    if (document === null) continue;

    for (const element of descendantElements(document.root)) {
      if (element.local !== 'ph' || namespaceOf(element) !== NS.p) continue;
      const type = attributeValue(element, 'type');
      if (type !== 'hdr' && type !== 'sldImg') continue;
      ctx.add(
        'V022',
        attributeLocation(part, element, 'type'),
        'type="' +
          type +
          '" is a whole-package refusal here, alone and with no other change. The other seven ' +
          'content types were built as one-type packages in the same bisection and every one ' +
          'opens: obj, chart, tbl, clipArt, dgm, media, pic. Nothing in the schema says so - ' +
          'CT_Placeholder is one complex type shared by every sheet family and ' +
          'ST_PlaceholderType is one enumeration holding all sixteen values - but hdr and ' +
          'sldImg belong to the notes and handout families and PowerPoint enforces it.',
      );
    }
  }
}

/**
 * The built-in geometry guides.
 *
 * Deliberately generous, and deliberately temporary. The authoritative table is
 * sub-phase 2.2's - forty-four seeded guides beside the seventeen `fmla`
 * operators - and it belongs in `@pptx-studio/geometry` rather than here. Until
 * it exists this list plus the two shape tests below stand in for it.
 *
 * Generous in the safe direction on purpose. This rule is fatal, so a name we
 * wrongly think is undefined refuses a file that works, while a name we wrongly
 * accept only means we miss one instance of a defect whose real-world form is a
 * typo or a deleted `a:gd` - and neither of those looks anything like `wd12`.
 */
const BUILTIN_GUIDES = new Set([
  '3cd4',
  '3cd8',
  '5cd8',
  '7cd8',
  'b',
  'cd2',
  'cd4',
  'cd8',
  'h',
  'hc',
  'hd2',
  'hd3',
  'hd4',
  'hd5',
  'hd6',
  'hd8',
  'l',
  'ls',
  'r',
  'ss',
  'ssd2',
  'ssd4',
  'ssd6',
  'ssd8',
  'ssd16',
  'ssd32',
  't',
  'vc',
  'w',
  'wd2',
  'wd3',
  'wd4',
  'wd5',
  'wd6',
  'wd8',
  'wd10',
  'wd32',
]);

/** `wd12`, `hd10`, `ssd12` - the same shape as a built-in, with a divisor we did not list. */
const BUILTIN_SHAPED = /^(?:w|h|ss|ls)d\d+$/;
/** `2cd4`, `5cd8` - a fraction of a full circle. */
const ANGLE_SHAPED = /^\d+cd\d+$/;

/** A literal number, in EMU or in 60000ths of a degree. Both may be negative. */
const LITERAL = /^-?\d+$/;

function isDefinedGuide(name: string, defined: ReadonlySet<string>): boolean {
  if (name === '' || LITERAL.test(name)) return true;
  if (defined.has(name) || BUILTIN_GUIDES.has(name)) return true;
  return BUILTIN_SHAPED.test(name) || ANGLE_SHAPED.test(name);
}

/** Attributes that hold a guide name rather than a value of their own. */
const GUIDE_VALUED: Readonly<Record<string, readonly string[]>> = {
  pt: ['x', 'y'],
  pos: ['x', 'y'],
  rect: ['l', 't', 'r', 'b'],
  ahXY: ['minX', 'maxX', 'minY', 'maxY', 'gdRefX', 'gdRefY'],
  ahPolar: ['minAng', 'maxAng', 'minR', 'maxR', 'gdRefAng', 'gdRefR'],
  cxn: ['ang'],
  arcTo: ['wR', 'hR', 'stAng', 'swAng'],
};

/**
 * Every geometry guide named is a guide that is defined.
 *
 * Scoped to one `a:custGeom` or `a:prstGeom` at a time, because that is the
 * scope a guide has: `adj1` in one shape has nothing to do with `adj1` in the
 * next. A `prstGeom`'s guides come from the preset definition rather than from
 * the markup, so only its `a:avLst` is checked - which is the whole of what a
 * preset shape can get wrong from here, and is exactly what dragging an adjust
 * handle writes.
 */
export function v023GeometryGuides(ctx: Context): void {
  forEachElement(ctx, (part, geometry) => {
    const isCustom = geometry.local === 'custGeom';
    if ((!isCustom && geometry.local !== 'prstGeom') || namespaceOf(geometry) !== NS.a) return;

    const defined = new Set<string>();
    for (const element of descendantElements(geometry)) {
      if (element.local !== 'gd') continue;
      const name = attributeValue(element, 'name');
      if (name !== undefined && name !== '') defined.add(name);
    }

    for (const element of descendantElements(geometry)) {
      if (element.local === 'gd') {
        // `fmla="*/ w adj1 100000"` - the operator, then two or three
        // arguments, each a guide name or a literal.
        const formula = attributeValue(element, 'fmla');
        if (formula === undefined) continue;
        const args = formula.trim().split(/\s+/).slice(1);
        for (const arg of args) {
          if (isDefinedGuide(arg, defined)) continue;
          report(ctx, part, element, 'fmla', arg, defined);
        }
        continue;
      }
      const names = GUIDE_VALUED[element.local];
      if (names === undefined || namespaceOf(element) !== NS.a) continue;
      for (const attributeName of names) {
        const value = attribute(element, attributeName);
        if (value === undefined || isDefinedGuide(value.value, defined)) continue;
        report(ctx, part, element, attributeName, value.value, defined);
      }
    }
  });
}

function report(
  ctx: Context,
  part: string,
  element: XElement,
  attributeName: string,
  name: string,
  defined: ReadonlySet<string>,
): void {
  ctx.add(
    'V023',
    attributeLocation(part, element, attributeName),
    '"' +
      name +
      '" names a guide nothing defines. This geometry defines ' +
      (defined.size === 0 ? 'none' : [...defined].sort().join(', ')) +
      '. ST_GeomGuideName is an unconstrained token, so nothing in the schema forbids it - and ' +
      'a package containing it is refused outright, not repaired and not drawn without the shape.',
  );
}

/** The chart namespace, for `V024`. */
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/**
 * A series' `c:tx` holds `c:strRef` or `c:v`, and nothing else.
 *
 * Scoped to `c:ser/c:tx` rather than to every `c:tx`, and the distinction is
 * real: a series' text is `CT_SerTx`, a choice of `c:strRef` or `c:v`, while a
 * title's or a data label's `c:tx` is `CT_Tx`, a choice of `c:strRef` or
 * `c:rich`. Checking both with one list would refuse every chart title in the
 * corpus.
 */
export function v024SeriesText(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'tx' || namespaceOf(element) !== CHART_NS) return;
    const parent = element.parent;
    if (parent === undefined || parent.local !== 'ser' || namespaceOf(parent) !== CHART_NS) return;

    for (const child of childElements(element)) {
      if (child.local === 'strRef' || child.local === 'v') continue;
      ctx.add(
        'V024',
        elementLocation(part, child),
        '<' +
          child.qname +
          '> inside a series <' +
          element.qname +
          '>. CT_SerTx is a choice of c:strRef or c:v and nothing else. The trap is that c:cat ' +
          'and c:val, two elements away in the same series, both accept the literal forms - so ' +
          'a c:strLit here looks like markup that ought to work, and it is a whole-package ' +
          'refusal.',
      );
    }
  });
}

/** No `p:control`. */
export function v025Control(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'control' || namespaceOf(element) !== NS.p) return;
    ctx.add(
      'V025',
      elementLocation(part, element),
      'a <p:control> is a whole-package refusal in all eight forms tried: bare, name-only, with ' +
        'and without r:id, with a p:pic preview, with an ActiveX part and its .bin, and inside a ' +
        'macro-enabled package. An empty <p:controls> is accepted, which places the refusal ' +
        'precisely on this element.',
    );
  });
}

/** The chart-style namespace and its thirty-one entries. */
const CHART_STYLE_NS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';

const CHART_STYLE_ENTRIES = [
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

/**
 * A `cs:chartStyle` carries all thirty-one of its entries.
 *
 * The shape of this finding is the interesting half. A chart with **no**
 * chart-style relationship at all opens; a chart whose style part is present
 * and has four of thirty-one does not. So the rule is not "the part is
 * required", it is "if the part exists it must be complete" - a subset is worse
 * than an absence, which is the opposite of what any partial-styling model
 * would predict, and is why we found it by shipping four entries rather than by
 * reasoning.
 */
export function v026ChartStyleComplete(ctx: Context): void {
  for (const part of ctx.parts()) {
    const document = ctx.document(part);
    if (document === null) continue;
    const root = document.root;
    if (root.local !== 'chartStyle' || namespaceOf(root) !== CHART_STYLE_NS) continue;

    const present = new Set(childElements(root).map((child) => child.local));
    const missing = CHART_STYLE_ENTRIES.filter((name) => !present.has(name));
    if (missing.length === 0) continue;
    ctx.add(
      'V026',
      elementLocation(part, root),
      'a chart style with ' +
        String(CHART_STYLE_ENTRIES.length - missing.length) +
        ' of its ' +
        String(CHART_STYLE_ENTRIES.length) +
        ' entries. Missing: ' +
        missing.join(', ') +
        '. Four entries were refused and thirty-one opened; a chart with no chart-style ' +
        'relationship at all is fine.',
    );
  }
}

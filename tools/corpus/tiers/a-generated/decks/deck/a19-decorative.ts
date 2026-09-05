import { placeholderXml, TITLE_BOX, type ProbeLayout, type ProbePart } from '../package.ts';
import { png } from '../png.ts';
import { grid, group, picture, scheme, shape, solidFill } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Accessibility: `@descr`, `@title`, `adec:decorative`, `@hidden`, and the
 * fact that `p:spTree` order **is** reading order.
 *
 * Sub-phase 12.3 has to announce a slide to a screen reader, and everything it
 * needs is on `p:cNvPr` - four attributes' worth of information that nothing
 * renders, that no test catches when it goes missing, and that a round trip
 * drops without a symptom. That is the whole case for a fixture.
 *
 * ## The extension GUID, measured rather than looked up
 *
 * `adec:decorative` is not in ECMA-376. It lives inside `p:cNvPr/a:extLst`
 * under an `a:ext/@uri` that appears in no public specification, so on
 * 2026-08-27 a rectangle was marked decorative in PowerPoint 16.0.20326 through
 * COM and the saved slide read back:
 *
 * ```xml
 * <p:cNvPr id="2" name="Rectangle 1">
 *   <a:extLst>
 *     <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
 *       <a16:creationId xmlns:a16="…/drawing/2014/main" id="{…}"/>
 *     </a:ext>
 *     <a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">
 *       <adec:decorative xmlns:adec="…/drawing/2017/decorative" val="1"/>
 *     </a:ext>
 *   </a:extLst>
 * </p:cNvPr>
 * ```
 *
 * Two things came out of that measurement beyond the GUID.
 *
 * **Decorative and alt text are mutually exclusive.** The same shape had
 * `AlternativeText` and `Title` set before `Decorative` was, and the saved
 * markup carries neither `@descr` nor `@title` - PowerPoint cleared them.
 * So "decorative" is not a flag beside the alt text, it replaces it, and a UI
 * that offers both at once is offering a state PowerPoint will not save.
 *
 * **`a:extLst` here holds several `a:ext` and order matters to nobody but us.**
 * ADR 0006 keeps `extLst` as an ordered opaque list keyed by `@uri` and never
 * rebuilds it, which is exactly what a list like this needs.
 *
 * ## `&quot;`, and how that question got answered
 *
 * The alt text on slide 1 contains a double quote. It is there because `@descr`
 * is the one attribute in a deck that holds free text a user typed, so it is
 * where an unescaped quote would first break a package - and because until this
 * deck the generator **threw** rather than escape one. The measurement above
 * settled it: renaming a shape to `a "quoted" name` made PowerPoint write
 * `name="a &quot;quoted&quot; name"`, two `&quot;` and no other entity in the
 * package. `escapeAttribute` now escapes, `powerpoint-conventions.json` records
 * the fourth entity, and this is the deck that exercises it.
 *
 * ## Reading order is document order, and this deck disagrees with itself
 *
 * A screen reader walks `p:spTree` in document order; PowerPoint's Selection
 * pane shows the same list, bottom-up. Slide 3 is laid out so that the visual
 * order - left to right, top to bottom - is the reverse of the document order,
 * so an implementation that sorts by position instead of by tree order reads
 * it backwards and a fixture that agreed with itself would never notice.
 */

const NS_ADEC = 'http://schemas.microsoft.com/office/drawing/2017/decorative';

/** The `a:ext/@uri` PowerPoint 16.0.20326 was measured writing for it. */
const DECORATIVE_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';

/**
 * `a:extLst` marking a shape decorative.
 *
 * `@val` is `ST_OnOff` and PowerPoint writes `1`. A shape marked decorative
 * gets `aria-hidden="true"` and `tabindex="-1"` in sub-phase 12.3 **and** this
 * marker on export, because NVDA is known to announce PowerPoint decoratives
 * anyway and one of the two is not enough.
 */
const DECORATIVE =
  `<a:extLst><a:ext uri="${DECORATIVE_URI}">` +
  `<adec:decorative xmlns:adec="${NS_ADEC}" val="1"/>` +
  '</a:ext></a:extLst>';

const cell = grid(3, 2);

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'titleOnly',
    name: 'Title Only',
    hasTitle: true,
    shapes: placeholderXml({
      id: 2,
      name: 'Title 1',
      type: 'title',
      // Alt text on a *layout* placeholder. A slide placeholder that inherits
      // this one inherits its geometry and not its description: `@descr` is
      // per-shape and does not cascade, which is why an empty layout title
      // leaves every slide built from it with no title alt text at all.
      descr: 'The title of the slide, from the layout',
      ...TITLE_BOX,
    }),
  },
  { type: 'blank', name: 'Blank' },
];

function labelled(spec: {
  readonly id: number;
  readonly index: number;
  readonly name: string;
  readonly descr?: string;
  readonly title?: string;
  readonly hidden?: boolean;
  readonly extLst?: string;
  readonly lines: readonly string[];
  readonly accent?: number;
}): string {
  return shape({
    id: spec.id,
    name: spec.name,
    ...(spec.descr === undefined ? {} : { descr: spec.descr }),
    ...(spec.title === undefined ? {} : { title: spec.title }),
    ...(spec.hidden === undefined ? {} : { hidden: spec.hidden }),
    ...(spec.extLst === undefined ? {} : { extLst: spec.extLst }),
    ...cell(spec.index),
    fill: solidFill(scheme(`accent${String(spec.accent ?? 1)}`, '<a:alpha val="25000"/>')),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: spec.lines.map((line) => textLine(line, { sz: 1200 }, { algn: 'ctr' })).join(''),
    }),
  });
}

/** Six by four pixels of flat colour: enough to be a picture, and no more. */
const DOT_PNG = (): Uint8Array => png(6, 4, () => 0x2980b9);

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/media/image1.png',
    bytes: DOT_PNG(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
];

export const a19Decorative: ProbeDeck = {
  id: 'a19-decorative',
  title: 'PPTX Studio corpus: a19 decorative',
  description:
    'The accessibility surface, which is four attributes on p:cNvPr that nothing renders: @descr, ' +
    '@title, @hidden, and the adec:decorative extension under the a:ext uri measured from ' +
    'PowerPoint 16.0.20326 rather than looked up. Decorative and alt text are mutually exclusive - ' +
    'PowerPoint clears @descr and @title when a shape is marked decorative. One @descr carries a ' +
    'double quote, which is the attribute where &amp;quot; first matters. Slide 3 lays its shapes out ' +
    'so that visual order is the reverse of p:spTree order, because p:spTree order is what a screen ' +
    'reader announces.',
  features: {
    // 22 `p:sp`: two on the master, one on the Title Only layout, seven on each
    // of slides 1 and 3, and five on slide 2 - where the two group children
    // count and the two pictures do not, because a `p:pic` is not a `p:sp`.
    shape: 22,
    picture: 2,
    group: 1,
    placeholder: 7,
    presetGeom: 17,
    // No `blipFill`. A `p:pic` holds a *PresentationML* `p:blipFill`, and the
    // census's `blipFill` key is the DrawingML `a:blipFill` - a fill on a shape.
    // Two pictures contribute nothing to it, which is the whole reason the two
    // elements have different namespaces.
    gradientFill: 2,
    // A shape, a group and a picture. Not the placeholder: PowerPoint clears
    // `@descr` and `@title` when a shape is marked decorative, and a
    // placeholder that announces nothing is a different defect from a
    // decoration.
    decorative: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a19 decorative',
    layouts: LAYOUTS,
    parts: PARTS,
    slides: [
      {
        title: 'a19 — the four states of a p:cNvPr',
        layout: 0,
        body:
          labelled({
            id: 10,
            index: 0,
            name: 'Described',
            descr: 'A bar chart showing revenue rising through the year',
            lines: ['@descr only', 'the ordinary case'],
          }) +
          labelled({
            id: 11,
            index: 1,
            name: 'Titled',
            title: 'Revenue',
            lines: ['@title only', 'announced before @descr'],
            accent: 2,
          }) +
          labelled({
            id: 12,
            index: 2,
            name: 'Both',
            descr: 'The quarterly figures, described at length for a reader who cannot see them',
            title: 'Quarterly figures',
            lines: ['@title and @descr', 'both, on one shape'],
            accent: 3,
          }) +
          labelled({
            id: 13,
            index: 3,
            name: 'Quoted description',
            descr: 'A pull quote reading "the numbers speak for themselves", set in 24pt',
            lines: ['@descr with a double quote', 'written &quot; by PowerPoint'],
            accent: 4,
          }) +
          labelled({
            id: 14,
            index: 4,
            name: 'Decorative',
            extLst: DECORATIVE,
            lines: ['adec:decorative val="1"', 'and no @descr — PowerPoint clears it'],
            accent: 5,
          }) +
          labelled({
            id: 15,
            index: 5,
            name: 'Neither',
            lines: ['no @descr, no @title,', 'no marker — the default, and a defect'],
            accent: 6,
          }),
      },
      {
        title: 'a19 — decorative on a group, a picture and a placeholder',
        layout: 0,
        body:
          // A group marked decorative. `p:cNvPr` is the same element on a
          // `p:grpSp`, so the marker goes in the same place - and a reader that
          // only looks at `p:sp` misses it.
          group({
            id: 10,
            name: 'Decorative group',
            extLst: DECORATIVE,
            x: 457200,
            y: 1188720,
            cx: 3474720,
            cy: 2286000,
            children:
              shape({
                id: 11,
                name: 'Group child A',
                x: 457200,
                y: 1188720,
                cx: 1600200,
                cy: 2286000,
                fill: solidFill(scheme('accent1')),
                caption: false,
              }) +
              shape({
                id: 12,
                name: 'Group child B',
                x: 2331720,
                y: 1188720,
                cx: 1600200,
                cy: 2286000,
                fill: solidFill(scheme('accent2')),
                caption: false,
              }),
          }) +
          picture({
            id: 13,
            name: 'Described picture',
            relId: 'rId2',
            x: 4267200,
            y: 1188720,
            cx: 3474720,
            cy: 2286000,
            description: 'A photograph of the building, taken from the north side',
          }) +
          // The same picture again, decorative instead of described. `p:pic`
          // has its own `p:cNvPr`, and this is where a decorative image most
          // often belongs: a divider, a texture, a logo already in the title.
          picture({
            id: 14,
            name: 'Decorative picture',
            relId: 'rId2',
            x: 8077200,
            y: 1188720,
            cx: 3474720,
            cy: 2286000,
            description: '',
            extLst: DECORATIVE,
          }) +
          placeholderXml({
            id: 15,
            name: 'Described placeholder',
            type: 'body',
            idx: 1,
            descr: 'The agenda, as a bulleted list',
            title: 'Agenda',
            x: 457200,
            y: 3703320,
            cx: 11277600,
            cy: 1600200,
            body: textLine('A placeholder carries @descr and @title like any other shape.'),
          }) +
          // Hidden is not decorative. The shape is in the tree, has alt text,
          // and is not drawn - three facts that have to survive independently.
          labelled({
            id: 16,
            index: 5,
            name: 'Hidden but described',
            descr: 'A note to the presenter, hidden on the slide',
            hidden: true,
            lines: ['@hidden="1"', 'in the tree, not drawn'],
            accent: 6,
          }),
        rels: [
          {
            id: 'rId2',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            target: '../media/image1.png',
          },
        ],
      },
      {
        title: 'a19 — reading order is p:spTree order, not position',
        layout: 0,
        body: [5, 4, 3, 2, 1, 0]
          .map((position, order) =>
            labelled({
              id: 10 + order,
              index: position,
              name: 'Reading order ' + String(order + 1),
              descr: `Shape ${String(order + 1)} in document order, at grid position ${String(position + 1)}`,
              lines: [
                `${String(order + 1)} in p:spTree`,
                `position ${String(position + 1)} on screen`,
              ],
              accent: (order % 6) + 1,
            }),
          )
          .join(''),
      },
    ],
  }),
};

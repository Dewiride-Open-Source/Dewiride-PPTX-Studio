import { placeholderXml, TITLE_BOX, type ProbeLayout } from '../package.ts';
import { textLine } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Placeholders: every type, and every branch of the five-tier matcher.
 *
 * This is the first deck the chassis grows layouts for, and the reason it is
 * `a02` rather than something later is that inheritance is what the product is
 * for. Everything in phase 7 - Change Layout, Reset, layout-compatibility
 * scoring - is the matcher plus a plan built on top of it, and a matcher with
 * no adversarial fixture is a function nobody can refactor.
 *
 * ## Where the geometry deliberately is not
 *
 * Almost every `p:ph` on a slide here has **no `a:xfrm`**. That absence is the
 * probe. A placeholder without explicit geometry takes its layout counterpart's
 * and follows it when the layout changes; one with an `a:xfrm` is pinned where
 * it is forever. Flattening the first into the second at parse time - which is
 * what every browser PowerPoint importer does - is precisely why their slides
 * cannot switch layout, so a fixture where the two are visibly different
 * matters more than any number of shapes.
 *
 * The one exception is slide 3's orphan, which has nothing to inherit from and
 * would render at zero size without one.
 *
 * ## The five tiers, one slide, one shape each
 *
 * Slide 3 is the matcher's test matrix in shape form. Reading it against
 * layout 2:
 *
 * | shape | on the slide            | in the layout      | tier                       |
 * | ----- | ----------------------- | ------------------ | -------------------------- |
 * | 10    | `ctrTitle` idx 9        | `title`, no idx    | 2 - any title, ignore idx  |
 * | 11    | `body` idx 1            | `body` idx 1       | 1 - exact `(type, idx)`    |
 * | 12    | `sldNum` idx 99         | `sldNum` idx 12    | 3 - raw type, ignore idx   |
 * | 13    | `body` idx 5            | one `body` only    | 4 - the sole body          |
 * | 14    | `pic` idx 42            | no `pic` at all    | 5 - orphan                 |
 *
 * Tier 2 is the one worth stating: `ctrTitle` normalizes to `title` before any
 * comparison, so a centred title on the slide binds to a plain title in the
 * layout. Miss the normalization and the title becomes an orphan on every deck
 * built from the Title Slide layout, which is most of them.
 *
 * ## Fourteen of the sixteen types, and why the other two are not here
 *
 * `hdr` and `sldImg` began with their own layout and their own slide, so that
 * if the deck stopped opening one bisection step would say which of the sixteen
 * types did it. It stopped opening, and it did.
 *
 * **Either one, alone, on a slide layout or on a slide, is a whole-package
 * refusal**: "the file or directory is corrupted and unreadable", no part name,
 * no element name. The other seven content types were built the same way in the
 * same bisection and every one of them opens - `obj`, `chart`, `tbl`,
 * `clipArt`, `dgm`, `media`, `pic`.
 *
 * Nothing in the schema says this. `CT_Placeholder` is one complex type used by
 * masters, layouts, slides, notes slides and handout masters alike, and
 * `ST_PlaceholderType` is one enumeration with all sixteen values in it. The
 * restriction is real and unwritten: a slide has no header, and the slide image
 * is what a notes page shows, so those two belong to the notes and handout
 * families only. `a14-notes` is where they will be exercised, on the sheets
 * that may legally hold them.
 *
 * So this deck covers **fourteen**, and the pair become fixtures under
 * `corpus/reject/` once that exists, with `C-REJECT` asserting the validator
 * catches them before an export is handed to anybody.
 */

const HF_ROW_Y = 6356350;
const HF_ROW_HEIGHT = 365125;

/** The date / footer / slide-number trio, at the geometry PowerPoint uses. */
function headerFooterTrio(firstId: number, firstIdx: number): string {
  return (
    placeholderXml({
      id: firstId,
      name: 'Date Placeholder',
      type: 'dt',
      idx: firstIdx,
      x: 457200,
      y: HF_ROW_Y,
      cx: 2743200,
      cy: HF_ROW_HEIGHT,
    }) +
    placeholderXml({
      id: firstId + 1,
      name: 'Footer Placeholder',
      type: 'ftr',
      idx: firstIdx + 1,
      x: 4038600,
      y: HF_ROW_Y,
      cx: 4114800,
      cy: HF_ROW_HEIGHT,
    }) +
    placeholderXml({
      id: firstId + 2,
      name: 'Slide Number Placeholder',
      type: 'sldNum',
      idx: firstIdx + 2,
      x: 8610600,
      y: HF_ROW_Y,
      cx: 2743200,
      cy: HF_ROW_HEIGHT,
    })
  );
}

/** The eight content types, laid out four across and two down under the title. */
const CONTENT_TYPES = ['body', 'obj', 'chart', 'tbl', 'clipArt', 'dgm', 'media', 'pic'] as const;

function contentGrid(): string {
  const width = 2743200;
  const height = 2286000;
  return CONTENT_TYPES.map((type, index) =>
    placeholderXml({
      id: 10 + index,
      name: 'Content Placeholder ' + String(index + 1),
      type,
      idx: index + 1,
      x: 457200 + (index % 4) * (width + 152400),
      y: 1188720 + Math.floor(index / 4) * (height + 152400),
      cx: width,
      cy: height,
    }),
  ).join('');
}

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'title',
    name: 'Title Slide',
    shapes:
      placeholderXml({
        id: 2,
        name: 'Title 1',
        type: 'ctrTitle',
        x: 1524000,
        y: 1857375,
        cx: 9144000,
        cy: 2387600,
      }) +
      placeholderXml({
        id: 3,
        name: 'Subtitle 2',
        type: 'subTitle',
        idx: 1,
        x: 1524000,
        y: 4344988,
        cx: 9144000,
        cy: 1655762,
      }),
  },
  {
    type: 'cust',
    name: 'Content Types',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }) + contentGrid(),
  },
  {
    type: 'cust',
    name: 'Matcher Cases',
    shapes:
      placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }) +
      placeholderXml({
        id: 3,
        name: 'Content Placeholder 2',
        type: 'body',
        idx: 1,
        x: 457200,
        y: 1188720,
        cx: 11277600,
        cy: 4114800,
      }) +
      headerFooterTrio(4, 10),
    // `p:hf` on a layout overrides the master's for slides bound here. Turning
    // the date off while the master has it on is the only way to see which of
    // the two a renderer actually read.
    hf: '<p:hf dt="0"/>',
  },
  // Bound to by nothing, and `preserve="1"`. An unused layout is normal - it is
  // what "change layout" changes *to* - and media GC in 1.3 must not collect it.
  { type: 'blank', name: 'Blank' },
];

export const a02Placeholders: ProbeDeck = {
  id: 'a02-placeholders',
  title: 'PPTX Studio corpus: a02 placeholders',
  description:
    'Fourteen of the sixteen ST_PlaceholderType values across four layouts and a master, p:hf on ' +
    'both the master and a layout, and one shape per branch of the five-tier placeholder matcher. ' +
    'Almost no slide placeholder carries an a:xfrm, because that absence is what makes Change ' +
    'Layout work. hdr and sldImg are absent by measurement, not oversight: either one alone, on a ' +
    'layout or a slide, is a whole-package refusal.',
  features: {
    shape: 37,
    placeholder: 37,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a02 placeholders',
    layouts: LAYOUTS,
    masterShapes: headerFooterTrio(4, 2),
    masterHf: '<p:hf sldNum="1" hdr="0" ftr="1" dt="1"/>',
    slides: [
      {
        title: 'a02 — title and subtitle',
        layout: 0,
        body:
          placeholderXml({
            id: 10,
            name: 'Title 1',
            type: 'ctrTitle',
            body: textLine('a02 — centred title'),
          }) +
          placeholderXml({
            id: 11,
            name: 'Subtitle 2',
            type: 'subTitle',
            idx: 1,
            body: textLine('Subtitle, inheriting its box from the layout'),
          }),
      },
      {
        title: 'a02 — the eight content types',
        layout: 1,
        body: CONTENT_TYPES.map((type, index) =>
          placeholderXml({
            id: 10 + index,
            name: 'Content Placeholder ' + String(index + 1),
            type,
            idx: index + 1,
            body: textLine(`type="${type}" idx="${String(index + 1)}"`),
          }),
        ).join(''),
      },
      {
        title: 'a02 — matcher cases',
        layout: 2,
        body:
          // Tier 2: ctrTitle normalizes to title, and a title matches a title
          // whatever the idx says.
          placeholderXml({
            id: 10,
            name: 'Title 1',
            type: 'ctrTitle',
            idx: 9,
            body: textLine('a02 — matcher cases'),
          }) +
          // Tier 1: exact (type, idx).
          placeholderXml({
            id: 11,
            name: 'Content Placeholder 2',
            type: 'body',
            idx: 1,
            body: textLine('tier 1 — body idx 1 matches body idx 1 exactly'),
          }) +
          // Tier 3: sldNum/dt/ftr/hdr match on raw type, idx ignored.
          placeholderXml({
            id: 12,
            name: 'Slide Number Placeholder 3',
            type: 'sldNum',
            idx: 99,
            body: textLine('3'),
          }) +
          // Tier 4: the layout has exactly one body, so an unmatched body takes it.
          placeholderXml({
            id: 13,
            name: 'Content Placeholder 4',
            type: 'body',
            idx: 5,
            body: textLine('tier 4 — body idx 5 falls back to the only body'),
          }) +
          // Tier 5: orphan. Nothing to inherit, so it carries its own geometry
          // or it renders at zero size.
          placeholderXml({
            id: 14,
            name: 'Picture Placeholder 5',
            type: 'pic',
            idx: 42,
            x: 457200,
            y: 5486400,
            cx: 5334000,
            cy: 685800,
            body: textLine('tier 5 — pic idx 42 is an orphan'),
          }),
      },
    ],
  }),
};

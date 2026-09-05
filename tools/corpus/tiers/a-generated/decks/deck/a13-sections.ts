import { grid, shape } from '../shapes.ts';
import { txBody, textLine } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Sections and custom shows: two ways of grouping slides, neither of which is
 * in ECMA-376 the same way.
 *
 * ## `p14:sectionLst` is an extension, and it is not optional in practice
 *
 * Sections are the feature every PowerPoint user knows and no version of the
 * standard has. They live in `p:presentation/p:extLst` under
 * `{521415D9-36F7-43E2-AB2F-B90AF26B5E84}`, in the PowerPoint 2010 namespace,
 * and a reader that treats `extLst` as opaque - which is exactly what ADR 0006
 * says to do with it - keeps them across a round trip without understanding
 * them. That is the right default and it is also why this deck exists: the
 * moment anything reorders or deletes slides, the section list has to be
 * updated in step, and `p14:sldId/@id` is a **value copy** of `p:sldId/@id`
 * rather than a relationship, so nothing dangles when it goes stale. It just
 * quietly stops matching.
 *
 * ## `p:custShowLst` is in the standard, and points by relationship id
 *
 * A custom show is a named ordering of slides, `p:sld/@r:id` pointing into
 * **`ppt/_rels/presentation.xml.rels`**. Three things follow that a naive
 * implementation gets wrong:
 *
 * - A slide may appear in a show **more than once**. "With a repeat" below
 *   lists slide 1 twice, which is legal and is why a show is a list rather
 *   than a set.
 * - A slide may be in several shows at once, or in none.
 * - Deleting a slide has to purge its `p:custShow` entries as well as its
 *   `p:sldId`, its relationship, its content-type override and its part.
 *   Sub-phase 5.6 names that last one explicitly because almost every
 *   home-grown implementation forgets it, and the symptom is a dangling
 *   relationship - which is fatal where an orphan is harmless.
 *
 * ## Order inside `p:presentation`
 *
 * `CT_Presentation` is a sequence, and these two are not adjacent:
 * `… sldSz, notesSz, smartTags, embeddedFontLst, custShowLst, photoAlbum,
 * custDataLst, kinsoku, defaultTextStyle, modifyVerifier, extLst`. So the
 * custom shows come ninth and the sections come last, inside `extLst`, with
 * six optional elements' worth of space between them.
 */

const SECTION_GUIDS = [
  '{2E9C4A6D-3B71-4F08-9C25-6A1D7E0B4F31}',
  '{7F1B85C0-9D46-4A23-B7E8-1C5A0F62D948}',
  '{B4D0E37A-58C2-4691-8FA3-0E7B926C15D2}',
] as const;

/** The section extension's own uri, which is what makes it findable. */
const SECTION_EXT_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

interface Section {
  readonly name: string;
  /** One-based slide numbers, which become `p:sldId/@id` of 255 + n. */
  readonly slides: readonly number[];
}

const SECTIONS: readonly Section[] = [
  { name: 'Opening', slides: [1] },
  { name: 'Body', slides: [2, 3] },
  { name: 'Closing', slides: [4] },
];

interface CustomShow {
  readonly name: string;
  readonly slides: readonly number[];
}

const SHOWS: readonly CustomShow[] = [
  { name: 'Short', slides: [1, 4] },
  { name: 'Body only', slides: [2, 3] },
  // Slide 1 twice. A show is an ordering, not a selection.
  { name: 'With a repeat', slides: [1, 2, 1] },
];

/** `p:sldId/@id` for the nth slide, matching what the chassis writes. */
const slideId = (n: number): string => String(255 + n);

/**
 * `p:sld/@r:id` for the nth slide.
 *
 * `rId1` is the slide master, so slide n is `rId(n+1)`. This deck reaching
 * into the chassis's relationship numbering is the price of custom shows
 * addressing slides by relationship rather than by id, and it is exactly the
 * coupling that makes "delete a slide" a five-part operation.
 */
const slideRid = (n: number): string => 'rId' + String(n + 1);

const customShowLst =
  '<p:custShowLst>' +
  SHOWS.map(
    (show, index) =>
      `<p:custShow name="${show.name}" id="${String(index)}"><p:sldLst>` +
      show.slides.map((n) => `<p:sld r:id="${slideRid(n)}"/>`).join('') +
      '</p:sldLst></p:custShow>',
  ).join('') +
  '</p:custShowLst>';

const sectionLst =
  `<p:extLst><p:ext uri="${SECTION_EXT_URI}">` +
  `<p14:sectionLst xmlns:p14="${NS_P14}">` +
  SECTIONS.map(
    (section, index) =>
      `<p14:section name="${section.name}" id="${SECTION_GUIDS[index] ?? ''}"><p14:sldIdLst>` +
      section.slides.map((n) => `<p14:sldId id="${slideId(n)}"/>`).join('') +
      '</p14:sldIdLst></p14:section>',
  ).join('') +
  '</p14:sectionLst></p:ext></p:extLst>';

/** What this slide belongs to, written on the slide so a human can check. */
function membership(id: number, number: number): string {
  const section = SECTIONS.find((entry) => entry.slides.includes(number));
  const shows = SHOWS.filter((show) => show.slides.includes(number));
  const cell = grid(1, 3)(0);
  return shape({
    id,
    name: 'Membership ' + String(number),
    ...cell,
    textBody: txBody({
      paras:
        textLine(`section: ${section?.name ?? 'none'}`, { sz: 1600 }) +
        textLine(
          'custom shows: ' + (shows.length === 0 ? 'none' : shows.map((s) => s.name).join(', ')),
          { sz: 1600 },
        ) +
        textLine(`p:sldId id="${slideId(number)}"  ·  r:id="${slideRid(number)}"`, { sz: 1600 }),
    }),
  });
}

export const a13Sections: ProbeDeck = {
  id: 'a13-sections',
  title: 'PPTX Studio corpus: a13 sections',
  description:
    'Four slides in three p14:sectionLst sections and three p:custShowLst custom shows, with one ' +
    'slide in two shows and one show listing the same slide twice. Sections address slides by a ' +
    'value copy of p:sldId/@id inside p:presentation/p:extLst; custom shows address them by ' +
    'relationship id, nine elements earlier in the same sequence. Neither edge dangles when it ' +
    'goes stale, which is what makes both worth a fixture.',
  features: {
    shape: 11,
    placeholder: 7,
    presetGeom: 4,
    gradientFill: 2,
    section: 3,
    customShow: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a13 sections',
    presentationTail: customShowLst + sectionLst,
    slides: [1, 2, 3, 4].map((number) => ({
      title: `a13 — slide ${String(number)} of four`,
      layout: 0,
      body: membership(10, number),
    })),
  }),
};

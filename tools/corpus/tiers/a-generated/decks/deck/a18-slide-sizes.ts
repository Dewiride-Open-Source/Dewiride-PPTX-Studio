import { scheme, shape, solidFill } from '../../markup/shapes.ts';
import { textLine, txBody } from '../../markup/text.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * A deck that is not 16:9, and a notes page that is not the slide.
 *
 * ## One `p:sldSz` per package, which is why this is one deck and not four
 *
 * The roster asked for 4:3, 16:9, A4 and a custom size. `p:sldSz` is a single
 * element on `p:presentation`, so a package has exactly one slide size and
 * those four need four decks. This one is `screen4x3` at PowerPoint's own
 * 9144000 x 6858000, every other Tier A deck is `screen16x9` at
 * 12192000 x 6858000, and A4 and custom are declared gaps rather than
 * quietly-absent ones - see `ROSTER.md`.
 *
 * 4:3 is the one worth having. Sub-phase 7.11's *Reflow on slide-size change*
 * is specifically 16:9 to 4:3 and back, with PowerPoint's Maximize and
 * Ensure-Fit modes and nine-slice anchoring so a corner logo stays in its
 * corner; a corpus with no 4:3 deck cannot test either direction.
 *
 * ## `@type` is a hint and `@cx`/`@cy` are the fact
 *
 * `ST_SlideSizeType` has sixteen values - `35mm`, `A3`, `A4`, `B4ISO`, `B4JIS`,
 * `B5ISO`, `B5JIS`, `banner`, `custom`, `hagakiCard`, `ledger`, `letter`,
 * `overhead`, `screen16x10`, `screen16x9`, `screen4x3` - and it says what the
 * deck was *meant* for, not how big it is. The two can disagree, `@type` is
 * optional, and every layout decision has to come from `@cx` and `@cy`. A
 * renderer that switches on `@type` is wrong on the first deck someone
 * resized by hand.
 *
 * ## `p:notesSz` is independent, and is required where `p:sldSz` is not
 *
 * The one genuinely counter-intuitive fact here. `p:sldSz` is `[0..1]` in
 * `CT_Presentation`; `p:notesSz` is `[1..1]`. So a deck may legally omit its
 * slide size and may not omit its notes size, which is the reverse of what
 * anyone expects, and it is in the must-not-break rules for that reason.
 *
 * Nor is one derived from the other. The chassis default is 6858000 x 9144000 -
 * portrait, and portrait under a landscape deck is the *normal* case. This deck
 * makes them disagree three ways at once: the slide is 4:3 landscape, the notes
 * page is A4 landscape at 10692000 x 7560000 - 297mm by 210mm, at 36000 EMU to
 * the millimetre - and the notes page is **wider than the slide**. Anything
 * that scales a notes page by the slide's aspect ratio is visibly wrong here.
 */

const WIDTH = 9144000;
const HEIGHT = 6858000;

/** 297mm x 210mm at 36000 EMU per millimetre. Landscape, unlike the default. */
const NOTES_WIDTH = 10692000;
const NOTES_HEIGHT = 7560000;

const MARKER = 914400;

/** A square in one corner, so a wrong slide extent puts it off the page. */
function corner(id: number, name: string, x: number, y: number): string {
  return shape({
    id,
    name,
    x,
    y,
    cx: MARKER,
    cy: MARKER,
    fill: solidFill(scheme('accent1')),
  });
}

function caption(id: number, name: string, y: number, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: MARKER + 91440,
    y,
    cx: WIDTH - 2 * (MARKER + 91440),
    cy: 1600200,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines.map((line) => textLine(line, { sz: 1600 }, { algn: 'ctr' })).join(''),
    }),
  });
}

export const a18SlideSizes: ProbeDeck = {
  id: 'a18-slide-sizes',
  title: 'PPTX Studio corpus: a18 slide sizes',
  description:
    'The corpus&apos;s only deck that is not 16:9: p:sldSz type="screen4x3" at 9144000 x 6858000, with ' +
    'corner markers a wrong extent pushes off the page and a shape that deliberately overhangs the ' +
    'right edge, which is legal. Its p:notesSz is A4 landscape at 10692000 x 7560000 - wider than ' +
    'the slide and a different aspect ratio, because the two are independent and p:notesSz is the ' +
    'one of the pair that is required. One package holds one slide size, so A4 and custom are ' +
    'declared gaps in ROSTER.md rather than absent ones.',
  features: {
    shape: 13,
    placeholder: 2,
    presetGeom: 11,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a18 slide sizes',
    slideWidth: WIDTH,
    slideHeight: HEIGHT,
    slideSizeType: 'screen4x3',
    notesWidth: NOTES_WIDTH,
    notesHeight: NOTES_HEIGHT,
    // Both layouts are empty: the chassis writes a slide's title into a box
    // 11277600 EMU wide, which is wider than this whole slide. Rather than
    // pin a title with an explicit a:xfrm on every slide - the one thing
    // `a02-placeholders` is about not doing - this deck writes its own text.
    layouts: [
      { type: 'blank', name: 'Blank 4:3' },
      { type: 'cust', name: 'Full Bleed 4:3' },
    ],
    slides: [
      {
        title: 'a18 — four corners of a 4:3 slide',
        layout: 0,
        body:
          corner(10, 'Top left', 0, 0) +
          corner(11, 'Top right', WIDTH - MARKER, 0) +
          corner(12, 'Bottom left', 0, HEIGHT - MARKER) +
          corner(13, 'Bottom right', WIDTH - MARKER, HEIGHT - MARKER) +
          caption(14, 'Size caption', 2628900, [
            'p:sldSz cx="9144000" cy="6858000" type="screen4x3"',
            'Four markers, each flush with a corner.',
            'A renderer holding 12192000 puts two of them off the page.',
          ]),
      },
      {
        title: 'a18 — content outside the extents',
        layout: 1,
        body:
          shape({
            id: 10,
            name: 'Overhanging band',
            x: WIDTH - 2 * MARKER,
            y: 1828800,
            cx: 3 * MARKER,
            cy: MARKER,
            fill: solidFill(scheme('accent2')),
          }) +
          shape({
            id: 11,
            name: 'Negative offset band',
            x: -MARKER,
            y: 3200400,
            cx: 3 * MARKER,
            cy: MARKER,
            fill: solidFill(scheme('accent3')),
          }) +
          caption(12, 'Overhang caption', 4572000, [
            'Objects within a slide can be specified outside these extents.',
            'One starts at a negative x; one runs past the right edge.',
            'Both are legal and both must clip rather than move.',
          ]),
      },
      {
        title: 'a18 — the notes page is a different shape',
        layout: 0,
        body:
          shape({
            id: 10,
            name: 'Notes page footprint, to scale',
            x: 457200,
            y: 457200,
            // The notes page at the same scale as the slide would be wider than
            // the slide, so this is drawn at a third and labelled.
            cx: Math.round(NOTES_WIDTH / 3),
            cy: Math.round(NOTES_HEIGHT / 3),
            fill: solidFill(scheme('accent4')),
          }) +
          shape({
            id: 11,
            name: 'Chassis default notes page, to scale',
            x: 457200 + Math.round(NOTES_WIDTH / 3) + 182880,
            y: 457200,
            cx: Math.round(6858000 / 3),
            cy: Math.round(9144000 / 3),
            fill: solidFill(scheme('accent5')),
          }) +
          caption(12, 'Notes caption', 4800600, [
            'p:notesSz cx="10692000" cy="7560000" — A4 landscape, wider than the slide.',
            'The other rectangle is the portrait default every other deck uses.',
            'p:notesSz is required; p:sldSz is optional.',
          ]),
      },
    ],
  }),
};

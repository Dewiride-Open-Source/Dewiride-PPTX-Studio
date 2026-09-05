import { scheme, shape, solidFill } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * `p:sldSz type="A4"`, and the proof that `@type` names nothing measurable.
 *
 * ## Why this is a package of its own
 *
 * `p:sldSz` is one element on `p:presentation`, so a package has exactly one
 * slide size. The roster's "4:3, 16:9, A4 and a custom size" is four decks:
 * `a18-slide-sizes` is 4:3, every other Tier A deck is 16:9 at the modern
 * default extent, this is A4, and `a42-custom-size` is the one with no `@type`
 * at all.
 *
 * ## The extent, measured
 *
 * 9906000 x 6858000 EMU - 780 x 540 points, 10.833 x 7.5 inches. Read out of
 * `PageSetup.SlideWidth`/`SlideHeight` after setting `PageSetup.SlideSize` to
 * `ppSlideSizeA4Paper` in PowerPoint 16.0.20326 on 2026-08-28, not derived from
 * the paper size. It is **not** A4: A4 is 297 x 210 mm, which is
 * 10692000 x 7560000 EMU, and this is 275.17 x 190.5 mm. PowerPoint's "A4
 * Paper (210x297 mm)" has always been the printable area rather than the sheet,
 * and a renderer that computes the extent from the paper name is wrong by 8mm
 * in one axis and 20mm in the other.
 *
 * Same measurement, same run, and the reason this deck exists rather than
 * `a18` growing a fourth slide:
 *
 * | `ST_SlideSizeType` | `@cx`      | `@cy`     |
 * | ------------------ | ---------- | --------- |
 * | `onScreen`         | 9144000    | 6858000   |
 * | `letter`           | 9144000    | 6858000   |
 * | `overhead`         | 9144000    | 6858000   |
 * | `A4`               | **9906000**| 6858000   |
 * | `35mm`             | 10287000   | 6858000   |
 * | `banner`           | 7315200    | 914400    |
 * | `screen16x9`       | 9144000    | 5143500   |
 * | `screen16x9`       | 12192000   | 6858000   |
 *
 * **Three of the sixteen enumeration values write the same extent** - a deck
 * typed `letter` and a deck typed `overhead` are byte-identical in `p:sldSz`
 * except for the hint - and **one value names two different extents**, because
 * Page Setup's "On-screen Show (16:9)" is 10 x 5.625 inches while the size a
 * new PowerPoint deck is created at is 13.333 x 7.5, and both are written
 * `type="screen16x9"`. So `@type` is not a size, is not a key, and cannot be
 * inverted. `@cx` and `@cy` are the fact, `@type` is a note about what the deck
 * was meant for, and `@type` is optional while a renderer's need for a size is
 * not.
 *
 * Slide 2 draws the first of those two facts and slide 3 the second, at one
 * scale, so a renderer that switched on `@type` produces a visibly wrong
 * picture rather than a subtly wrong number.
 *
 * ## What this deck does not carry
 *
 * No title placeholder on either layout. The chassis writes a slide's title
 * into a box 11277600 EMU wide, which is wider than this whole slide; rather
 * than pin an explicit `a:xfrm` on every title - the one thing
 * `a02-placeholders` exists to avoid - each slide writes its own text, exactly
 * as `a18` does.
 */

/** `ppSlideSizeA4Paper`, measured. 780 x 540 pt. */
const WIDTH = 9906000;
const HEIGHT = 6858000;

const MARKER = 685800;
const GUTTER = 91440;

/** A square flush with one corner, so a wrong extent puts it off the page. */
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

/** One of the extents in the table above, drawn to scale and labelled. */
function extent(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  accent: string,
  label: string,
): string {
  const divisor = 3;
  return shape({
    id,
    name,
    x,
    y,
    cx: Math.round(cx / divisor),
    cy: Math.round(cy / divisor),
    fill: solidFill(scheme(accent)),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: textLine(label, { sz: 1100, b: true }, { algn: 'ctr' }),
    }),
  });
}

function caption(id: number, name: string, y: number, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: MARKER + GUTTER,
    y,
    cx: WIDTH - 2 * (MARKER + GUTTER),
    cy: 1371600,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines.map((line) => textLine(line, { sz: 1400 }, { algn: 'ctr' })).join(''),
    }),
  });
}

export const a41A4: ProbeDeck = {
  id: 'a41-a4',
  title: 'PPTX Studio corpus: a41 A4',
  description:
    'p:sldSz type="A4" at 9906000 x 6858000 EMU, the extent measured out of PowerPoint 16.0.20326 ' +
    'rather than computed from the paper size - it is the printable area, 275.17 x 190.5 mm, and ' +
    'not the 297 x 210 mm sheet. A package holds one slide size, so this is a deck of its own. It ' +
    'also draws the two facts that make @type unusable as a key: onScreen, letter and overhead all ' +
    'write 9144000 x 6858000, and screen16x9 names both 9144000 x 5143500 and 12192000 x 6858000.',
  features: {
    shape: 14,
    placeholder: 2,
    presetGeom: 12,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a41 A4',
    slideWidth: WIDTH,
    slideHeight: HEIGHT,
    slideSizeType: 'A4',
    layouts: [
      { type: 'blank', name: 'Blank A4' },
      { type: 'cust', name: 'Full Bleed A4' },
    ],
    slides: [
      {
        title: 'a41 — the four corners of an A4 slide',
        layout: 0,
        body:
          corner(10, 'Top left', 0, 0) +
          corner(11, 'Top right', WIDTH - MARKER, 0) +
          corner(12, 'Bottom left', 0, HEIGHT - MARKER) +
          corner(13, 'Bottom right', WIDTH - MARKER, HEIGHT - MARKER) +
          caption(14, 'Extent caption', 2743200, [
            'p:sldSz cx="9906000" cy="6858000" type="A4"',
            '780 x 540 pt. Measured from PageSetup, not from the paper size:',
            'true A4 is 10692000 x 7560000, which this is not.',
          ]),
      },
      {
        title: 'a41 — three type values, one extent',
        layout: 1,
        body:
          extent(10, 'onScreen', 457200, 1143000, 9144000, 6858000, 'accent2', 'onScreen') +
          extent(
            11,
            'letter',
            457200 + Math.round(9144000 / 3) + GUTTER,
            1143000,
            9144000,
            6858000,
            'accent3',
            'letter',
          ) +
          extent(
            12,
            'overhead',
            457200 + 2 * (Math.round(9144000 / 3) + GUTTER),
            1143000,
            9144000,
            6858000,
            'accent4',
            'overhead',
          ) +
          caption(13, 'Collision caption', 4114800, [
            'Three ST_SlideSizeType values. One extent: 9144000 x 6858000.',
            'Nothing distinguishes these three packages but the hint,',
            'so @type cannot be inverted to a size.',
          ]),
      },
      {
        title: 'a41 — one type value, two extents',
        layout: 1,
        body:
          extent(
            10,
            'screen16x9 as Page Setup writes it',
            457200,
            1143000,
            9144000,
            5143500,
            'accent5',
            '9144000 x 5143500',
          ) +
          extent(
            11,
            'screen16x9 as a new deck is created',
            457200 + Math.round(9144000 / 3) + GUTTER,
            1143000,
            12192000,
            6858000,
            'accent6',
            '12192000 x 6858000',
          ) +
          caption(12, 'Ambiguity caption', 4114800, [
            'Both of these are written type="screen16x9".',
            '10 x 5.625 in from Page Setup; 13.333 x 7.5 in for a new deck.',
            'Every other deck in this corpus is the second one.',
          ]),
      },
    ],
  }),
};

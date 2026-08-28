import { scheme, shape, solidFill } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * A `p:sldSz` with no `@type`, at an extent no enumeration names - and it is
 * portrait.
 *
 * ## `custom` is a value nothing writes
 *
 * `ST_SlideSizeType` has a `custom` member, and PowerPoint does not use it.
 * Measured on 2026-08-28 against 16.0.20326: setting `PageSetup.SlideWidth` to
 * 1000pt and `SlideHeight` to 400pt makes `PageSetup.SlideSize` read back as
 * `ppSlideSizeCustom` (7), and the saved package contains
 *
 * ```
 * <p:sldSz cx="12700000" cy="5080000"/>
 * ```
 *
 * with no `@type` attribute at all. The enumeration member exists in the
 * schema and the COM property reports it, and the file says nothing. So the
 * absent attribute is not an oddity to be defended against - it is the normal
 * serialization of every deck anybody has ever resized, and `@type` being
 * `[0..1]` in `CT_SlideSize` is the schema agreeing.
 *
 * This deck therefore omits it, which is also what the chassis does by default
 * for any extent that is not the 16:9 one. `a41-a4` is the deck that writes a
 * `@type`.
 *
 * ## Why portrait, and why this particular extent
 *
 * 6858000 x 12192000 is the exact **transpose** of the default 16:9 extent
 * every other deck in the corpus uses, and no `ST_SlideSizeType` value names
 * it: all sixteen presets are landscape or square-ish, the tallest relative to
 * its width being `hagakiCard` at 4572000 x 2971800.
 *
 * The transpose is chosen deliberately. The failure it catches is the one that
 * hides: code that reads `@cx` into a height and `@cy` into a width renders
 * this deck as a perfectly ordinary 16:9 slide and looks *correct*. Nothing
 * about the aspect ratio, the theme or the fills gives it away. The corner
 * markers are what make it visible - swap the axes and all four land in the
 * wrong places, two of them off the page entirely.
 *
 * Portrait is also the case sub-phase 7.11 has to survive. *Reflow on
 * slide-size change* is specified against PowerPoint's Maximize and
 * Ensure-Fit modes with nine-slice anchoring, and every landscape-to-landscape
 * change leaves the sign of `cx - cy` alone. This one does not.
 *
 * ## And the notes page is landscape
 *
 * `p:notesSz` is independent of `p:sldSz` and required where `p:sldSz` is
 * optional - `[1..1]` against `[0..1]` in `CT_Presentation`, which is the
 * reverse of what anyone expects. `a18-slide-sizes` makes them disagree with a
 * landscape slide under a landscape notes page. This one inverts both: a
 * portrait slide under a **landscape** notes page, at 9144000 x 6858000. Every
 * other deck in the corpus is the other way round in both, so anything that
 * derived one from the other is wrong here in two directions at once.
 */

/** The transpose of the default 16:9 extent. Named by no ST_SlideSizeType value. */
const WIDTH = 6858000;
const HEIGHT = 12192000;

/** Landscape, under a portrait slide. */
const NOTES_WIDTH = 9144000;
const NOTES_HEIGHT = 6858000;

const MARKER = 685800;
const GUTTER = 91440;
const INSET = 342900;

/** A square flush with one corner. Transpose the axes and two go off the page. */
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

/** An extent drawn at a quarter scale and labelled. */
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
  return shape({
    id,
    name,
    x,
    y,
    cx: Math.round(cx / 4),
    cy: Math.round(cy / 4),
    fill: solidFill(scheme(accent)),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: textLine(label, { sz: 1000, b: true }, { algn: 'ctr' }),
    }),
  });
}

function caption(id: number, name: string, y: number, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: INSET,
    y,
    cx: WIDTH - 2 * INSET,
    cy: 1828800,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines.map((line) => textLine(line, { sz: 1200 }, { algn: 'ctr' })).join(''),
    }),
  });
}

export const a42CustomSize: ProbeDeck = {
  id: 'a42-custom-size',
  title: 'PPTX Studio corpus: a42 custom size',
  description:
    'A p:sldSz with no @type at all, which is what PowerPoint writes for every resized deck - ' +
    'measured: setting a custom size makes PageSetup.SlideSize report ppSlideSizeCustom while the ' +
    'saved file omits the attribute, so ST_SlideSizeType&apos;s custom member is a value nothing ' +
    'emits. The extent is 6858000 x 12192000, the exact transpose of the corpus default, which no ' +
    'enumeration value names and which renders as a plausible 16:9 slide if cx and cy are swapped. ' +
    'Its p:notesSz is landscape under a portrait slide, inverting both of a18&apos;s disagreements.',
  features: {
    shape: 13,
    placeholder: 2,
    presetGeom: 11,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a42 custom size',
    slideWidth: WIDTH,
    slideHeight: HEIGHT,
    // No `slideSizeType`: the chassis omits `@type` for any extent that is not
    // the default one, which is what PowerPoint does. Stated here because an
    // absent option is easy to read as an oversight.
    notesWidth: NOTES_WIDTH,
    notesHeight: NOTES_HEIGHT,
    layouts: [
      { type: 'blank', name: 'Blank portrait' },
      { type: 'cust', name: 'Full bleed portrait' },
    ],
    slides: [
      {
        title: 'a42 — the four corners of a portrait slide',
        layout: 0,
        body:
          corner(10, 'Top left', 0, 0) +
          corner(11, 'Top right', WIDTH - MARKER, 0) +
          corner(12, 'Bottom left', 0, HEIGHT - MARKER) +
          corner(13, 'Bottom right', WIDTH - MARKER, HEIGHT - MARKER) +
          caption(14, 'Extent caption', 5181600, [
            'p:sldSz cx="6858000" cy="12192000"',
            'No @type. This is what a resized deck looks like.',
            'Taller than it is wide, which no ST_SlideSizeType value is.',
          ]),
      },
      {
        title: 'a42 — the transpose renders as a plausible slide',
        layout: 1,
        body:
          extent(10, 'This slide', INSET, 1371600, WIDTH, HEIGHT, 'accent2', '6858000 x 12192000') +
          extent(
            11,
            'Its transpose, which is the corpus default',
            INSET + Math.round(WIDTH / 4) + GUTTER,
            1371600,
            HEIGHT,
            WIDTH,
            'accent3',
            '12192000 x 6858000',
          ) +
          caption(12, 'Transpose caption', 5486400, [
            'Read cx as a height and cy as a width and you get',
            'the shape on the right: an ordinary 16:9 slide,',
            'wrong in a way nothing on the page announces.',
            'The corner markers are what give it away.',
          ]),
      },
      {
        title: 'a42 — a landscape notes page under a portrait slide',
        layout: 1,
        body:
          extent(
            10,
            'This slide',
            INSET,
            1371600,
            WIDTH,
            HEIGHT,
            'accent4',
            'slide 6858000 x 12192000',
          ) +
          extent(
            11,
            'Its notes page',
            INSET + Math.round(WIDTH / 4) + GUTTER,
            1371600,
            NOTES_WIDTH,
            NOTES_HEIGHT,
            'accent5',
            'notes 9144000 x 6858000',
          ) +
          caption(12, 'Notes caption', 5486400, [
            'p:notesSz is required; p:sldSz is optional.',
            'Neither is derived from the other, and here they',
            'disagree in aspect ratio and in orientation.',
          ]),
      },
    ],
  }),
};

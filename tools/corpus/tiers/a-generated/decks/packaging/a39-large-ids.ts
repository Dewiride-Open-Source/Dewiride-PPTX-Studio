import { grid, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The four id spaces, at their edges - and they are four, never one allocator.
 *
 * ## The spaces
 *
 * | id                                | range                        | unique within |
 * | --------------------------------- | ---------------------------- | -------------- |
 * | `p:sldId/@id`                     | 256 .. 2147483647            | the package    |
 * | `p:sldMasterId/@id`               | 2147483648 .. 4294967295     | the package    |
 * | `p:sldLayoutId/@id`               | 2147483648 .. 4294967295     | the package    |
 * | `p:cNvPr/@id`                     | 0 .. 4294967295 in the schema | **one part**  |
 * |                                   | 0 .. 2147483647, and 4294967295, in PowerPoint | |
 *
 * The first three are disjoint by range from the fourth and from each other -
 * except that rows two and three **are the same space**, which is the fact that
 * cost a bisection. `ST_SlideMasterId` and `ST_SlideLayoutId` are separately
 * named types with identical facets, and nothing in any schema says an id from
 * one may not equal an id from the other. PowerPoint refuses a package where
 * they collide, and allocates from a single running counter - master, its
 * layouts, next master, its layouts. `a12-masters` is where that was found:
 * two counters are indistinguishable from correct with one master and collide
 * on the second, so **every** two-master package was a whole-package refusal.
 * See `ROSTER.md`'s "What PowerPoint refuses".
 *
 * ## What this deck sits at
 *
 * **`p:sldId` at both ends, and out of order.** The three slides are
 * 2147483647, 256 and 1073741823 - the top of the range, the bottom, and the
 * middle - in that order. Nothing requires `p:sldIdLst` to ascend; the
 * presentation order is the **element order**, and the ids are labels. A
 * renderer that sorts by id shows this deck backwards, and one that treats an
 * id as an index reads off the end of an array immediately.
 *
 * 256 is the load-bearing end. `ST_SlideId`'s minimum is 256, not 0 and not 1,
 * so a new-slide allocator that starts at zero writes a package no consumer has
 * to accept, and one that starts at `max + 1` fails on the deck below rather
 * than on the thousandth slide.
 *
 * **`p:sldMasterId` and `p:sldLayoutId` at the top.** The master is
 * 4294967293 and its two layouts are 4294967294 and 4294967295 - the last three
 * values the space has. Any allocator that adds one to the highest existing id
 * overflows here. In a language where these are `number` that is silent: the
 * next id is 4294967296, which is not representable as an `xsd:unsignedInt` and
 * which PowerPoint has no reason to accept.
 *
 * **`p:cNvPr/@id` is not the space the schema says it is.** This one was
 * measured rather than read, because the first version of this deck was a
 * whole-package refusal and bisecting it took three rounds. Sweeping one id
 * per package through PowerPoint 16.0.20326 on 2026-08-28:
 *
 * ```
 * 0 .. 2147483647             accepted
 * 2147483648 .. 4294967294    refused, whole package
 * 4294967295                  accepted
 * ```
 *
 * `ST_DrawingElementId` is `xsd:unsignedInt`, so all 4294967296 values are
 * schema-legal and PowerPoint accepts 2147483649 of them. It is reading the
 * attribute as a **signed** 32-bit integer and rejecting everything that comes
 * out negative - except `0xFFFFFFFF`, which is minus one, and which it
 * evidently keeps as a sentinel of its own rather than as an id. The two
 * survivors sit at opposite ends of the declared range with two billion
 * refusals between them, which is not a shape any schema can express and is
 * exactly the sort of thing this corpus exists to write down.
 *
 * Slide 3 carries both survivors and a 3, in one part - because shape ids are
 * unique **within one part**. `a12`'s masters and every slide in this corpus
 * reuse 10, 11, 12 freely, and that is correct; a global registry is the bug,
 * and it is the same bug as the rId one, since rIds are `xsd:ID` scoped to one
 * `.rels`.
 *
 * ## The sentinel is tolerated, not accepted
 *
 * Re-saving through PowerPoint on the same day settles what the open could not.
 * Slide 3's ids go in as `1, 2, 3, 2147483647, 4294967295, 10` and come back as
 * `1, 2, 3, 2147483647, 4, 10`: **`4294967295` is renumbered to `4`** while
 * `2147483647` is kept exactly. So the top of the range is not an id PowerPoint
 * accepts, it is a value it declines to refuse and then replaces - which is
 * what a sentinel looks like from outside, and which no amount of opening the
 * file would have shown.
 *
 * Everything else survives untouched, including the parts most likely to be
 * "helpfully" normalised: `p:sldIdLst` comes back with 2147483647 first and 256
 * second, in that order, and `p:sldMasterId id="4294967293"` is unchanged.
 *
 * ## Why the ids are on the deck and not only in a test
 *
 * Because the failure is not an exception. Every one of these produces a
 * package that opens, and then a renderer that shows the wrong slide first, or
 * a writer that allocates an id that already exists and hands over a file
 * PowerPoint repairs. `a39` is a fixture the allocator can be run against.
 */

/** `ST_SlideId`: 256 .. 2147483647, in the order the deck presents them. */
const SLIDE_IDS = [2147483647, 256, 1073741823] as const;

/** The top of `ST_SlideMasterId`, leaving exactly enough for two layouts. */
const FIRST_SHEET_ID = 4294967293;

/** The top of `ST_DrawingElementId`, accepted only because it is `0xFFFFFFFF`. */
const SENTINEL_SHAPE_ID = 4294967295;

/** The largest id PowerPoint actually accepts: `INT32_MAX`, measured. */
const MAX_SHAPE_ID = 2147483647;

const cell = grid(2, 2);

function box(id: number, name: string, c: Cell, accent: string, lines: readonly string[]): string {
  return shape({
    id,
    name,
    x: c.x,
    y: c.y,
    cx: c.cx,
    cy: c.cy,
    geometry: prstGeom('roundRect'),
    fill: solidFill(scheme(accent, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="ctr"/>',
      paras: lines
        .map((line, i) => textLine(line, { sz: 1100, b: i === 0 }, { algn: 'ctr' }))
        .join(''),
    }),
  });
}

export const a39LargeIds: ProbeDeck = {
  id: 'a39-large-ids',
  title: 'PPTX Studio corpus: a39 large ids',
  description:
    'The four id spaces at their edges. p:sldId at 2147483647, 256 and 1073741823, in that order, ' +
    'because p:sldIdLst is ordered by element and not by id and 256 is the minimum rather than 0. ' +
    'p:sldMasterId at 4294967293 with its layouts at 4294967294 and 4294967295, the last three ' +
    'values the space has, so any allocator that adds one to the maximum overflows. p:cNvPr at ' +
    '2147483647 and 4294967295 beside 3 in one part - which is the whole of what PowerPoint accepts ' +
    'for a p:cNvPr id: swept one value per package, 0 to 2147483647 open, 2147483648 to 4294967294 ' +
    'are whole-package refusals, and 4294967295 opens. It reads a schema-declared xsd:unsignedInt as ' +
    'a signed int32 and keeps 0xFFFFFFFF as a sentinel. Master and layout ids share one space, which ' +
    'is unwritten and which PowerPoint enforces.',
  features: {
    shape: 18,
    placeholder: 6,
    presetGeom: 12,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a39 large ids',
    slideIds: SLIDE_IDS,
    firstSheetId: FIRST_SHEET_ID,
    slides: [
      {
        title: 'a39 — this slide is p:sldId id="2147483647"',
        body:
          box(10, 'Top of the range', cell(0), 'accent1', [
            'p:sldId id="2147483647"',
            'the maximum ST_SlideId,',
            'and the first slide in the deck',
          ]) +
          box(11, 'Order is elements, not ids', cell(1), 'accent2', [
            'p:sldIdLst is a sequence.',
            'The next slide is id 256.',
            'Sort by id and this deck plays backwards.',
          ]) +
          box(12, 'Not an index', cell(2), 'accent3', [
            'An id is a label, not a position.',
            'slides[sldId] reads off the end',
            'of the array on this very slide.',
          ]) +
          box(13, 'The master above this slide', cell(3), 'accent4', [
            'p:sldMasterId id="4294967293"',
            'layouts 4294967294 and 4294967295.',
            'Nothing is left in the space.',
          ]),
      },
      {
        title: 'a39 — this slide is p:sldId id="256"',
        body:
          box(10, 'Bottom of the range', cell(0), 'accent5', [
            'p:sldId id="256"',
            'the minimum ST_SlideId.',
            'Not 0, and not 1.',
          ]) +
          box(11, 'What starts at zero', cell(1), 'accent6', [
            'An allocator that numbers from 0',
            'writes a package no consumer',
            'has to accept.',
          ]) +
          box(12, 'What starts at max + 1', cell(2), 'accent1', [
            'An allocator that adds one to the highest',
            'existing id fails on slide 1 of this deck,',
            'not on the thousandth slide.',
          ]) +
          box(13, 'The same shape ids as slide 1', cell(3), 'accent2', [
            'ids 10 to 13 again, in a different part.',
            'Correct: p:cNvPr is unique within a part.',
            'A package-wide registry is the bug.',
          ]),
      },
      {
        title: 'a39 — shape ids at the top of their own space',
        body:
          box(3, 'A very small id', cell(0), 'accent3', [
            'p:cNvPr id="3"',
            'beside the largest one there is,',
            'in the same p:spTree',
          ]) +
          box(MAX_SHAPE_ID, 'The largest id PowerPoint accepts', cell(1), 'accent4', [
            'p:cNvPr id="2147483647"',
            'INT32_MAX, measured — not the',
            'xsd:unsignedInt maximum the schema says',
          ]) +
          box(SENTINEL_SHAPE_ID, 'The sentinel', cell(2), 'accent5', [
            'p:cNvPr id="4294967295"',
            'accepted, while every value from',
            '2147483648 to 4294967294 is refused',
          ]) +
          box(10, 'And an ordinary one', cell(3), 'accent6', [
            'p:cNvPr id="10"',
            'Three ids, one part, no ordering',
            'and no contiguity required.',
          ]),
      },
    ],
  }),
};

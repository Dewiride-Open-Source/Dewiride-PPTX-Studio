import { grid, group, prstGeom, scheme, shape, solidFill, type Cell } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The empty cases: zero extents, a zero child space, an empty group and an
 * empty shape tree.
 *
 * ## Why a deck of nothings is worth committing
 *
 * Every probe here is schema-legal, and every one of them is a division, an
 * index or a dereference that a renderer does without checking. They are not
 * hostile input - `opc`'s fuzzing handles that - they are the *ordinary* output
 * of a user dragging a shape until it has no width, or of a program that wrote
 * a group and never put anything in it.
 *
 * The list, and what each one breaks:
 *
 * | probe                                  | what it divides by, or dereferences  |
 * | -------------------------------------- | ------------------------------------- |
 * | `a:ext cx="0" cy="0"`                  | any aspect-ratio or scale computation |
 * | `a:ext` zero in one axis only          | the same, in one axis, which is worse |
 * | `a:ext cx="1" cy="1"`                  | rounding to a device pixel            |
 * | `a:chExt cx="0" cy="0"`                | the group child transform             |
 * | a group with no children               | `children[0]`, bounds-of-children     |
 * | `a:path w="0" h="0"`                   | path-space scaling                    |
 * | an empty `p:spTree`                    | first-shape, z-order, hit-test        |
 *
 * ## The two that have a specified answer, and it is not "throw"
 *
 * **`a:chExt` of zero means a scale of 1.0.** A group child at `x` renders at
 * `off.x + (x - chOff.x) * (ext.cx / chExt.cx)`, and the plan's sub-phase 2.10
 * says in as many words that `chExt` of 0 is scale 1, not a divide-by-zero.
 * Slide 2 puts a child in such a group at a known offset so the answer is
 * visible rather than argued: it must land at `off + (x - chOff)`.
 *
 * **`a:path/@w` or `@h` of zero means no path-space scaling**, from sub-phase
 * 2.4, for the same reason and with the same trap. The path's coordinates are
 * then slide coordinates within the shape rather than fractions of it.
 *
 * ## The one that is missing, and why
 *
 * This deck was written with a ninth probe: **a slide with no `slideLayout`
 * relationship**. A slide part does not name its layout anywhere in its own
 * markup - the binding lives only in `ppt/slides/slideN.xml.rels` - so the way
 * to have no layout is to have no relationship, and `CT_Slide` does not require
 * one.
 *
 * PowerPoint refuses it, and the refusal is total. Bisected on 2026-08-28
 * against 16.0.20326, four variants, one change each, all rejected with
 * "PowerPoint could not open the file":
 *
 * - no `.rels` part for the slide at all
 * - a `.rels` part with one relationship, of a different type
 * - a `slideLayout` relationship whose target is a **master**
 * - a `slideLayout` relationship whose target does not exist
 *
 * So the rule is not "the relationship is optional" and not even "at most one":
 * **a slide has exactly one `slideLayout` relationship and it resolves to a
 * slide layout part.** Two of them is a refusal as well. That belongs in
 * sub-phase 1.2's rules, it is written nowhere in any schema, and `ROSTER.md`
 * carries it beside the others.
 *
 * The markup is not in this deck because a Tier A deck opens; it is the first
 * fixture `corpus/reject/` should carry when that collection is built.
 */

const cell = grid(2, 2);
const EMU = 914400;

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

/** A shape with an explicit extent and no caption, so the extent is the probe. */
function extent(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  accent: string,
): string {
  return shape({
    id,
    name,
    x,
    y,
    cx,
    cy,
    fill: solidFill(scheme(accent)),
    caption: false,
  });
}

/**
 * `a:custGeom` whose path space is zero in both axes.
 *
 * `a:path/@w` and `@h` are the coordinate space the path's own numbers are in,
 * and the shape's extent is what they scale to. Zero means no scaling, so these
 * coordinates are EMU inside the shape rather than fractions of it.
 */
const ZERO_PATH_GEOMETRY =
  '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>' +
  '<a:rect l="0" t="0" r="0" b="0"/>' +
  '<a:pathLst><a:path w="0" h="0">' +
  '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
  `<a:lnTo><a:pt x="${String(EMU)}" y="0"/></a:lnTo>` +
  `<a:lnTo><a:pt x="${String(EMU)}" y="${String(EMU / 2)}"/></a:lnTo>` +
  `<a:lnTo><a:pt x="0" y="${String(EMU / 2)}"/></a:lnTo>` +
  '<a:close/>' +
  '</a:path></a:pathLst></a:custGeom>';

/** A text body with the minimum `CT_TextBody` allows: a bodyPr and one empty paragraph. */
const EMPTY_TEXT_BODY = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>';

export const a38Degenerate: ProbeDeck = {
  id: 'a38-degenerate',
  title: 'PPTX Studio corpus: a38 degenerate',
  description:
    'Eight schema-legal empty cases, each of which is a division or a dereference a renderer does ' +
    'without checking: a:ext of zero in both axes and in one, an extent of one EMU, a:chExt of zero ' +
    'on a group, a group with no children, a:path with a zero path space, a p:txBody holding one ' +
    'empty paragraph, and a slide whose p:spTree has no shapes at all. A ninth was built and removed: ' +
    'a slide with no slideLayout relationship, which PowerPoint refuses outright, as it does a ' +
    'relationship pointing at a master, one pointing at nothing, and two of them - so a slide has ' +
    'exactly one, resolving to a layout part, which no schema says.',
  features: {
    shape: 16,
    // Five, not six: slide 3 is on the Blank layout, so it gets no title
    // placeholder from the chassis and there is nothing on it at all.
    placeholder: 5,
    presetGeom: 10,
    customGeom: 1,
    group: 2,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a38 degenerate',
    slides: [
      {
        title: 'a38 — extents that are zero, or nearly',
        body:
          box(10, 'Zero in both axes', cell(0), 'accent1', [
            'a:ext cx="0" cy="0"',
            'the marker below it has no area',
            'and is still a shape in the tree',
          ]) +
          extent(
            11,
            'Zero extent marker',
            cell(0).x + 91440,
            cell(0).y + cell(0).cy - 91440,
            0,
            0,
            'accent1',
          ) +
          box(12, 'Zero in one axis', cell(1), 'accent2', [
            'a:ext cx="0" cy="914400"',
            'an aspect ratio of zero,',
            'not of one',
          ]) +
          extent(
            13,
            'Zero width marker',
            cell(1).x + 91440,
            cell(1).y + 91440,
            0,
            EMU / 2,
            'accent2',
          ) +
          box(14, 'One EMU', cell(2), 'accent3', [
            'a:ext cx="1" cy="1"',
            'a 1/914400 inch shape.',
            'Legal, and rounds to nothing.',
          ]) +
          extent(15, 'One EMU marker', cell(2).x + 91440, cell(2).y + 91440, 1, 1, 'accent3') +
          shape({
            id: 16,
            name: 'Zero path space',
            x: cell(3).x,
            y: cell(3).y,
            cx: cell(3).cx,
            cy: cell(3).cy,
            geometry: ZERO_PATH_GEOMETRY,
            fill: solidFill(scheme('accent4', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
            textBody: txBody({
              bodyPr: '<a:bodyPr wrap="square" anchor="b"/>',
              paras:
                textLine('a:path w="0" h="0"', { sz: 1100, b: true }, { algn: 'ctr' }) +
                textLine('no path-space scaling:', { sz: 1100 }, { algn: 'ctr' }) +
                textLine('the pt values are EMU', { sz: 1100 }, { algn: 'ctr' }),
            }),
          }),
      },
      {
        title: 'a38 — groups with no space and no children',
        body:
          box(10, 'Zero child space', cell(0), 'accent5', [
            'a:chExt cx="0" cy="0"',
            'specified as a scale of 1.0,',
            'not as a division by zero',
          ]) +
          group({
            id: 11,
            name: 'Group with a zero child space',
            x: cell(1).x,
            y: cell(1).y,
            cx: cell(1).cx,
            cy: cell(1).cy,
            childOffsetX: 0,
            childOffsetY: 0,
            childWidth: 0,
            childHeight: 0,
            // At scale 1.0 this child lands at off + (x - chOff), which is
            // exactly the group's own origin plus 91440 in each axis.
            children: extent(12, 'Child at scale one', 91440, 91440, EMU, EMU / 2, 'accent5'),
          }) +
          group({
            id: 13,
            name: 'Group with no children at all',
            x: cell(2).x,
            y: cell(2).y,
            cx: cell(2).cx,
            cy: cell(2).cy,
            children: '',
          }) +
          box(14, 'Empty group', cell(2), 'accent6', [
            'p:grpSp with no children',
            'legal: CT_GroupShape allows zero.',
            'Bounds-of-children has no answer.',
          ]) +
          shape({
            id: 15,
            name: 'Empty text body',
            x: cell(3).x,
            y: cell(3).y,
            cx: cell(3).cx,
            cy: cell(3).cy,
            geometry: prstGeom('roundRect'),
            fill: solidFill(scheme('accent1', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')),
            textBody: EMPTY_TEXT_BODY,
          }),
      },
      {
        // Layout 1 is Blank, so the chassis writes no title placeholder and
        // this slide's `p:spTree` holds nothing but its own prologue.
        title: 'a38 — an empty spTree',
        layout: 1,
        body: '',
      },
    ],
  }),
};

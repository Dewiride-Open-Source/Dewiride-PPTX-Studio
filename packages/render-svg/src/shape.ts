/**
 * One placed shape, as SVG nodes.
 *
 * ## One `<path>` per `a:path`, and that is not a stylistic choice
 *
 * The 187 presets hold 320 paths and 60 of them have paths that disagree about
 * fill or stroke - `smileyFace` draws a filled head and an unfilled mouth, and
 * every one of the 17 action buttons draws a filled plate with an unfilled glyph
 * on it. Merging a shape's paths into one `d` cannot render any of those, which
 * is why `@pptx-studio/geometry` hands over a list and this keeps it a list.
 *
 * ## Text is a sibling, and that is not a stylistic choice either
 *
 * A `flipH` mirrors the outline and leaves the glyphs alone - measured 20 of 20
 * in T8 - so the text cannot sit inside the group that carries the mirror. It
 * gets its own group with its own turn, emitted beside this one.
 *
 * ## The transform, and why the mirror is innermost
 *
 * SVG applies a transform list right to left, so
 * `translate(centre) rotate(a) scale(±1) translate(-half)` moves the local box
 * to the origin, mirrors it, rotates it and puts it on the slide - which is the
 * order C6 measured: PowerPoint asked to mirror a shape already at 30 degrees
 * writes minus thirty with a flip, and only a renderer that flips first has to.
 */

import type { ResolvedPath } from '@pptx-studio/geometry';
import { resolveLine } from '@pptx-studio/paint';

import {
  type Defs,
  effectFilterAttribute,
  fillAttributes,
  strokeAttributes,
  type Attrs,
} from './paint.js';
import { element, type SvgElement, type SvgNode } from './node.js';
import type { Placed } from './layout.js';
import { shapeTextNodes, type TextEngine } from './text/draw.js';
import { frameTransform, type Box } from './transform.js';

/** Identifying attributes, so a caller can hit-test and a human can read a diff. */
function identity(placed: Placed): Attrs {
  return {
    'data-shape': String(placed.shape.cNvPrId),
    'data-name': placed.shape.name === '' ? null : placed.shape.name,
    ...(placed.shape.descr === null ? {} : { 'aria-label': placed.shape.descr }),
  };
}

/** The fill's rectangle, moved into the shape's own coordinate space. */
function localFillBox(placed: Placed): Box {
  return {
    x: placed.fillBox.x - placed.frame.x,
    y: placed.fillBox.y - placed.frame.y,
    cx: placed.fillBox.cx,
    cy: placed.fillBox.cy,
  };
}

function drawablePaths(placed: Placed): readonly ResolvedPath[] {
  // A path whose evaluation reached a non-finite number emits an empty `d`, and
  // a `d` holding a NaN is not merely wrong - a browser drops the whole element
  // and says nothing. `geometry` already refuses to emit one; this drops it.
  return placed.geometry?.paths.filter((path) => path.finite && path.d !== '') ?? [];
}

/**
 * The nodes for one shape, in painting order.
 *
 * A group contributes only its children: measured in C6, **a group's own fill
 * is never painted** - it exists to be asked for by a child's `a:grpFill` - so
 * a group emits a `<g>` for structure and nothing that puts ink down.
 */
export function shapeNodes(
  placed: Placed,
  defs: Defs,
  text: TextEngine | null = null,
): readonly SvgNode[] {
  if (placed.shape.kind === 'grpSp') {
    return [
      element(
        'g',
        identity(placed),
        placed.children.flatMap((child) => shapeNodes(child, defs, text)),
      ),
    ];
  }

  // The shape's own group carries the mirror a `flipH` asks for and text is
  // never mirrored, so the text group is a sibling and not a child.
  const glyphs = text === null ? [] : shapeTextNodes(placed, text);
  const paths = drawablePaths(placed);
  if (paths.length === 0) return glyphs;

  const box: Box = { x: 0, y: 0, cx: placed.frame.cx, cy: placed.frame.cy };
  const fillBox = localFillBox(placed);
  const fill = fillAttributes(placed.fill, placed.colorContext, fillBox, defs);
  const stroke = strokeAttributes(
    placed.appearance.line === null ? null : resolveLine(placed.appearance.line),
    placed.colorContext,
    fillBox,
    defs,
  );

  // `algn="in"` puts the whole stroke band inside the shape. Doubling the width
  // and clipping to the shape's own outline leaves exactly the inner half.
  let clip: string | null = null;
  if (stroke !== null && stroke.inset) {
    const id = defs.id();
    defs.add(
      element(
        'clipPath',
        { id, clipPathUnits: 'userSpaceOnUse' },
        paths.map((path) => element('path', { d: path.d })),
      ),
    );
    clip = `url(#${id})`;
  }

  const children: SvgElement[] = paths.map((path) =>
    element('path', {
      d: path.d,
      // `a:path/@fill="none"` is the path saying it is an outline, not the shape
      // saying it has no fill - `smileyFace`'s mouth against its face.
      ...(path.fill === 'none' ? { fill: 'none' } : fill),
      ...(path.stroke && stroke !== null ? stroke.attrs : { stroke: 'none' }),
      ...(clip !== null && path.stroke ? { 'clip-path': clip } : {}),
    }),
  );

  const transform = frameTransform(placed.frame);
  const filter = effectFilterAttribute(
    placed.appearance.effects ?? [],
    placed.colorContext,
    box,
    defs,
  );

  return [
    element(
      'g',
      { ...identity(placed), ...(transform === '' ? {} : { transform }), ...filter },
      children,
    ),
    ...glyphs,
  ];
}

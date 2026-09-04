/**
 * A resolved fill, stroke or effect into SVG.
 *
 * Every decision behind the numbers is `@pptx-studio/paint`'s and was measured
 * in 2.6 to 2.8; this module only chooses the SVG construct that reproduces
 * them. Two of those choices are load-bearing and both are recorded there:
 *
 * - **`gradientUnits="userSpaceOnUse"`, never `objectBoundingBox`.** The latter
 *   reproduces a gradient's bend and also shears its ramp - the corner of a 3:1
 *   shape lands at 0.21 where PowerPoint puts it at 0.01.
 * - **`patternUnits="userSpaceOnUse"`.** A pattern tile is a physical size, six
 *   points square, not a fraction of the shape. `objectBoundingBox` makes a wide
 *   shape's hatching coarser than a narrow one's and slants every diagonal.
 */

import {
  PATTERN_PIXEL_EMU,
  effectFilter,
  linearGradientVector,
  pathGeometry,
  resolveColor,
  resolvePattern,
  svgStops,
  svgStroke,
  toHexColor,
  type ColorContext,
  type Effect,
  type Fill,
  type ResolvedLine,
  type Rgba,
} from '@pptx-studio/paint';

import { element, num, type AttributeValue, type SvgElement } from './node.js';
import type { Box } from './transform.js';

export type Attrs = Record<string, AttributeValue>;

/** Everything that has to live in `<defs>`, with the ids that reach it. */
export class Defs {
  private readonly nodes: SvgElement[] = [];
  private next = 0;
  /** A per-slide prefix, so two slides in one document cannot collide. */
  private readonly prefix: string;

  // Written out rather than declared as a parameter property: the repository
  // lints for `erasableSyntaxOnly`, so nothing here may need a runtime shim.
  constructor(prefix: string) {
    this.prefix = prefix;
  }

  id(): string {
    this.next += 1;
    return `${this.prefix}-${String(this.next)}`;
  }

  add(node: SvgElement): void {
    this.nodes.push(node);
  }

  get length(): number {
    return this.nodes.length;
  }

  toNode(): SvgElement | null {
    return this.nodes.length === 0 ? null : element('defs', {}, this.nodes);
  }
}

/** `toHexColor` writes OOXML's bare `RRGGBB`; CSS and SVG want the hash. */
function css(color: Rgba): string {
  return `#${toHexColor(color)}`;
}

function colorAttrs(color: Rgba, prefix: 'fill' | 'stroke' | 'flood'): Attrs {
  const attrs: Attrs = { [prefix]: css(color) };
  if (color.a < 1) attrs[`${prefix}-opacity`] = Math.round(color.a * 1000) / 1000;
  return attrs;
}

/* -------------------------------------------------------------------------- */
/* gradients                                                                  */
/* -------------------------------------------------------------------------- */

function stopNodes(fill: Extract<Fill, { type: 'gradient' }>, ctx: ColorContext): SvgElement[] {
  return svgStops(fill, ctx).map((stop) =>
    element('stop', {
      // `svgStops` already returns a fraction. Dividing by a hundred thousand
      // here - as if the offset were still `ST_PositiveFixedPercentage` - put
      // every stop at zero and painted the whole shape in the last stop's
      // colour, which no unit test noticed and the pixel comparison against
      // PowerPoint caught on seventeen samples at once.
      offset: num(stop.offset, 6),
      'stop-color': stop.color,
      ...(stop.opacity < 1 ? { 'stop-opacity': stop.opacity } : {}),
    }),
  );
}

function gradientPaint(
  fill: Extract<Fill, { type: 'gradient' }>,
  ctx: ColorContext,
  box: Box,
  defs: Defs,
): Attrs {
  const id = defs.id();
  const stops = stopNodes(fill, ctx);

  if (fill.shade === null || fill.shade.kind === 'linear') {
    // No `a:lin` and no `a:path` at all: PowerPoint paints the ramp left to
    // right, which is what a zero angle gives.
    const shade = fill.shade ?? { kind: 'linear' as const, ang: 0, scaled: false };
    const vector = linearGradientVector(shade, box.cx, box.cy);
    defs.add(
      element(
        'linearGradient',
        {
          id,
          gradientUnits: 'userSpaceOnUse',
          x1: box.x + vector.x1,
          y1: box.y + vector.y1,
          x2: box.x + vector.x2,
          y2: box.y + vector.y2,
        },
        stops,
      ),
    );
    return { fill: `url(#${id})` };
  }

  const geometry = pathGeometry(fill.shade, box.cx, box.cy);
  const cx = box.x + geometry.fx * box.cx;
  const cy = box.y + geometry.fy * box.cy;

  // A `path="circle"` really is a circle in shape units, so it goes out as one.
  // `rect` and `shape` are a Chebyshev distance - concentric rectangles - which
  // SVG has no primitive for; an ellipse inscribed in the shape is the closest
  // single element and is visibly nearer than a circle on anything but a square.
  // Recorded as an open question rather than as a solved one.
  if (geometry.metric === 'radial') {
    defs.add(
      element(
        'radialGradient',
        { id, gradientUnits: 'userSpaceOnUse', cx, cy, r: geometry.radius, fx: cx, fy: cy },
        stops,
      ),
    );
  } else {
    const rx = Math.max(box.cx, 1) / 2;
    const ry = Math.max(box.cy, 1) / 2;
    defs.add(
      element(
        'radialGradient',
        {
          id,
          gradientUnits: 'userSpaceOnUse',
          cx: 0,
          cy: 0,
          r: 1,
          gradientTransform: `translate(${num(cx)} ${num(cy)}) scale(${num(rx)} ${num(ry)})`,
        },
        stops,
      ),
    );
  }
  return { fill: `url(#${id})` };
}

/* -------------------------------------------------------------------------- */
/* patterns                                                                   */
/* -------------------------------------------------------------------------- */

function patternPaint(
  fill: Extract<Fill, { type: 'pattern' }>,
  ctx: ColorContext,
  box: Box,
  defs: Defs,
): Attrs {
  const id = defs.id();
  const tile = resolvePattern(fill, ctx);
  const cells: SvgElement[] = [];
  for (let row = 0; row < tile.rows; row++) {
    for (let col = 0; col < tile.cols; col++) {
      const coverage = tile.coverage[row * tile.cols + col] ?? 0;
      if (coverage <= 0) continue;
      cells.push(
        element('rect', {
          x: col * PATTERN_PIXEL_EMU,
          y: row * PATTERN_PIXEL_EMU,
          width: PATTERN_PIXEL_EMU,
          height: PATTERN_PIXEL_EMU,
          ...colorAttrs({ ...tile.fg, a: tile.fg.a * coverage }, 'fill'),
        }),
      );
    }
  }
  defs.add(
    element(
      'pattern',
      {
        id,
        patternUnits: 'userSpaceOnUse',
        // The tile is anchored to the box the fill is laid out over, so a
        // `grpFill` child continues the group's tiling rather than restarting it.
        x: box.x,
        y: box.y,
        width: tile.width,
        height: tile.height,
      },
      [
        element('rect', {
          x: 0,
          y: 0,
          width: tile.width,
          height: tile.height,
          ...colorAttrs(tile.bg, 'fill'),
        }),
        ...cells,
      ],
    ),
  );
  return { fill: `url(#${id})` };
}

/* -------------------------------------------------------------------------- */
/* fills                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * The `fill` attributes for one shape.
 *
 * `box` is the rectangle the fill is laid out over, in the same coordinate
 * space as the path being filled. For an ordinary shape that is its own frame;
 * for an `a:grpFill` it is the enclosing group's content bounds, which is the
 * measured rule and the reason this takes a box at all rather than a size.
 */
export function fillAttributes(fill: Fill | null, ctx: ColorContext, box: Box, defs: Defs): Attrs {
  if (fill === null) return { fill: 'none' };
  switch (fill.type) {
    case 'solid':
      return colorAttrs(resolveColor(fill.color, ctx), 'fill');
    case 'gradient':
      return gradientPaint(fill, ctx, box, defs);
    case 'pattern':
      return patternPaint(fill, ctx, box, defs);
    case 'blip':
      // 2.10 draws geometry. The image pipeline is Phase 4's, and a shape whose
      // picture has not been decoded paints nothing rather than a grey box -
      // which is also what happens to one whose blip is missing.
      return { fill: 'none' };
    case 'group':
      // A `grpFill` that reached no enclosing group paints nothing. Measured on
      // a top-level shape and on a group whose own fill is `a:noFill`.
      return { fill: 'none' };
    case 'none':
    default:
      return { fill: 'none' };
  }
}

/* -------------------------------------------------------------------------- */
/* strokes                                                                    */
/* -------------------------------------------------------------------------- */

export interface StrokePaint {
  readonly attrs: Attrs;
  /** `algn="in"`: the stroke is drawn double-width and clipped to the shape. */
  readonly inset: boolean;
  readonly line: ResolvedLine;
}

/**
 * The `stroke` attributes for one shape, or `null` when nothing is stroked.
 *
 * `@cmpd` is honoured as a single rail, per the plan: the five compound forms
 * subdivide the width into rails at signed offsets from the geometry, and SVG
 * cannot offset a path. The attribute is preserved in the model and round-trips
 * either way, and `compoundRails` in `paint` has the measured table for when a
 * renderer can use it.
 */
export function strokeAttributes(
  line: ResolvedLine | null,
  ctx: ColorContext,
  box: Box,
  defs: Defs,
): StrokePaint | null {
  if (line === null || line.fill.type === 'none' || line.width <= 0) return null;

  const svg = svgStroke(line);
  const attrs: Attrs = {
    'stroke-linecap': svg['stroke-linecap'],
    'stroke-linejoin': svg['stroke-linejoin'],
    'stroke-miterlimit': svg['stroke-miterlimit'],
  };
  // Inset alignment puts the whole band inside the shape. Drawing it at double
  // width and clipping to the shape's own path leaves exactly the inner half,
  // which is the standard construction and the one the plan names.
  attrs['stroke-width'] = line.algn === 'in' ? line.width * 2 : line.width;
  if (svg['stroke-dasharray'] !== null) {
    const array = line.algn === 'in' ? svgStroke(line, line.width * 2) : svg;
    attrs['stroke-dasharray'] = (array['stroke-dasharray'] ?? [])
      .map((value) => num(value))
      .join(' ');
  }

  if (line.fill.type === 'solid') {
    Object.assign(attrs, colorAttrs(resolveColor(line.fill.color, ctx), 'stroke'));
  } else {
    // A gradient or pattern stroke: the same paint server, referenced from
    // `stroke` instead of `fill`.
    const paint = fillAttributes(line.fill, ctx, box, defs);
    attrs['stroke'] = paint['fill'] ?? 'none';
    if (paint['fill-opacity'] !== undefined) attrs['stroke-opacity'] = paint['fill-opacity'];
  }

  return { attrs, inset: line.algn === 'in', line };
}

/* -------------------------------------------------------------------------- */
/* effects                                                                    */
/* -------------------------------------------------------------------------- */

const PRIMITIVE_TAG = {
  gaussian: 'feGaussianBlur',
  offset: 'feOffset',
  flood: 'feFlood',
  composite: 'feComposite',
  merge: 'feMerge',
  morphology: 'feMorphology',
  matrix: 'feTransform',
} as const;

/**
 * The `filter` attribute for one shape's effects, or `null` for none.
 *
 * `paint` builds the graph and this writes it out. One primitive has no SVG
 * counterpart: an outer shadow's full transform - scale, skew and offset
 * together - is a `matrix`, and SVG filters have no affine primitive. It is
 * emitted as an `feOffset` of the matrix's translation, which is exact for the
 * shadows that only offset and is the visible part of the rest.
 */
export function effectFilterAttribute(
  effects: readonly Effect[],
  ctx: ColorContext,
  box: Box,
  defs: Defs,
): Attrs {
  if (effects.length === 0) return {};
  const graph = effectFilter(effects, { x: 0, y: 0, w: box.cx, h: box.cy }, (color) => {
    const rgba = resolveColor(color, ctx);
    return { css: css(rgba), alpha: rgba.a };
  });
  if (graph.primitives.length === 0) return {};

  const nodes = graph.primitives.map((primitive) => {
    switch (primitive.op) {
      case 'gaussian':
        return element('feGaussianBlur', {
          in: primitive.in,
          result: primitive.result,
          stdDeviation: primitive.stdDeviation,
        });
      case 'offset':
        return element('feOffset', {
          in: primitive.in,
          result: primitive.result,
          dx: primitive.dx,
          dy: primitive.dy,
        });
      case 'flood':
        return element('feFlood', {
          result: primitive.result,
          'flood-color': primitive.color,
          'flood-opacity': primitive.opacity,
        });
      case 'composite':
        return element('feComposite', {
          in: primitive.in,
          in2: primitive.in2,
          result: primitive.result,
          operator: primitive.operator,
        });
      case 'merge':
        return element(
          'feMerge',
          { result: primitive.result },
          primitive.inputs.map((input) => element('feMergeNode', { in: input })),
        );
      case 'morphology':
        return element('feMorphology', {
          in: primitive.in,
          result: primitive.result,
          operator: primitive.operator,
          radius: primitive.radius,
        });
      case 'matrix':
      default:
        return element('feOffset', {
          in: primitive.in,
          result: primitive.result,
          dx: primitive.matrix[4],
          dy: primitive.matrix[5],
        });
    }
  });

  const id = defs.id();
  // The region has to hold every effect's reach or SVG clips it. The default
  // -10%/+20% is not enough for an ordinary 24pt shadow on a small shape.
  const left = graph.margin.left;
  const top = graph.margin.top;
  defs.add(
    element(
      'filter',
      {
        id,
        filterUnits: 'userSpaceOnUse',
        x: -left,
        y: -top,
        width: box.cx + left + graph.margin.right,
        height: box.cy + top + graph.margin.bottom,
      },
      nodes,
    ),
  );
  return { filter: `url(#${id})` };
}

export { PRIMITIVE_TAG };

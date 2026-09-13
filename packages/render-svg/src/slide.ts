/**
 * A whole sheet, as one `<svg>`.
 *
 * ## User units are EMU
 *
 * The viewBox is the slide in English Metric Units, which is what
 * `p:sldSz` says and what every number the model and `paint` hand over is
 * already in - a stroke width, a pattern tile's six points, an effect's blur
 * radius. Rescaling any of them at the boundary is a conversion that has to be
 * right in a dozen places instead of none, and the plan's fixed-space rule says
 * the same thing from the other side: lay out once, and let one `transform:
 * scale(z)` do the zooming, because font metrics are not linear in point size.
 */

import type { MediaResolver } from './image/blip.js';
import { RenderError } from './errors.js';
import { colorContextOf, resolveBackground, type Sheet } from '@pptx-studio/model';

import { layoutSheet, layoutSlide, type Placed } from './layout.js';
import { element, serializeSvg, type SvgElement } from './node.js';
import { Defs, fillAttributes } from './paint.js';
import { shapeNodes } from './shape.js';
import { createTextEngine, type TextOptions } from './text/draw.js';
import { EMU_PER_POINT, type Box } from './transform.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** The slide size in EMU, from `p:presentation/p:sldSz`. */
export interface SlideSize {
  readonly cx: number;
  readonly cy: number;
}

export interface RenderOptions {
  /**
   * A prefix for every generated `id`.
   *
   * Two slides in one HTML document share a URL space, so a `url(#g-1)` from
   * the second silently paints with the first's gradient unless the ids differ.
   * Defaults to something unique per call.
   */
  readonly idPrefix?: string;
  /** The `width` attribute in CSS pixels, which also names the device scale strokes round to. */
  readonly width?: number;
  /** The `height` attribute in CSS pixels: a box the slide is stretched into, as the export is (F2). */
  readonly height?: number;
  /** Device pixels per CSS pixel, so a 2x display rounds strokes to its own pixels. Default 1. */
  readonly devicePixelRatio?: number;
  /** Leave off `xmlns`, for a fragment going into an existing SVG document. */
  readonly inline?: boolean;
  /** Draw the master's and layout's shapes beneath the sheet's own. Default true. */
  readonly inherited?: boolean;
  /**
   * How an image fill or picture reaches its bytes.
   *
   * The renderer resolves no relationships: a fill written in a layout means
   * the layout part own rels, and only the caller knows which part it came from.
   */
  readonly media?: MediaResolver;
  /**
   * How to draw text, or `false` to draw none.
   *
   * Text needs a measuring surface, so a caller with no `OffscreenCanvas` - a
   * Node script, a test asserting geometry alone - turns it off rather than
   * getting a throw from a package it did not know it was using.
   */
  readonly text?: TextOptions | false;
}

/** Device pixels to the point, or `null` when no width names a device. */
function devicePxPerPt(options: RenderOptions, size: SlideSize): number | null {
  const ratio = options.devicePixelRatio ?? 1;
  if (!Number.isFinite(ratio) || ratio <= 0) {
    throw new RenderError(
      'RENDER_DEVICE_PIXEL_RATIO',
      `a device pixel ratio of ${String(ratio)} is not a positive number`,
      String(ratio),
    );
  }
  if (options.width === undefined) {
    if (options.devicePixelRatio !== undefined) {
      throw new RenderError(
        'RENDER_DEVICE_PIXEL_RATIO',
        'a device pixel ratio without a width names no device',
        String(ratio),
      );
    }
    return null;
  }
  return (options.width * ratio) / (size.cx / EMU_PER_POINT);
}

let counter = 0;

function nextPrefix(): string {
  counter += 1;
  return `px${String(counter)}`;
}

/** The background rectangle, or `null` when no sheet in the chain declares one. */
function backgroundNode(sheet: Sheet, size: SlideSize, defs: Defs): SvgElement | null {
  const background = resolveBackground(sheet);
  if (background === null) return null;
  const box: Box = { x: 0, y: 0, cx: size.cx, cy: size.cy };
  const ctx = colorContextOf(
    background.sheet,
    background.phClr === null ? undefined : background.phClr,
  );
  const attrs = fillAttributes(background.fill, ctx, box, defs);
  if (attrs['fill'] === 'none') return null;
  return element('rect', {
    'data-role': 'background',
    x: 0,
    y: 0,
    width: size.cx,
    height: size.cy,
    ...attrs,
  });
}

export interface SlideRender {
  readonly node: SvgElement;
  readonly placed: readonly Placed[];
}

/**
 * The node tree for one sheet.
 *
 * Returned alongside the placed shapes rather than instead of them, because
 * `render-dom` wants both: the tree to mount, and the layout to hit-test
 * against without walking the DOM back.
 */
export function slideNode(sheet: Sheet, size: SlideSize, options: RenderOptions = {}): SlideRender {
  // A named width names the device scale, and strokes are drawn as the export draws them at it.
  const pxPerPt = devicePxPerPt(options, size);
  const defs = new Defs(options.idPrefix ?? nextPrefix(), options.media, pxPerPt);
  const placed = options.inherited === false ? layoutSheet(sheet) : layoutSlide(sheet);
  const background = backgroundNode(sheet, size, defs);
  const engine = options.text === false ? null : createTextEngine(options.text ?? {});
  const shapes = placed.flatMap((shape) => shapeNodes(shape, defs, engine));

  // The defs are collected while the shapes are built, so the node has to be
  // asked for afterwards - and it goes first, because a `url(#...)` reference
  // that precedes its target is legal but reads badly in a diff.
  const defsNode = defs.toNode();
  const children = [
    ...(defsNode === null ? [] : [defsNode]),
    ...(background === null ? [] : [background]),
    ...shapes,
  ];

  return {
    node: element(
      'svg',
      {
        ...(options.inline === true ? {} : { xmlns: SVG_NS }),
        viewBox: `0 0 ${String(size.cx)} ${String(size.cy)}`,
        ...(options.width === undefined ? {} : { width: options.width }),
        ...(options.height === undefined ? {} : { height: options.height }),
        // A named box is filled to the pixel, as the export fills its own; otherwise the aspect holds.
        preserveAspectRatio: options.height === undefined ? 'xMidYMid meet' : 'none',
      },
      children,
    ),
    placed,
  };
}

/** One sheet as an SVG string. */
export function renderSlide(sheet: Sheet, size: SlideSize, options: RenderOptions = {}): string {
  return serializeSvg(slideNode(sheet, size, options).node);
}

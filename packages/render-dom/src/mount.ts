/**
 * The live renderer: the same node tree as `render-svg`, as real elements.
 *
 * ## Why this is thirty lines of DOM and not a second renderer
 *
 * The plan asks for two renderers "over the same layout engine", so that a
 * thumbnail and the editor cannot disagree about what is on the screen. Sharing
 * a *layout* is not enough for that on its own: emitting a `<linearGradient>`
 * twice, once as text and once as `createElementNS`, is two chances to write the
 * same gradient wrong in two different ways.
 *
 * So the shared surface goes a level deeper. `render-svg` builds a node tree and
 * serialises it; this builds the identical tree and instantiates it. Everything
 * either of them knows about DrawingML is in that tree before either touches it,
 * and the only thing this file decides is which DOM call makes a node.
 *
 * ## The namespace, and the four attributes that need their own
 *
 * `document.createElement` makes an HTML element that happens to be spelled
 * `path`, and it draws nothing at all - silently. Every element here goes
 * through `createElementNS` with the SVG namespace. Attributes are the mirror
 * image: they belong to **no** namespace and must go through `setAttribute`,
 * with the single exception of `xlink:href`, which this package never emits.
 *
 * ## Zoom is a transform, never a re-layout
 *
 * The mounted `<svg>` carries a `viewBox` in EMU and is sized in CSS pixels, so
 * scaling it is one attribute and no arithmetic. The plan's rule is the reason:
 * font metrics are not linear in point size, so a renderer that re-lays out per
 * zoom level changes the line breaks as the user zooms.
 */

import {
  serializeSvg,
  slideNode,
  type Placed,
  type RenderOptions,
  type SlideSize,
  type SvgElement,
  type SvgNode,
} from '@pptx-studio/render-svg';
import type { Sheet } from '@pptx-studio/model';

import { RenderDomError } from './errors.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** One node into one element. Exported because the debug overlay wants it too. */
export function createNode(document: Document, node: SvgNode): Node {
  if (node.kind === 'text') return document.createTextNode(node.text);
  const el = document.createElementNS(SVG_NS, node.tag);
  for (const [name, value] of Object.entries(node.attrs)) {
    if (value === null || value === undefined) continue;
    el.setAttribute(name, typeof value === 'number' ? String(value) : value);
  }
  for (const child of node.children) el.appendChild(createNode(document, child));
  return el;
}

export interface MountOptions extends RenderOptions {
  /**
   * The `document` to build in.
   *
   * Passed rather than reached for, because this package must work in a Web
   * Worker with an `OffscreenCanvas` and no global `document` - and because a
   * test that has to stub a global is a test that can pass for the wrong reason.
   */
  readonly document?: Document;
}

export interface MountedSlide {
  readonly root: SVGSVGElement;
  readonly placed: readonly Placed[];
  /** The element drawing a given `p:cNvPr/@id`, for hit-testing and selection. */
  element(cNvPrId: number): SVGElement | null;
  /** Resize without re-laying anything out: the viewBox does the work. */
  resize(width: number, height: number): void;
  /** Detach from the host. The nodes stay valid, so a caller may re-attach. */
  unmount(): void;
}

function ownerDocument(host: Element, options: MountOptions): Document {
  const document = options.document ?? host.ownerDocument;

  if (document === null || document === undefined) {
    throw new RenderDomError('RENDER_DOM_NO_DOCUMENT', 'the host element has no ownerDocument');
  }
  return document;
}

/**
 * Render a sheet into a host element.
 *
 * The host is emptied first. Anything a caller wants to keep beside the slide -
 * the overlay canvas of 2.11, the text layer of 3.8 - is a sibling of the host,
 * not a child of it.
 */
export function mountSlide(
  host: Element,
  sheet: Sheet,
  size: SlideSize,
  options: MountOptions = {},
): MountedSlide {
  // `typeof null` is 'object', so the null check has to come first or the
  // guard reads a property off it and throws the wrong error.
  if (host === null || typeof host !== 'object' || typeof host.appendChild !== 'function') {
    throw new RenderDomError('RENDER_DOM_NO_HOST', 'mountSlide needs an element to build into');
  }
  const document = ownerDocument(host, options);
  const rendered = slideNode(sheet, size, options);
  const root = createNode(document, rendered.node) as SVGSVGElement;

  host.replaceChildren(root);

  const byId = new Map<number, SVGElement>();
  const walk = (node: Element): void => {
    const id = node.getAttribute('data-shape');
    if (id !== null && !byId.has(Number(id))) byId.set(Number(id), node as SVGElement);
    for (const child of Array.from(node.children)) walk(child);
  };
  walk(root);

  return {
    root,
    placed: rendered.placed,
    element(cNvPrId: number): SVGElement | null {
      return byId.get(cNvPrId) ?? null;
    },
    resize(width: number, height: number): void {
      root.setAttribute('width', String(width));
      root.setAttribute('height', String(height));
    },
    unmount(): void {
      root.remove();
    },
  };
}

/**
 * The same slide as a string, from this package.
 *
 * A convenience rather than a second implementation: it calls `render-svg`'s
 * serializer on the same tree `mountSlide` instantiates, which is the property
 * the tests check.
 */
export function renderSlideMarkup(
  sheet: Sheet,
  size: SlideSize,
  options: RenderOptions = {},
): string {
  return serializeSvg(slideNode(sheet, size, options).node);
}

export type { SvgElement };

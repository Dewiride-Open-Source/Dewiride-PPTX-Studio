/**
 * The debug overlay, live: the same nodes `render-svg` builds, plus a pointer.
 *
 * Sub-phase 2.11. Everything drawn here comes from `shapeOverlay`, so the
 * handle a user grabs is at the position a thumbnail would have drawn it - the
 * overlay is not a second opinion about where the geometry went. What this file
 * adds is the four things a pointer needs and a node tree cannot have: an
 * element to mount into, hit-testing, the drag, and the inverse transform that
 * gets a pointer from the slide into the shape's own space.
 *
 * ## Slide space, shape space, and why the inverse has to be exact
 *
 * A shape's adjust handles live in *its* coordinates: `(0, 0)` at the top left
 * of `a:ext`, before any rotation or mirroring. A pointer arrives in the
 * slide's. `inverseFramePoint` undoes exactly what `frameTransform` did, in the
 * reverse order - so the mirror is undone *after* the rotation, because it was
 * applied *before* it. Get that backwards and handles work perfectly on every
 * unrotated shape and drift on precisely the flipped, turned ones where a user
 * would never guess the renderer was at fault.
 *
 * ## The drag writes adjust values, not positions
 *
 * `dragHandle` inverts the handle's own formula and returns `{ guideName:
 * value }` for the guides that handle controls and nothing else, already
 * rounded to the integers the file will hold. So a drag produces an `a:avLst`
 * edit, the shape is re-evaluated from it, and the handle lands wherever the
 * geometry says it lands - which is why a handle confined to a curve does not
 * follow the pointer off it. That is the correct behaviour and PowerPoint's.
 */

import { dragHandle, type Point, type PresetGuide } from '@pptx-studio/geometry';
import {
  inverseFramePoint,
  shapeOverlay,
  type OverlayOptions,
  type Placed,
  type ShapeOverlay,
} from '@pptx-studio/render-svg';

import { RenderDomError } from './errors.js';
import { createNode } from './mount.js';

/** What a completed drag produced. */
export interface HandleEdit {
  /** Which `a:ahLst` entry was dragged. */
  readonly handle: number;
  /** The adjust values this handle controls, rounded as the file will hold them. */
  readonly adjust: Readonly<Record<string, number>>;
  /** The shape's whole `a:avLst` after the edit, ready to write back. */
  readonly avLst: readonly PresetGuide[];
}

export interface OverlayMountOptions extends OverlayOptions {
  readonly document?: Document;
  /** Called on every pointer move during a drag, and once more on release. */
  onChange?: (edit: HandleEdit) => void;
  /** Called once on release. */
  onCommit?: (edit: HandleEdit) => void;
}

export interface MountedOverlay {
  readonly root: SVGGElement;
  readonly overlay: ShapeOverlay;
  /** Redraw against a new placement - after a drag, or a resize. */
  update(placed: Placed): void;
  /** True while a handle is being dragged. */
  dragging(): boolean;
  unmount(): void;
}

/**
 * A point in the host `<svg>`'s user space - EMU - from a pointer event.
 *
 * Through the SVG CTM rather than by arithmetic on `getBoundingClientRect`,
 * because the slide may be inside any number of CSS transforms and the browser
 * already knows the composed matrix. `getScreenCTM` returns null for an element
 * that is not being rendered, which a drag cannot reach but a stray event can.
 */
export function slidePoint(root: SVGSVGElement, clientX: number, clientY: number): Point | null {
  const ctm = root.getScreenCTM();
  if (ctm === null) return null;
  const point = root.createSVGPoint();
  point.x = clientX;
  point.y = clientY;
  const local = point.matrixTransform(ctm.inverse());
  return { x: local.x, y: local.y };
}

/**
 * The `a:avLst` a drag leaves behind.
 *
 * The dragged guides replace their namesakes and everything else is kept in
 * place, because an `a:avLst` is an ordered list and a shape with three adjusts
 * whose second one moves must still write three, in order. A guide the shape did
 * not previously declare is appended: the preset's default was in force and the
 * drag has just overridden it.
 */
export function mergeAdjust(
  avLst: readonly PresetGuide[],
  adjust: Readonly<Record<string, number>>,
): readonly PresetGuide[] {
  const seen = new Set<string>();
  const merged: PresetGuide[] = avLst.map((guide) => {
    const value = adjust[guide.name];
    if (value === undefined) return guide;
    seen.add(guide.name);
    return { name: guide.name, fmla: ['val', String(Math.round(value))] };
  });
  for (const [name, value] of Object.entries(adjust)) {
    if (seen.has(name)) continue;
    merged.push({ name, fmla: ['val', String(Math.round(value))] });
  }
  return merged;
}

/**
 * Draw the overlay for one shape into a host `<svg>`, and make its handles drag.
 *
 * The host must be the same `<svg>` the slide is in, because the overlay is
 * emitted in slide coordinates and carries the shape's own transform - the same
 * arrangement `mountSlide` produces, so an editor appends this to that root.
 */
export function mountOverlay(
  root: SVGSVGElement,
  placed: Placed,
  options: OverlayMountOptions = {},
): MountedOverlay {
  if (root === null || typeof root !== 'object' || typeof root.appendChild !== 'function') {
    throw new RenderDomError(
      'RENDER_DOM_NO_HOST',
      'mountOverlay needs the slide <svg> to build in',
    );
  }
  const document = options.document ?? root.ownerDocument;
  if (document === null || document === undefined) {
    throw new RenderDomError('RENDER_DOM_NO_DOCUMENT', 'the host element has no ownerDocument');
  }

  let current = placed;
  let overlay = shapeOverlay(current, options);
  let group = createNode(document, overlay.node) as SVGGElement;
  root.appendChild(group);

  let active: number | null = null;
  let pointer: number | null = null;

  function edit(at: Point, index: number): HandleEdit | null {
    const source = current.geometrySource;
    const handle = overlay.handles.find((h) => h.index === index);
    if (source === null || handle === undefined) return null;
    const local = inverseFramePoint(current.frame, at);
    const adjust = dragHandle(
      source.geometry,
      { w: current.frame.cx, h: current.frame.cy },
      handle.handle,
      local,
      { adjust: source.adjust },
    );
    if (Object.keys(adjust).length === 0) return null;
    return { handle: index, adjust, avLst: mergeAdjust(source.adjust, adjust) };
  }

  /**
   * Hold the pointer for the duration of the drag.
   *
   * On the root rather than on the handle, because the handle element is
   * replaced on every update and a captured element that leaves the tree takes
   * the capture with it - the drag would stop the moment the shape redrew.
   *
   * `setPointerCapture` throws `NotFoundError` for a pointer that is no longer
   * active, which is reachable whenever the button is released between the
   * event being queued and this running. Losing the capture degrades the drag
   * to one that stops at the window edge; throwing would abandon it entirely.
   */
  function capture(pointerId: number): void {
    try {
      root.setPointerCapture(pointerId);
    } catch {
      // No capture, and the drag continues without it.
    }
  }

  function onPointerDown(event: PointerEvent): void {
    const target = event.target as Element | null;
    const owner = target?.closest?.('[data-role="handle"]');
    const index = owner?.getAttribute('data-handle');
    // A handle with no `gdRef` at all controls nothing and must not swallow
    // the gesture. No preset has one - all 120 with an `a:ahLst` declare at
    // least one axis - but an `a:custGeom` in a real file may.
    if (index === null || index === undefined || owner?.getAttribute('data-axes') === '') return;
    active = Number(index);
    pointer = event.pointerId;
    capture(event.pointerId);
    event.preventDefault();
  }

  function onPointerMove(event: PointerEvent): void {
    if (active === null || event.pointerId !== pointer) return;
    const at = slidePoint(root, event.clientX, event.clientY);
    if (at === null) return;
    const result = edit(at, active);
    if (result !== null) options.onChange?.(result);
  }

  function onPointerUp(event: PointerEvent): void {
    if (active === null || event.pointerId !== pointer) return;
    const at = slidePoint(root, event.clientX, event.clientY);
    const result = at === null ? null : edit(at, active);
    active = null;
    pointer = null;
    if (root.hasPointerCapture(event.pointerId)) root.releasePointerCapture(event.pointerId);
    if (result !== null) options.onCommit?.(result);
  }

  // One signal for all four, so `unmount` cannot detach three of them and
  // leave the fourth listening on a root that outlives this overlay. Four
  // paired `removeEventListener` calls are four chances to forget one, and a
  // half-removed listener is invisible until a second overlay is mounted on the
  // same slide and the dead one starts editing a shape that is no longer there.
  const listening = new AbortController();
  const on = { signal: listening.signal };
  root.addEventListener('pointerdown', onPointerDown, on);
  root.addEventListener('pointermove', onPointerMove, on);
  root.addEventListener('pointerup', onPointerUp, on);
  root.addEventListener('pointercancel', onPointerUp, on);

  return {
    get root(): SVGGElement {
      return group;
    },
    get overlay(): ShapeOverlay {
      return overlay;
    },
    update(next: Placed): void {
      current = next;
      overlay = shapeOverlay(current, options);
      const replacement = createNode(document, overlay.node) as SVGGElement;
      group.replaceWith(replacement);
      group = replacement;
    },
    dragging(): boolean {
      return active !== null;
    },
    unmount(): void {
      listening.abort();
      group.remove();
    },
  };
}

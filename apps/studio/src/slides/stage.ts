/**
 * The stage: one slide, at one zoom, with the debug overlay on whichever shape is selected.
 *
 * A zoom is a fresh mount at that width, not a resize: strokes are rounded to the device
 * pixels of the mounted width (F2), so the SVG for 25 % is not the SVG for 400 %. The
 * text layout is the same at every zoom.
 */

import {
  mountOverlay,
  mountSlide,
  type MountedOverlay,
  type MountedSlide,
} from '@pptx-studio/render-dom';
import { flatten, type Placed } from '@pptx-studio/render-svg';

import { el } from '../element.js';
import type { Deck } from './deck.js';
import { describe } from './inspector.js';
import { nextFrame, type ShowTimings } from './timings.js';
import { stageSizeAt, unitAt } from './zoom.js';

/** A shape as drawn, in slide EMU, for a harness that scores the stage. */
export interface ShapeFrame {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly rot: number;
}

/** A rectangle in page coordinates, which is what a screenshot clips to. */
export interface PageBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface Stage {
  /** Show a slide at a zoom; resolves after the frame that painted it. */
  show(index: number, zoom: number): Promise<ShowTimings>;
  /** The mounted `<svg>` serialised, for comparison across zooms; never for re-parsing. */
  svg(): string;
  shapes(): readonly ShapeFrame[];
  box(): PageBox;
  setOverlay(on: boolean): void;
  unmount(): void;
}

/** The page rectangle of an element, scrolled into view first so the box is what is painted. */
export function pageBoxOf(element: Element): PageBox {
  element.scrollIntoView({ block: 'start', inline: 'start' });
  const rect = element.getBoundingClientRect();
  return {
    x: rect.left + window.scrollX,
    y: rect.top + window.scrollY,
    width: rect.width,
    height: rect.height,
  };
}

export function createStage(host: HTMLElement, inspector: HTMLElement, deck: Deck): Stage {
  host.dataset['role'] = 'stage';
  let mounted: MountedSlide | null = null;
  let shown: { index: number; zoom: number; ratio: number } | null = null;
  let overlay: MountedOverlay | null = null;
  let showOverlay = true;
  let selected: number | null = null;

  function placed(): readonly Placed[] {
    return mounted === null ? [] : flatten(mounted.placed);
  }

  function drawOverlay(): void {
    overlay?.unmount();
    overlay = null;
    const target = placed().find((p) => p.shape.cNvPrId === selected);
    inspector.replaceChildren(
      target === undefined ? el('p', 'note', 'Click a shape to inspect it.') : describe(target),
    );
    if (mounted === null || shown === null || target === undefined || !showOverlay) return;
    const { width } = stageSizeAt(shown.zoom, deck.size);
    overlay = mountOverlay(mounted.root, target, { unit: unitAt(width, deck.size) });
  }

  // One listener on the host: the innermost `data-shape` wins, so clicking inside a
  // group selects the child. Entering and leaving groups deliberately is 5.2.
  host.addEventListener('click', (event) => {
    const target = event.target as Element | null;
    const id = target?.closest?.('[data-shape]')?.getAttribute('data-shape');
    selected = id === null || id === undefined ? null : Number(id);
    drawOverlay();
  });

  return {
    async show(index, zoom) {
      const { width, height } = stageSizeAt(zoom, deck.size);
      const ratio = window.devicePixelRatio;
      if (shown !== null && shown.index === index && shown.zoom === zoom && shown.ratio === ratio) {
        return { mountMs: 0, width, height };
      }
      const sheet = deck.slides[index];
      if (sheet === undefined) throw new RangeError(`no slide ${String(index)}`);
      overlay?.unmount();
      overlay = null;
      mounted?.unmount();
      mounted = null;
      shown = null;
      selected = null;
      const started = performance.now();
      try {
        mounted = mountSlide(host, sheet, deck.size, {
          width,
          height,
          devicePixelRatio: ratio,
          idPrefix: `slide${String(index)}`,
          media: deck.media,
          text: deck.text,
        });
      } catch (error) {
        host.replaceChildren(
          el('span', 'bad', error instanceof Error ? error.message : String(error)),
        );
        throw error;
      }
      shown = { index, zoom, ratio };
      const mountMs = performance.now() - started;
      drawOverlay();
      await nextFrame();
      return { mountMs, width, height };
    },
    svg() {
      if (mounted === null) throw new Error('the stage is empty');
      return new XMLSerializer().serializeToString(mounted.root);
    },
    shapes() {
      return placed().map((p) => ({
        id: p.shape.cNvPrId,
        name: p.shape.name,
        x: p.frame.x,
        y: p.frame.y,
        cx: p.frame.cx,
        cy: p.frame.cy,
        rot: p.frame.rot,
      }));
    },
    box() {
      if (mounted === null) throw new Error('the stage is empty');
      return pageBoxOf(mounted.root);
    },
    setOverlay(on) {
      showOverlay = on;
      drawOverlay();
    },
    unmount() {
      overlay?.unmount();
      mounted?.unmount();
      overlay = null;
      mounted = null;
      shown = null;
    },
  };
}

/**
 * The dropped deck, drawn: a strip of every slide and a stage for one, at a zoom.
 *
 * This runs on the main thread while the census runs in a Worker: a rendered slide is a DOM tree,
 * and a DOM is not transferable. Slide virtualization (12.1) is where that changes.
 */

import type { SlideSize } from '@pptx-studio/render-svg';

import { el } from '../element.js';
import { openDeck, type Deck } from './deck.js';
import { createStage, type PageBox, type ShapeFrame, type Stage } from './stage.js';
import { createStrip, type Strip, type StripDrawn } from './strip.js';
import { heapBytes, nextFrame, type OpenTimings, type ShowTimings } from './timings.js';
import { fitZoom, watchDevicePixelRatio, ZOOMS, type ZoomChoice } from './zoom.js';

const EMU_PER_POINT = 12700;

/** What a deck came to once its first slide was painted and its strip was drawn. */
export interface DeckOpened {
  readonly slides: number;
  readonly size: SlideSize;
  readonly timings: OpenTimings;
  /** Thumbnails the renderer refused, as `<slide>: <reason>`. */
  readonly failed: readonly string[];
}

export interface SlidesView {
  readonly root: HTMLElement;
  readonly ready: Promise<DeckOpened>;
  show(index: number, zoom: ZoomChoice): Promise<ShowTimings>;
  stageSvg(): string;
  thumbSvg(index: number): string;
  stageShapes(): readonly ShapeFrame[];
  stageBox(): PageBox;
  thumbBox(index: number): PageBox;
  /** Resolves when the strip's latest draw has finished. */
  stripDone(): Promise<StripDrawn>;
  /** Stop drawing: the next deck's strip must not share the thread with this one's. */
  dispose(): void;
}

/** The width the stage may fit into, which is the section's, less the frame around it. */
function viewportWidth(root: HTMLElement): number {
  const width = root.clientWidth;
  return width > 0 ? width - 2 : 960;
}

export function slidesView(read: () => Promise<ArrayBuffer>): SlidesView {
  const root = el('div', 'slides');
  const status = el('p', 'note', 'Reading the deck…');
  const timing = el('p', 'note timings');
  const strip = el('div', 'strip');
  const controls = el('div', 'stage-controls');
  const stageHost = el('div', 'stage');
  const inspector = el('div', 'shape-inspector');

  const zoomSelect = el('select', 'zoom');
  for (const zoom of ZOOMS) {
    const option = el('option', '', `${String(zoom * 100)} %`);
    option.value = String(zoom);
    zoomSelect.append(option);
  }
  const fit = el('option', '', 'fit');
  fit.value = 'fit';
  zoomSelect.append(fit);
  zoomSelect.value = 'fit';
  const toggle = el('button', 'plain', 'Debug overlay: on');
  toggle.setAttribute('type', 'button');
  controls.append(el('label', '', 'Zoom '), zoomSelect, toggle);
  root.append(status, timing, strip, controls, stageHost, inspector);

  let deck: Deck | null = null;
  let stage: Stage | null = null;
  let stripView: Strip | null = null;
  let current = 0;
  let choice: ZoomChoice = 'fit';
  let disposed = false;

  const zoomOf = (wanted: ZoomChoice): number =>
    wanted === 'fit' ? fitZoom(viewportWidth(root), deck?.size ?? { cx: 1, cy: 1 }) : wanted;

  async function show(index: number, wanted: ZoomChoice): Promise<ShowTimings> {
    if (deck === null || stage === null || stripView === null) throw new Error('no deck is open');
    current = index;
    choice = wanted;
    zoomSelect.value = wanted === 'fit' ? 'fit' : String(wanted);
    stripView.select(index);
    return stage.show(index, zoomOf(wanted));
  }

  zoomSelect.addEventListener('change', () => {
    const value = zoomSelect.value;
    void show(current, value === 'fit' ? 'fit' : Number(value)).catch(() => undefined);
  });
  toggle.addEventListener('click', () => {
    const on = toggle.textContent.endsWith('off');
    toggle.textContent = `Debug overlay: ${on ? 'on' : 'off'}`;
    stage?.setOverlay(on);
  });
  window.addEventListener('resize', () => {
    if (choice === 'fit' && stage !== null) void show(current, 'fit').catch(() => undefined);
  });
  // A window dragged to another display: the strokes are rounded to its pixels, so redraw.
  const unwatch = watchDevicePixelRatio(window, () => {
    if (disposed || stage === null) return;
    void show(current, choice).catch(() => undefined);
    void stripView?.redraw();
  });

  const ready = (async (): Promise<DeckOpened> => {
    const started = performance.now();
    const bytes = new Uint8Array(await read());
    const readMs = performance.now() - started;
    const opened = openDeck(bytes);
    const parseMs = performance.now() - started - readMs;
    if (opened.slides.length === 0) {
      status.textContent = 'This package has no slides.';
      throw new Error('This package has no slides.');
    }
    if (disposed) throw new Error('disposed');
    deck = opened;
    status.textContent =
      `${String(opened.slides.length)} slide${opened.slides.length === 1 ? '' : 's'} · ` +
      `${(opened.size.cx / EMU_PER_POINT).toFixed(0)} × ` +
      `${(opened.size.cy / EMU_PER_POINT).toFixed(0)} pt` +
      (opened.problems.length === 0
        ? ''
        : ` · ${String(opened.problems.length)} model problem(s), carried across unchanged`);

    // The stage first, so the first slide is painted before the strip starts on the rest.
    stage = createStage(stageHost, inspector, opened);
    stripView = createStrip(strip, opened, (index) => {
      void show(index, choice).catch(() => undefined);
    });
    stripView.select(0);
    await stage.show(0, zoomOf('fit'));
    await nextFrame();
    const firstStageMs = performance.now() - started;
    const { failed, cpuMs } = await stripView.done;
    const stripMs = performance.now() - started;
    const timings: OpenTimings = {
      readMs,
      parseMs,
      firstStageMs,
      stripMs,
      stripCpuMs: cpuMs,
      heapBytes: heapBytes(),
    };
    timing.textContent =
      `parsed in ${parseMs.toFixed(0)} ms · first slide at ${firstStageMs.toFixed(0)} ms · ` +
      `${String(opened.slides.length)} thumbnails at ${stripMs.toFixed(0)} ms` +
      (failed.length === 0 ? '' : ` · ${String(failed.length)} not drawn`);
    return { slides: opened.slides.length, size: opened.size, timings, failed };
  })();
  ready.catch((error: unknown) => {
    status.textContent =
      'The renderer could not read this deck: ' +
      (error instanceof Error ? error.message : String(error));
  });

  return {
    root,
    ready,
    show,
    stageSvg: () => {
      if (stage === null) throw new Error('no deck is open');
      return stage.svg();
    },
    thumbSvg: (index) => {
      if (stripView === null) throw new Error('no deck is open');
      return stripView.svg(index);
    },
    stageShapes: () => {
      if (stage === null) throw new Error('no deck is open');
      return stage.shapes();
    },
    stageBox: () => {
      if (stage === null) throw new Error('no deck is open');
      return stage.box();
    },
    thumbBox: (index) => {
      if (stripView === null) throw new Error('no deck is open');
      return stripView.box(index);
    },
    stripDone: () => {
      if (stripView === null) throw new Error('no deck is open');
      return stripView.done;
    },
    dispose: () => {
      disposed = true;
      unwatch();
      stripView?.cancel();
      stage?.unmount();
    },
  };
}

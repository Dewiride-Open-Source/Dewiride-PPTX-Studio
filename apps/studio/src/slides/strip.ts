/**
 * The strip: every slide as a 120-pixel thumbnail, mounted in frame-sized chunks after the
 * stage has painted, so the first slide is on screen before the hundredth is drawn.
 *
 * Every slide still mounts; nothing is virtualised (that is 12.1). What a chunked loop buys
 * is a page that answers clicks while it draws, and a strip the next deck can cancel.
 */

import { mountSlide, type MountedSlide } from '@pptx-studio/render-dom';

import { el } from '../element.js';
import type { Deck } from './deck.js';
import { pageBoxOf, type PageBox } from './stage.js';
import { nextFrame } from './timings.js';
import { stageSizeAt, THUMB_ZOOM } from './zoom.js';

/** Main-thread time one frame's chunk may take before the loop yields. */
const CHUNK_BUDGET_MS = 12;

export interface StripDrawn {
  /** Thumbnails the renderer refused, as `<slide>: <reason>`. */
  readonly failed: readonly string[];
  /** Main-thread time the mounts took, across every chunk. */
  readonly cpuMs: number;
}

export interface Strip {
  /** Resolves when the last thumbnail of the latest draw has mounted. */
  readonly done: Promise<StripDrawn>;
  /** Draw every thumbnail again, at the display's current device pixel ratio. */
  redraw(): Promise<StripDrawn>;
  select(index: number): void;
  svg(index: number): string;
  box(index: number): PageBox;
  cancel(): void;
}

export function createStrip(
  host: HTMLElement,
  deck: Deck,
  onSelect: (index: number) => void,
): Strip {
  host.dataset['role'] = 'strip';
  const { width, height } = stageSizeAt(THUMB_ZOOM, deck.size);
  const holders: HTMLElement[] = [];
  const cells: HTMLElement[] = [];
  const roots = new Map<number, SVGSVGElement>();
  const mounts = new Map<number, MountedSlide>();
  // A draw stops as soon as it is not the latest one asked for.
  let generation = 0;

  host.replaceChildren();
  deck.slides.forEach((_sheet, index) => {
    const cell = el('button', 'thumb');
    cell.setAttribute('type', 'button');
    cell.dataset['thumb'] = String(index);
    const holder = el('div', 'thumb-svg');
    holder.style.width = `${String(width)}px`;
    holder.style.height = `${String(height)}px`;
    cell.append(holder, el('span', 'n', String(index + 1)));
    host.append(cell);
    holders.push(holder);
    cells.push(cell);
  });

  host.addEventListener('click', (event) => {
    const target = event.target as Element | null;
    const index = target?.closest?.('[data-thumb]')?.getAttribute('data-thumb');
    if (index !== null && index !== undefined) onSelect(Number(index));
  });

  async function draw(mine: number): Promise<StripDrawn> {
    // The first chunk waits a frame, so the stage mounted beside this strip paints first.
    await nextFrame();
    const ratio = window.devicePixelRatio;
    const failed: string[] = [];
    let cpuMs = 0;
    let index = 0;
    while (index < deck.slides.length && mine === generation) {
      const started = performance.now();
      while (index < deck.slides.length && performance.now() - started < CHUNK_BUDGET_MS) {
        const sheet = deck.slides[index]!;
        const holder = holders[index]!;
        try {
          const mounted = mountSlide(holder, sheet, deck.size, {
            width,
            height,
            devicePixelRatio: ratio,
            idPrefix: `thumb${String(index)}`,
            media: deck.media,
            text: deck.text,
          });
          roots.set(index, mounted.root);
          mounts.set(index, mounted);
        } catch (error) {
          // One slide the renderer cannot draw must not take the deck with it.
          const message = error instanceof Error ? error.message : String(error);
          holder.append(el('span', 'bad', message));
          failed.push(`${String(index + 1)}: ${message}`);
        }
        index += 1;
      }
      cpuMs += performance.now() - started;
      if (index < deck.slides.length) await nextFrame();
    }
    return { failed, cpuMs };
  }

  function clear(): void {
    for (const mounted of mounts.values()) mounted.unmount();
    mounts.clear();
    roots.clear();
    for (const holder of holders) holder.replaceChildren();
  }

  let done = draw(generation);

  return {
    get done() {
      return done;
    },
    redraw() {
      generation += 1;
      clear();
      done = draw(generation);
      return done;
    },
    select(index) {
      cells.forEach((cell, at) => cell.classList.toggle('on', at === index));
    },
    svg(index) {
      const root = roots.get(index);
      if (root === undefined) throw new Error(`thumbnail ${String(index)} is not mounted`);
      return new XMLSerializer().serializeToString(root);
    },
    box(index) {
      const root = roots.get(index);
      if (root === undefined) throw new Error(`thumbnail ${String(index)} is not mounted`);
      return pageBoxOf(root);
    },
    cancel() {
      generation += 1;
    },
  };
}

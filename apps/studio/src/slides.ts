/**
 * The dropped deck, drawn.
 *
 * Geometry, fills, gradients, patterns, strokes, effects and nested rotated
 * groups, with every shape selectable and the 2.11 debug overlay on whichever
 * one is selected, and the text drawn over it as real `<text>`.
 *
 * ## Why this runs on the main thread while the census runs in a Worker
 *
 * The census answers a question about bytes and hands back plain JSON, so it
 * belongs off the main thread. A rendered slide is a DOM tree, and a DOM is not
 * transferable. Moving the parse into the Worker would mean serialising the
 * whole `Sheet` chain - every `XNode`, with the raw byte offsets that make the
 * round trip work - across `postMessage` for every slide. Slide virtualization
 * (12.1) and an `OffscreenCanvas` raster path are where that changes.
 */

import { PartStore } from '@pptx-studio/opc';
import { loadDocument, type ListStyle, type Sheet } from '@pptx-studio/model';
import { mountOverlay, mountSlide, type MountedOverlay } from '@pptx-studio/render-dom';
import { flatten, type Placed, type SlideSize } from '@pptx-studio/render-svg';

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className = '',
  text = '',
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== '') node.className = className;
  if (text !== '') node.textContent = text;
  return node;
}

const EMU_PER_POINT = 12700;

interface Deck {
  readonly slides: readonly Sheet[];
  readonly size: SlideSize;
  /** The text cascade's seventh source, which lives on `ppt/presentation.xml`. */
  readonly defaultTextStyle: ListStyle | undefined;
  readonly problems: readonly { message: string }[];
}

/** The stage width in CSS pixels. Fixed, because zoom is a transform, never a re-layout. */
const STAGE_PX = 900;

interface State {
  deck: Deck;
  slide: number;
  selected: number | null;
  overlay: MountedOverlay | null;
  showOverlay: boolean;
}

function describe(placed: Placed): HTMLElement {
  const rows = el('div', 'rows');
  const add = (key: string, value: string): void => {
    const row = el('div', 'row');
    row.append(el('span', 'k', key), el('span', 'v', value));
    rows.append(row);
  };
  const frame = placed.frame;
  const pt = (value: number): string => (value / EMU_PER_POINT).toFixed(1);

  add('name', placed.shape.name === '' ? `#${String(placed.shape.cNvPrId)}` : placed.shape.name);
  add('kind', placed.shape.kind);
  add('position', `${pt(frame.x)}, ${pt(frame.y)} pt`);
  add('size', `${pt(frame.cx)} × ${pt(frame.cy)} pt`);
  if (frame.rot !== 0) add('rotation', `${frame.rot.toFixed(2)}°`);
  if (frame.flipH || frame.flipV) {
    add('mirrored', [frame.flipH ? 'H' : '', frame.flipV ? 'V' : ''].filter(Boolean).join(' + '));
  }
  const geometry = placed.geometrySource?.geometry;
  add('geometry', geometry === undefined ? 'none' : (geometry.name ?? 'a:custGeom'));
  if (placed.geometry !== null) {
    const broken = placed.geometry.paths.filter((p) => !p.finite).length;
    add(
      'paths',
      String(placed.geometry.paths.length) + (broken === 0 ? '' : ` (${String(broken)} broken)`),
    );
  }
  add('fill', placed.fill === null ? 'none' : placed.fill.type);
  add('sheet', placed.sheet.kind);
  return rows;
}

export function slidesView(read: () => Promise<ArrayBuffer>): HTMLElement {
  const root = el('div', 'slides');
  const strip = el('div', 'strip');
  const stage = el('div', 'stage');
  const inspector = el('div', 'shape-inspector');
  const status = el('p', 'note', 'Reading the deck…');
  const controls = el('div', 'stage-controls');

  const toggle = el('button', 'plain', 'Debug overlay: on');
  toggle.setAttribute('type', 'button');
  controls.append(toggle);

  root.append(status, strip, controls, stage, inspector);

  let state: State | null = null;

  function drawStage(): void {
    // Bound to a local, because the narrowing does not survive into a closure
    // and every callback below needs it.
    const current = state;
    if (current === null) return;
    const sheet = current.deck.slides[current.slide];
    if (sheet === undefined) return;
    const height = Math.round((STAGE_PX * current.deck.size.cy) / current.deck.size.cx);
    const mounted = mountSlide(stage, sheet, current.deck.size, {
      width: STAGE_PX,
      height,
      idPrefix: `slide${String(current.slide)}`,
      text: { defaultTextStyle: current.deck.defaultTextStyle },
    });
    const all = flatten(mounted.placed);

    current.overlay?.unmount();
    current.overlay = null;
    const selected = all.find((p) => p.shape.cNvPrId === current.selected);
    if (selected !== undefined && current.showOverlay) {
      current.overlay = mountOverlay(mounted.root, selected, {
        unit: current.deck.size.cx / STAGE_PX,
      });
    }
    inspector.replaceChildren(
      selected === undefined ? el('p', 'note', 'Click a shape to inspect it.') : describe(selected),
    );

    mounted.root.addEventListener('click', (event) => {
      const target = event.target as Element | null;
      // The innermost `data-shape` wins, so clicking inside a group selects the
      // child rather than the group. Entering and leaving groups deliberately
      // is sub-phase 5.2.
      const id = target?.closest?.('[data-shape]')?.getAttribute('data-shape');
      current.selected = id === null || id === undefined ? null : Number(id);
      drawStage();
    });
  }

  function drawStrip(): void {
    const current = state;
    if (current === null) return;
    strip.replaceChildren();
    current.deck.slides.forEach((sheet, index) => {
      const cell = el('button', `thumb${index === current.slide ? ' on' : ''}`);
      cell.setAttribute('type', 'button');
      const holder = el('div', 'thumb-svg');
      try {
        mountSlide(holder, sheet, current.deck.size, {
          width: 132,
          height: Math.round((132 * current.deck.size.cy) / current.deck.size.cx),
          idPrefix: `thumb${String(index)}`,
          text: { defaultTextStyle: current.deck.defaultTextStyle },
        });
      } catch (error) {
        // One slide the renderer cannot draw must not take the deck with it.
        holder.append(el('span', 'bad', error instanceof Error ? error.message : String(error)));
      }
      cell.append(holder, el('span', 'n', String(index + 1)));
      cell.addEventListener('click', () => {
        current.slide = index;
        current.selected = null;
        drawStrip();
        drawStage();
      });
      strip.append(cell);
    });
  }

  toggle.addEventListener('click', () => {
    if (state === null) return;
    state.showOverlay = !state.showOverlay;
    toggle.textContent = `Debug overlay: ${state.showOverlay ? 'on' : 'off'}`;
    drawStage();
  });

  void (async () => {
    try {
      const document_ = loadDocument(PartStore.open(new Uint8Array(await read())));
      const deck: Deck = {
        slides: document_.slides,
        size: document_.slideSize,
        defaultTextStyle: document_.defaultTextStyle,
        problems: document_.problems,
      };
      if (deck.slides.length === 0) {
        status.textContent = 'This package has no slides.';
        return;
      }
      state = { deck, slide: 0, selected: null, overlay: null, showOverlay: true };
      status.textContent =
        `${String(deck.slides.length)} slide${deck.slides.length === 1 ? '' : 's'} · ` +
        `${(deck.size.cx / EMU_PER_POINT).toFixed(0)} × ` +
        `${(deck.size.cy / EMU_PER_POINT).toFixed(0)} pt` +
        (deck.problems.length === 0
          ? ''
          : ` · ${String(deck.problems.length)} model problem(s), carried across unchanged`);
      drawStrip();
      drawStage();
    } catch (error) {
      status.textContent =
        'The renderer could not read this deck: ' +
        (error instanceof Error ? error.message : String(error));
    }
  })();

  return root;
}

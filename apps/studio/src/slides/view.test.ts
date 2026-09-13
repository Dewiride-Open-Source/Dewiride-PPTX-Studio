import { describe, expect, it } from 'vitest';

import { el } from '../element.js';
import { openDeck } from './deck.js';
import { createStage } from './stage.js';
import { createStrip } from './strip.js';
import { nextFrame } from './timings.js';
import { slidesView } from './view.js';

/**
 * The slides view against the committed corpus, over Vitest's own server: the one
 * place a package test cannot go, because browser mode has no filesystem.
 */

const MINIMAL = '/corpus/decks/a01-minimal.pptx';
const FILLS = '/corpus/decks/a03-fills.pptx';
const LINES = '/corpus/decks/a06-lines.pptx';

async function bytesOf(path: string): Promise<ArrayBuffer> {
  const response = await fetch(path);
  if (!response.ok) throw new Error(`could not fetch ${path}: ${String(response.status)}`);
  return response.arrayBuffer();
}

function host(): HTMLElement {
  const node = document.createElement('div');
  node.style.width = '1000px';
  document.body.append(node);
  return node;
}

describe('slidesView', () => {
  it('paints the first slide, then every thumbnail, and says how long each took', async () => {
    const view = slidesView(() => bytesOf(FILLS));
    host().append(view.root);
    const opened = await view.ready;
    expect(opened.slides).toBe(3);
    expect(opened.failed).toEqual([]);
    expect(opened.timings.firstStageMs).toBeGreaterThan(0);
    expect(opened.timings.stripMs).toBeGreaterThanOrEqual(opened.timings.firstStageMs);
    expect(view.root.querySelectorAll('.thumb-svg > svg')).toHaveLength(3);
    // A thumbnail is exactly the oracle's smallest export.
    const thumb = view.root.querySelector('.thumb-svg > svg');
    expect(thumb?.getAttribute('width')).toBe('120');
    expect(thumb?.getAttribute('height')).toBe('68');
    view.dispose();
  });

  it('mounts the stage again at a new zoom, and not at all for the zoom it shows', async () => {
    const view = slidesView(() => bytesOf(MINIMAL));
    host().append(view.root);
    await view.ready;
    const first = await view.show(0, 1);
    expect(first.width).toBe(960);
    const before = view.root.querySelector('.stage > svg');
    const again = await view.show(0, 1);
    expect(again.mountMs).toBe(0);
    expect(view.root.querySelector('.stage > svg')).toBe(before);
    const doubled = await view.show(0, 2);
    expect(doubled.width).toBe(1920);
    expect(doubled.mountMs).toBeGreaterThan(0);
    const after = view.root.querySelector('.stage > svg');
    expect(after).not.toBe(before);
    expect(after?.getAttribute('width')).toBe('1920');
    view.dispose();
  });

  it('draws a picture fill on the stage', async () => {
    const view = slidesView(() => bytesOf(FILLS));
    host().append(view.root);
    await view.ready;
    await view.show(1, 1);
    expect(view.stageSvg()).toContain('data:image/png;base64');
    view.dispose();
  });

  it('serialises the stage and reports its shapes and its page box', async () => {
    const view = slidesView(() => bytesOf(MINIMAL));
    host().append(view.root);
    await view.ready;
    await view.show(0, 0.25);
    expect(view.stageSvg().startsWith('<svg')).toBe(true);
    expect(view.stageShapes().length).toBeGreaterThan(0);
    const box = view.stageBox();
    expect(box.width).toBe(240);
    expect(box.height).toBe(135);
    expect(view.thumbBox(0).width).toBe(120);
    expect(view.thumbSvg(0)).toContain('viewBox');
    view.dispose();
  });

  it('draws the strip again on request and reports the latest draw', async () => {
    const view = slidesView(() => bytesOf(FILLS));
    host().append(view.root);
    await view.ready;
    const before = [...view.root.querySelectorAll('.thumb-svg > svg')];
    expect(before).toHaveLength(3);
    const deck = openDeck(new Uint8Array(await bytesOf(FILLS)));
    const strip = createStrip(el('div', 'strip'), deck, () => undefined);
    await strip.done;
    const first = strip.svg(0);
    const again = await strip.redraw();
    expect(again.failed).toEqual([]);
    expect(await strip.done).toBe(again);
    expect(strip.svg(0)).toBe(first);
    view.dispose();
  });

  it('mounts no thumbnail until the frame after the one the stage paints in', async () => {
    const deck = openDeck(new Uint8Array(await bytesOf(FILLS)));
    const host = el('div', 'strip');
    const strip = createStrip(host, deck, () => undefined);
    expect(host.querySelectorAll('svg')).toHaveLength(0);
    await nextFrame();
    expect(host.querySelectorAll('svg')).toHaveLength(0);
    await strip.done;
    expect(host.querySelectorAll('svg')).toHaveLength(3);
  });

  it('stops a running draw when cancelled, and mounts nothing more', async () => {
    const deck = openDeck(new Uint8Array(await bytesOf(FILLS)));
    const host = el('div', 'strip');
    const strip = createStrip(host, deck, () => undefined);
    strip.cancel();
    const drawn = await strip.done;
    expect(drawn.failed).toEqual([]);
    expect(host.querySelectorAll('svg')).toHaveLength(0);
    // A redraw after a cancel is a fresh draw and completes.
    await strip.redraw();
    expect(host.querySelectorAll('svg')).toHaveLength(3);
  });

  it('mounts the stage again when the display ratio changes, at the same zoom', async () => {
    // a06's second slide: 2.25-pt lines, two pixels at 1 px/pt and five at 2 px/pt.
    const deck = openDeck(new Uint8Array(await bytesOf(LINES)));
    const stageHost = el('div', 'stage');
    const stage = createStage(stageHost, el('div', 'inspector'), deck);
    const was = window.devicePixelRatio;
    try {
      await stage.show(1, 1);
      const one = stage.svg();
      expect((await stage.show(1, 1)).mountMs).toBe(0);
      Object.defineProperty(window, 'devicePixelRatio', { value: was * 2, configurable: true });
      expect((await stage.show(1, 1)).mountMs).toBeGreaterThan(0);
      // The same slide at the same CSS size, with its strokes rounded to twice the pixels.
      const widthOf = (svg: string): string =>
        /<g data-shape.*?stroke-width="([^"]+)"/.exec(svg)?.[1] ?? '';
      expect(widthOf(one)).toBe('25400');
      expect(widthOf(stage.svg())).toBe('31750');
      expect(stage.svg()).toContain('width="960"');
    } finally {
      Object.defineProperty(window, 'devicePixelRatio', { value: was, configurable: true });
      stage.unmount();
    }
  });

  it('stops drawing the strip when disposed', async () => {
    const view = slidesView(() => bytesOf(FILLS));
    host().append(view.root);
    view.dispose();
    await expect(view.ready).rejects.toThrow();
    expect(view.root.querySelectorAll('.thumb-svg > svg')).toHaveLength(0);
  });
});

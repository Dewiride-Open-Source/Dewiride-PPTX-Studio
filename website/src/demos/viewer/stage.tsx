'use client';

import { useEffect, useMemo, useRef, type KeyboardEvent } from 'react';

import type { Sheet } from '@pptx-studio/model';
import { mountOverlay, mountSlide, mountTextLayer, type LayerBlock } from '@pptx-studio/render-dom';
import {
  flatten,
  layoutSlide,
  textBlockOf,
  type Placed,
  type SlideSize,
} from '@pptx-studio/render-svg';

import type { ParsedDeck } from '@/deck/document';
import { deviceRatioAt, pointsOf, stageSizeAt, unitAt } from './zoom';

const EMU_PER_POINT = 12700;

interface StageProps {
  readonly deck: ParsedDeck;
  readonly sheet: Sheet;
  readonly size: SlideSize;
  /** CSS pixels per point. */
  readonly zoom: number;
  readonly devicePixelRatio: number;
  readonly overlay: boolean;
  readonly textLayer: boolean;
  readonly selected: number | null;
  readonly onSelect: (id: number | null) => void;
  readonly onMounted?: (placed: readonly Placed[], ms: number) => void;
}

/**
 * One slide, mounted live.
 *
 * A zoom is a fresh mount at that width, not a resize: strokes are rounded to
 * the device pixels of the mounted width, so the SVG for 25 % is not the SVG
 * for 400 %. The text layout is the same at every zoom.
 */
export function Stage({
  deck,
  sheet,
  size,
  zoom,
  devicePixelRatio,
  overlay,
  textLayer,
  selected,
  onSelect,
  onMounted,
}: StageProps) {
  const host = useRef<HTMLDivElement>(null);
  const textHost = useRef<HTMLDivElement>(null);
  // The same layout mountSlide makes, so the overlay and the text layer never wait on the mount.
  const placed = useMemo<readonly Placed[]>(() => flatten(layoutSlide(sheet)), [sheet]);
  const mountedRoot = useRef<SVGSVGElement | null>(null);
  const { width, height } = stageSizeAt(zoom, size);

  useEffect(() => {
    const element = host.current;
    if (element === null) return;
    const startedAt = performance.now();
    const mounted = mountSlide(element, sheet, size, {
      width,
      devicePixelRatio: deviceRatioAt(zoom, devicePixelRatio),
      idPrefix: `v${String(width)}`,
      media: deck.media,
      text: deck.text,
    });
    mountedRoot.current = mounted.root;
    onMounted?.(flatten(mounted.placed), performance.now() - startedAt);
    return () => {
      mounted.unmount();
      mountedRoot.current = null;
    };
  }, [deck, sheet, size, zoom, width, devicePixelRatio, onMounted]);

  useEffect(() => {
    const root = mountedRoot.current;
    if (root === null || !overlay || selected === null) return;
    const target = placed.find((one) => one.shape.cNvPrId === selected);
    if (target === undefined) return;
    const chrome = mountOverlay(root, target, { unit: unitAt(width, size) });
    return () => chrome.unmount();
  }, [placed, overlay, selected, width, size]);

  useEffect(() => {
    const element = textHost.current;
    if (element === null || !textLayer) return;
    const blocks: LayerBlock[] = [];
    for (const one of placed) {
      const block = textBlockOf(one, deck.text);
      if (block === null) continue;
      blocks.push({
        block,
        leftPt: one.frame.x / EMU_PER_POINT,
        topPt: one.frame.y / EMU_PER_POINT,
        widthPt: one.frame.cx / EMU_PER_POINT,
        heightPt: one.frame.cy / EMU_PER_POINT,
        cNvPrId: one.shape.cNvPrId,
        name: one.shape.name,
      });
    }
    const layer = mountTextLayer(element, blocks, pointsOf(size), {
      scale: zoom,
      interactive: true,
    });
    return () => layer.unmount();
  }, [placed, textLayer, deck, size, zoom]);

  const pick = (target: EventTarget | null): number | null => {
    const hit = (target as Element | null)?.closest?.('[data-shape], [data-text]');
    const id = hit?.getAttribute('data-shape') ?? hit?.getAttribute('data-text');
    return id === null || id === undefined ? null : Number(id);
  };

  const keys = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') onSelect(null);
    if (placed.length === 0 || (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft')) return;
    event.preventDefault();
    const at = placed.findIndex((one) => one.shape.cNvPrId === selected);
    const next = event.key === 'ArrowRight' ? at + 1 : at - 1;
    const wrapped = ((next % placed.length) + placed.length) % placed.length;
    onSelect(placed[wrapped]?.shape.cNvPrId ?? null);
  };

  return (
    <div
      role="application"
      aria-label={
        selected === null
          ? 'Slide: click a shape, or use the arrow keys'
          : `Slide, shape ${String(selected)} selected`
      }
      tabIndex={0}
      className="stage-surface relative shrink-0 overflow-hidden"
      style={{ width, height }}
      onKeyDown={keys}
      onPointerDown={(event) => onSelect(pick(event.target))}
    >
      <div ref={host} className="absolute inset-0 [&>svg]:block [&>svg]:h-full [&>svg]:w-full" />
      <div
        ref={textHost}
        className={`absolute inset-0 ${textLayer ? '' : 'pointer-events-none'}`}
      />
    </div>
  );
}

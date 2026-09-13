'use client';

import { useEffect, useRef } from 'react';

import { mountSlide } from '@pptx-studio/render-dom';
import type { SlideSize } from '@pptx-studio/render-svg';

import type { ParsedDeck } from '@/deck/document';
import { THUMB_ZOOM, deviceRatioAt, stageSizeAt } from './zoom';

/** How long one frame may spend mounting thumbnails before yielding to the page. */
const FRAME_BUDGET_MS = 12;

interface StripProps {
  readonly deck: ParsedDeck;
  readonly size: SlideSize;
  readonly current: number;
  readonly onSelect: (index: number) => void;
  readonly horizontal?: boolean;
}

/**
 * Every slide as a thumbnail, mounted in frame-budgeted chunks after the stage has painted.
 *
 * A hundred thumbnails mounted at once would hold the main thread for the whole
 * mount; twelve milliseconds a frame keeps the page responsive while they fill
 * in, and a slide scrolled out of view is not mounted until it comes near.
 */
export function Strip({ deck, size, current, onSelect, horizontal = false }: StripProps) {
  const hosts = useRef<(HTMLDivElement | null)[]>([]);
  const { width, height } = stageSizeAt(THUMB_ZOOM, size);
  const slides = deck.document.slides;

  useEffect(() => {
    const mounted = new Map<number, () => void>();
    const queue: number[] = [];
    let frame = 0;

    const pump = (): void => {
      frame = 0;
      const startedAt = performance.now();
      while (queue.length > 0 && performance.now() - startedAt < FRAME_BUDGET_MS) {
        const index = queue.shift();
        if (index === undefined) break;
        const host = hosts.current[index];
        const sheet = slides[index];
        if (host === null || host === undefined || sheet === undefined || mounted.has(index))
          continue;
        try {
          const thumb = mountSlide(host, sheet, size, {
            width,
            devicePixelRatio: deviceRatioAt(THUMB_ZOOM, window.devicePixelRatio),
            idPrefix: `t${String(index)}`,
            media: deck.media,
            text: deck.text,
          });
          mounted.set(index, () => thumb.unmount());
        } catch (cause) {
          host.textContent = cause instanceof Error ? cause.message : String(cause);
          host.classList.add('text-bad', 'text-[9px]', 'p-1');
          mounted.set(index, () => {
            host.textContent = '';
          });
        }
      }
      if (queue.length > 0) frame = requestAnimationFrame(pump);
    };

    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const index = Number((entry.target as HTMLElement).dataset['index']);
          if (!mounted.has(index) && !queue.includes(index)) queue.push(index);
        }
        if (queue.length > 0 && frame === 0) frame = requestAnimationFrame(pump);
      },
      { rootMargin: '200px' },
    );
    for (const host of hosts.current) if (host !== null) observer.observe(host);

    return () => {
      observer.disconnect();
      if (frame !== 0) cancelAnimationFrame(frame);
      for (const unmount of mounted.values()) unmount();
    };
  }, [deck, slides, size, width]);

  return (
    <ol
      aria-label="Slides"
      className={`flex gap-2 ${horizontal ? 'overflow-x-auto pb-2' : 'max-h-[70vh] flex-col overflow-y-auto pr-1'}`}
    >
      {slides.map((sheet, index) => (
        <li key={sheet.partName} className="shrink-0">
          <button
            type="button"
            onClick={() => onSelect(index)}
            aria-label={`Slide ${String(index + 1)}`}
            aria-current={index === current ? 'true' : undefined}
            className={`flex items-start gap-2 rounded-control border p-1 text-left transition-colors ${
              index === current
                ? 'border-accent bg-accent-soft'
                : 'border-transparent hover:border-line-strong'
            }`}
          >
            <span className="w-5 pt-0.5 text-right font-mono text-[10px] text-fg-faint">
              {index + 1}
            </span>
            <div
              ref={(element) => {
                hosts.current[index] = element;
              }}
              data-index={index}
              className="stage-surface overflow-hidden [&>svg]:block [&>svg]:h-full [&>svg]:w-full"
              style={{ width, height }}
            />
          </button>
        </li>
      ))}
    </ol>
  );
}

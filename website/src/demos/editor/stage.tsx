'use client';

import { useMemo, type KeyboardEvent, type PointerEvent as ReactPointerEvent } from 'react';

import type { Sheet } from '@pptx-studio/model';
import {
  flatten,
  serializeSvg,
  shapeOverlay,
  slideNode,
  type MediaResolver,
  type Placed,
  type SlideSize,
  type TextEngine,
} from '@pptx-studio/render-svg';

export interface StageContent {
  readonly svg: string;
  readonly placed: readonly Placed[];
  readonly error: string | null;
}

const EMPTY: StageContent = { svg: '', placed: [], error: null };

/** The sheet drawn once as a string, with every shape flattened out for hit-testing. */
export function useStage(
  sheet: Sheet | undefined,
  size: SlideSize | undefined,
  text: TextEngine | false | null,
  idPrefix: string,
  media: MediaResolver | undefined,
): StageContent {
  return useMemo(() => {
    if (sheet === undefined || size === undefined || text === null || media === undefined)
      return EMPTY;
    try {
      const { node, placed } = slideNode(sheet, size, { idPrefix, media, text });
      return { svg: serializeSvg(node), placed: flatten(placed), error: null };
    } catch (cause) {
      return { svg: '', placed: [], error: cause instanceof Error ? cause.message : String(cause) };
    }
  }, [sheet, size, text, idPrefix, media]);
}

export interface StageProps {
  readonly content: StageContent;
  readonly size: SlideSize;
  readonly width: number;
  readonly selected?: number | null;
  readonly overlay?: boolean;
  readonly onSelect?: (id: number | null) => void;
  readonly onShapePointerDown?: (placed: Placed, event: ReactPointerEvent) => void;
  /** Arrow keys on the selected shape, in EMU; Shift moves ten times as far. */
  readonly onNudge?: (id: number, dx: number, dy: number) => void;
}

const NUDGE_EMU = 12700;

/**
 * One slide on screen.
 *
 * Zoom is a CSS width on a fixed EMU viewBox and never a re-layout - font
 * metrics are not linear in point size, so laying out again at a new zoom would
 * move the line breaks.
 */
export function Stage({
  content,
  size,
  width,
  selected = null,
  overlay = false,
  onSelect,
  onShapePointerDown,
  onNudge,
}: StageProps) {
  const { svg, placed, error } = content;
  const height = Math.round((width * size.cy) / size.cx);

  const chrome = useMemo(() => {
    if (!overlay || selected === null) return null;
    const found = placed.find((one) => one.shape.cNvPrId === selected);
    if (found === undefined) return null;
    try {
      return serializeSvg(shapeOverlay(found, { unit: size.cx / width }).node);
    } catch {
      // A shape whose geometry will not resolve has no overlay, which is not a
      // reason to stop drawing the slide.
      return null;
    }
  }, [overlay, selected, placed, size.cx, width]);

  if (error !== null) {
    return (
      <div
        style={{ width, height }}
        className="flex items-center justify-center rounded-control border border-bad/40 bg-bad/5 p-4 text-center font-mono text-[12px] text-bad"
      >
        {error}
      </div>
    );
  }

  const pick = (event: ReactPointerEvent<HTMLDivElement>): number | null => {
    const target = event.target as Element | null;
    // Text is drawn in a group that is a *sibling* of the shape's geometry, not
    // a child of it, so a click on a word never reaches `[data-shape]`. Both
    // carry the same `cNvPrId`.
    const hit = target?.closest?.('[data-shape], [data-text]');
    const id = hit?.getAttribute('data-shape') ?? hit?.getAttribute('data-text');
    return id === null || id === undefined ? null : Number(id);
  };

  const keys = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      onSelect?.(null);
      return;
    }
    if (selected === null || onNudge === undefined) return;
    const step = NUDGE_EMU * (event.shiftKey ? 10 : 1);
    const delta: Record<string, [number, number]> = {
      ArrowLeft: [-step, 0],
      ArrowRight: [step, 0],
      ArrowUp: [0, -step],
      ArrowDown: [0, step],
    };
    const found = delta[event.key];
    if (found === undefined) return;
    event.preventDefault();
    onNudge(selected, found[0], found[1]);
  };

  return (
    <div
      role="application"
      aria-label={
        selected === null
          ? 'Slide: click a shape to select it'
          : `Slide, shape ${String(selected)} selected; arrow keys move it`
      }
      tabIndex={0}
      className="stage-surface relative shrink-0 overflow-hidden"
      style={{ width, height }}
      onKeyDown={keys}
      onPointerDown={(event) => {
        const id = pick(event);
        onSelect?.(id);
        if (id === null) return;
        const found = placed.find((one) => one.shape.cNvPrId === id);
        if (found !== undefined) onShapePointerDown?.(found, event);
      }}
    >
      <div
        className="absolute inset-0 [&>svg]:h-full [&>svg]:w-full"
        dangerouslySetInnerHTML={{ __html: svg }}
      />
      {chrome === null ? null : (
        <svg
          viewBox={`0 0 ${String(size.cx)} ${String(size.cy)}`}
          className="pointer-events-none absolute inset-0 h-full w-full"
          dangerouslySetInnerHTML={{ __html: chrome }}
        />
      )}
    </div>
  );
}

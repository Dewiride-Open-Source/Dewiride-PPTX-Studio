'use client';

import { useMemo, type PointerEvent as ReactPointerEvent } from 'react';

import type { ListStyle, Sheet } from '@pptx-studio/model';
import {
  flatten,
  serializeSvg,
  shapeOverlay,
  slideNode,
  type MediaResolver,
  type Placed,
  type SlideSize,
} from '@pptx-studio/render-svg';

export interface StageProps {
  readonly sheet: Sheet;
  readonly size: SlideSize;
  readonly defaultTextStyle: ListStyle | undefined;
  readonly media?: MediaResolver;
  readonly width: number;
  readonly idPrefix: string;
  readonly selected?: number | null;
  readonly overlay?: boolean;
  onSelect?(id: number | null): void;
  onShapePointerDown?(placed: Placed, event: ReactPointerEvent): void;
}

export interface StageContent {
  readonly svg: string;
  readonly placed: readonly Placed[];
  readonly error: string | null;
}

/** The sheet drawn once, with every shape flattened out for hit-testing. */
export function useStage(
  sheet: Sheet,
  size: SlideSize,
  defaultTextStyle: ListStyle | undefined,
  idPrefix: string,
  media: MediaResolver | undefined,
): StageContent {
  return useMemo(() => {
    try {
      const { node, placed } = slideNode(sheet, size, {
        idPrefix,
        media,
        text: { defaultTextStyle },
      });
      return { svg: serializeSvg(node), placed: flatten(placed), error: null };
    } catch (cause) {
      return {
        svg: '',
        placed: [],
        error: cause instanceof Error ? cause.message : String(cause),
      };
    }
  }, [sheet, size, defaultTextStyle, idPrefix, media]);
}

/**
 * One slide on screen.
 *
 * Zoom is a CSS width on a fixed EMU viewBox and never a re-layout - font
 * metrics are not linear in point size, so laying out again at a new zoom would
 * move the line breaks.
 */
export function Stage({
  sheet,
  size,
  defaultTextStyle,
  media,
  width,
  idPrefix,
  selected = null,
  overlay = false,
  onSelect,
  onShapePointerDown,
}: StageProps) {
  const { svg, placed, error } = useStage(sheet, size, defaultTextStyle, idPrefix, media);
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
        className="flex items-center justify-center rounded border border-handle/40 bg-handle/5 p-4 text-center font-mono text-[12px] text-handle"
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

  return (
    <div
      className="stage-surface relative shrink-0 overflow-hidden"
      style={{ width, height }}
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
          viewBox={`0 0 ${size.cx} ${size.cy}`}
          className="pointer-events-none absolute inset-0 h-full w-full"
          dangerouslySetInnerHTML={{ __html: chrome }}
        />
      )}
    </div>
  );
}

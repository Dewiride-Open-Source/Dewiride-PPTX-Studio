'use client';

import {
  useEffect,
  useMemo,
  useRef,
  type KeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
} from 'react';

import type { Sheet } from '@pptx-studio/model';
import {
  flatten,
  serializeSvg,
  shapeOverlay,
  slideNode,
  type Frame,
  type MediaResolver,
  type Placed,
  type SlideSize,
  type TextEngine,
} from '@pptx-studio/render-svg';

import { graphicKind } from './locate.ts';

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
  text: TextEngine | null,
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

/** A click landed on a shape: the slide's own, or furniture the layout or master draws. */
export interface Hit {
  readonly placed: Placed;
  readonly own: boolean;
}

export interface StageProps {
  readonly content: StageContent;
  readonly size: SlideSize;
  readonly width: number;
  readonly sheet: Sheet;
  readonly selected: number | null;
  readonly debug: boolean;
  /** The shape whose drawn text is hidden beneath the text box. */
  readonly editing: number | null;
  readonly onSelect: (hit: Hit | null) => void;
  readonly onDragStart: (placed: Placed, event: ReactPointerEvent) => void;
  readonly onOpenText: (placed: Placed) => void;
  /** Arrow keys on the selected shape, in EMU; Shift moves ten times as far. */
  readonly onNudge: (id: number, dx: number, dy: number) => void;
  readonly onDelete: (id: number) => void;
  /** Drawn in the slide's own EMU space, over the selection. */
  readonly inSlide?: ReactNode;
  /** Drawn in pixels over the stage. */
  readonly children?: ReactNode;
}

const NUDGE_EMU = 12700;

/** True when a slide-EMU point lies inside a frame, rotation included. */
function inside(frame: Frame, x: number, y: number): boolean {
  const cx = frame.x + frame.cx / 2;
  const cy = frame.y + frame.cy / 2;
  const angle = (-frame.rot * Math.PI) / 180;
  const dx = x - cx;
  const dy = y - cy;
  const lx = dx * Math.cos(angle) - dy * Math.sin(angle);
  const ly = dx * Math.sin(angle) + dy * Math.cos(angle);
  return Math.abs(lx) <= frame.cx / 2 && Math.abs(ly) <= frame.cy / 2;
}

function Selection({ frame, unit }: { frame: Frame; unit: number }) {
  const stroke = 1.5 * unit;
  const handle = 7 * unit;
  const corners = [
    [frame.x, frame.y],
    [frame.x + frame.cx, frame.y],
    [frame.x, frame.y + frame.cy],
    [frame.x + frame.cx, frame.y + frame.cy],
  ];
  return (
    <g
      transform={`rotate(${String(frame.rot)} ${String(frame.x + frame.cx / 2)} ${String(frame.y + frame.cy / 2)})`}
    >
      <rect
        x={frame.x}
        y={frame.y}
        width={frame.cx}
        height={frame.cy}
        fill="none"
        stroke="var(--accent)"
        strokeWidth={stroke}
      />
      {corners.map(([x, y]) => (
        <rect
          key={`${String(x)}-${String(y)}`}
          x={(x ?? 0) - handle / 2}
          y={(y ?? 0) - handle / 2}
          width={handle}
          height={handle}
          fill="var(--surface)"
          stroke="var(--accent)"
          strokeWidth={stroke}
        />
      ))}
    </g>
  );
}

/**
 * The frames render-svg does not draw - charts, tables, diagrams, ink - as
 * labelled boxes the slide's own can be selected, moved and deleted from.
 */
function Frames({
  placed,
  sheet,
  unit,
}: {
  placed: readonly Placed[];
  sheet: Sheet;
  unit: number;
}) {
  const frames = placed.filter(
    (one) =>
      one.sheet === sheet &&
      (one.shape.kind === 'graphicFrame' || one.shape.kind === 'contentPart'),
  );
  return (
    <>
      {frames.map((one) => {
        const { frame } = one;
        const kind = one.shape.kind === 'contentPart' ? 'ink' : graphicKind(one.shape.node);
        const label = `${one.shape.name === '' ? kind : one.shape.name} · ${kind}, preserved and not drawn yet`;
        return (
          <g
            key={one.shape.cNvPrId}
            data-shape={one.shape.cNvPrId}
            data-name={one.shape.name}
            transform={`rotate(${String(frame.rot)} ${String(frame.x + frame.cx / 2)} ${String(frame.y + frame.cy / 2)})`}
            style={{ pointerEvents: 'auto', cursor: 'move' }}
          >
            <rect
              x={frame.x}
              y={frame.y}
              width={frame.cx}
              height={frame.cy}
              fill="var(--sunken)"
              fillOpacity={0.6}
              stroke="var(--line-strong)"
              strokeWidth={unit}
              strokeDasharray={`${String(6 * unit)} ${String(4 * unit)}`}
            />
            <text
              x={frame.x + frame.cx / 2}
              y={frame.y + frame.cy / 2}
              textAnchor="middle"
              dominantBaseline="middle"
              fontSize={13 * unit}
              fontFamily="var(--font-sans)"
              fill="var(--fg-muted)"
            >
              {label}
            </text>
          </g>
        );
      })}
    </>
  );
}

/**
 * One slide on screen, with the slide's own shapes selectable.
 *
 * Zoom is a CSS width on a fixed EMU viewBox and never a re-layout - font
 * metrics are not linear in point size, so laying out again at a new zoom would
 * move the line breaks.
 */
export function Stage({
  content,
  size,
  width,
  sheet,
  selected,
  debug,
  editing,
  onSelect,
  onDragStart,
  onOpenText,
  onNudge,
  onDelete,
  inSlide,
  children,
}: StageProps) {
  const { svg, placed, error } = content;
  const surface = useRef<HTMLDivElement>(null);
  const height = Math.round((width * size.cy) / size.cx);
  const unit = size.cx / width;
  const chosen =
    placed.find((one) => one.shape.cNvPrId === selected && one.sheet === sheet) ?? null;

  const overlay = useMemo(() => {
    if (!debug || chosen === null) return null;
    try {
      return serializeSvg(shapeOverlay(chosen, { unit }).node);
    } catch {
      // A shape whose geometry will not resolve has no overlay, which is no reason
      // to stop drawing the slide.
      return null;
    }
  }, [debug, chosen, unit]);

  // The drawn text under an open text box would show through it twice.
  useEffect(() => {
    const element = surface.current;
    if (element === null || editing === null) return;
    const groups = element.querySelectorAll<SVGGElement>(`[data-text="${String(editing)}"]`);
    const group = groups[groups.length - 1];
    if (group === undefined) return;
    group.style.visibility = 'hidden';
    return () => {
      group.style.visibility = '';
    };
  }, [editing, svg]);

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

  const pick = (event: ReactMouseEvent<HTMLDivElement>): Hit | null => {
    const target = event.target as Element | null;
    // Text is drawn in a group beside the shape's geometry, not inside it, so a
    // click on a word never reaches `[data-shape]`.
    const hit = target?.closest?.('[data-shape], [data-text]');
    const id = hit?.getAttribute('data-shape') ?? hit?.getAttribute('data-text');
    if (id === null || id === undefined) return null;
    const candidates = placed.filter((one) => one.shape.cNvPrId === Number(id));
    const own = candidates.find((one) => one.sheet === sheet);
    const first = candidates[0];
    if (own === undefined) return first === undefined ? null : { placed: first, own: false };
    if (candidates.length === 1) return { placed: own, own: true };
    const box = event.currentTarget.getBoundingClientRect();
    const x = (event.clientX - box.left) * unit;
    const y = (event.clientY - box.top) * unit;
    return inside(own.frame, x, y)
      ? { placed: own, own: true }
      : { placed: first ?? own, own: false };
  };

  const keys = (event: KeyboardEvent<HTMLDivElement>) => {
    if (editing !== null) return;
    if (event.key === 'Escape') {
      onSelect(null);
      return;
    }
    if (chosen === null) return;
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      onDelete(chosen.shape.cNvPrId);
      return;
    }
    if (event.key === 'Enter' || event.key === 'F2') {
      event.preventDefault();
      onOpenText(chosen);
      return;
    }
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
    onNudge(chosen.shape.cNvPrId, found[0], found[1]);
  };

  return (
    <div
      ref={surface}
      role="application"
      aria-label={
        chosen === null
          ? 'Slide: click a shape to select it, double-click to type in it'
          : `Slide, ${chosen.shape.name === '' ? 'a shape' : chosen.shape.name} selected: drag or arrow keys move it, Enter edits its text, Delete removes it`
      }
      tabIndex={0}
      className="stage-surface relative shrink-0 overflow-hidden outline-none focus-visible:ring-2 focus-visible:ring-accent"
      style={{ width, height, touchAction: 'none' }}
      onKeyDown={keys}
      onPointerDown={(event) => {
        const hit = pick(event);
        onSelect(hit);
        if (hit?.own === true) onDragStart(hit.placed, event);
      }}
      onDoubleClick={(event) => {
        const hit = pick(event);
        if (hit?.own === true) onOpenText(hit.placed);
      }}
    >
      <div
        className="absolute inset-0 [&>svg]:h-full [&>svg]:w-full"
        dangerouslySetInnerHTML={{ __html: svg }}
      />
      <svg
        viewBox={`0 0 ${String(size.cx)} ${String(size.cy)}`}
        className="pointer-events-none absolute inset-0 h-full w-full"
        aria-hidden={editing === null}
      >
        <Frames placed={placed} sheet={sheet} unit={unit} />
        {overlay === null ? null : <g dangerouslySetInnerHTML={{ __html: overlay }} />}
        {chosen === null || editing !== null ? null : (
          <Selection frame={chosen.frame} unit={unit} />
        )}
        {inSlide}
      </svg>
      {children}
    </div>
  );
}

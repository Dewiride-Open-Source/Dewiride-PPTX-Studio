'use client';

import { Palette, Trash2, Type } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

import type { Frame, Placed } from '@pptx-studio/render-svg';

const SWATCHES = [
  '#E8453C',
  '#D98A2B',
  '#F2C744',
  '#2FA869',
  '#0B62B0',
  '#7B1FA2',
  '#111111',
  '#FFFFFF',
];

/** The topmost point and the horizontal centre of a rotated frame, in EMU. */
function above(frame: Frame): { top: number; bottom: number; centre: number } {
  const cx = frame.x + frame.cx / 2;
  const cy = frame.y + frame.cy / 2;
  const angle = (frame.rot * Math.PI) / 180;
  const ys = [
    [-frame.cx / 2, -frame.cy / 2],
    [frame.cx / 2, -frame.cy / 2],
    [-frame.cx / 2, frame.cy / 2],
    [frame.cx / 2, frame.cy / 2],
  ].map(([x, y]) => cy + (x ?? 0) * Math.sin(angle) + (y ?? 0) * Math.cos(angle));
  return { top: Math.min(...ys), bottom: Math.max(...ys), centre: cx };
}

export interface ToolbarProps {
  readonly placed: Placed;
  /** EMU per stage pixel. */
  readonly unit: number;
  readonly stage: { readonly width: number; readonly height: number };
  /** The shape's colour as drawn, for the swatch that shows it. */
  readonly colour: string | null;
  /** Why the colour cannot be changed, or null when it can. */
  readonly colourRefused: string | null;
  readonly textRefused: string | null;
  readonly onColour: (hex: string, live: boolean) => void;
  readonly onColourEnd: () => void;
  readonly onText: () => void;
  readonly onDelete: () => void;
}

const HEIGHT = 36;

/** Three controls over the selected shape: its colour, its text, and delete. */
export function Toolbar({
  placed,
  unit,
  stage,
  colour,
  colourRefused,
  textRefused,
  onColour,
  onColourEnd,
  onText,
  onDelete,
}: ToolbarProps) {
  const [open, setOpen] = useState(false);
  const root = useRef<HTMLDivElement>(null);
  const picking = useRef(false);
  const { top, bottom, centre } = above(placed.frame);
  const gap = 10;
  const wanted = top / unit - HEIGHT - gap;
  const y = wanted >= 4 ? wanted : Math.min(bottom / unit + gap, stage.height - HEIGHT - 4);
  const x = Math.min(Math.max(centre / unit, 96), stage.width - 96);

  useEffect(() => {
    if (!open) return;
    const away = (event: PointerEvent): void => {
      if (root.current?.contains(event.target as Node) === false) setOpen(false);
    };
    window.addEventListener('pointerdown', away);
    return () => window.removeEventListener('pointerdown', away);
  }, [open]);

  const name = placed.shape.name === '' ? 'this shape' : placed.shape.name;

  return (
    <div
      ref={root}
      role="toolbar"
      aria-label={`Edit ${name}`}
      className="absolute z-10 flex -translate-x-1/2 items-center gap-0.5 rounded-control border border-line-strong bg-surface p-1 shadow-panel"
      style={{ left: x, top: y, height: HEIGHT }}
      onPointerDown={(event) => event.stopPropagation()}
    >
      <button
        type="button"
        title={colourRefused ?? 'Colour'}
        aria-label={colourRefused ?? 'Colour'}
        aria-expanded={open}
        disabled={colourRefused !== null}
        onClick={() => setOpen((was) => !was)}
        className="inline-flex h-7 items-center gap-1.5 rounded-control px-2 text-[12px] text-fg hover:bg-sunken disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Palette size={15} aria-hidden />
        <span
          className="h-3.5 w-3.5 rounded-sm border border-line-strong"
          style={{ background: colour ?? 'transparent' }}
        />
      </button>
      <button
        type="button"
        title={textRefused ?? 'Edit the text (or double-click the shape)'}
        aria-label={textRefused ?? 'Edit the text'}
        disabled={textRefused !== null}
        onClick={onText}
        className="inline-flex h-7 items-center gap-1.5 rounded-control px-2 text-[12px] text-fg hover:bg-sunken disabled:cursor-not-allowed disabled:opacity-40"
      >
        <Type size={15} aria-hidden />
        Text
      </button>
      <button
        type="button"
        title="Delete (Delete key)"
        aria-label={`Delete ${name}`}
        onClick={onDelete}
        className="inline-flex h-7 items-center rounded-control px-2 text-[12px] text-fg hover:bg-bad/10 hover:text-bad"
      >
        <Trash2 size={15} aria-hidden />
      </button>

      {open ? (
        <div
          role="group"
          aria-label="Colours"
          className="absolute top-full left-0 mt-1.5 flex w-52 flex-wrap items-center gap-1.5 rounded-control border border-line-strong bg-surface p-2 shadow-panel"
        >
          {SWATCHES.map((hex) => (
            <button
              key={hex}
              type="button"
              title={hex}
              aria-label={`Colour ${name} ${hex}`}
              aria-pressed={colour?.toUpperCase() === hex}
              onClick={() => {
                onColour(hex, false);
                setOpen(false);
              }}
              className={`h-7 w-7 rounded-control border ${colour?.toUpperCase() === hex ? 'border-accent ring-2 ring-accent/40' : 'border-line-strong'}`}
              style={{ background: hex }}
            />
          ))}
          <label className="ml-auto inline-flex h-7 cursor-pointer items-center gap-1.5 rounded-control border border-line-strong px-2 text-[12px] text-fg-muted hover:text-fg">
            Any
            <input
              type="color"
              aria-label={`Any colour for ${name}`}
              defaultValue={colour ?? '#0B62B0'}
              className="h-5 w-6 cursor-pointer border-0 bg-transparent p-0"
              onInput={(event) => {
                picking.current = true;
                onColour(event.currentTarget.value, true);
              }}
              onChange={() => {
                if (!picking.current) return;
                picking.current = false;
                onColourEnd();
              }}
            />
          </label>
        </div>
      ) : null}
    </div>
  );
}

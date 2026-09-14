'use client';

import { Baseline, PaintBucket, Pencil, Trash2 } from 'lucide-react';
import { useEffect, useRef, useState, type ReactNode } from 'react';

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

/** One colour the shape carries, as drawn, and why it cannot be changed when it cannot. */
export interface Colour {
  readonly value: string | null;
  readonly refused: string | null;
  readonly onPick: (hex: string, live: boolean) => void;
  readonly onEnd: () => void;
}

export interface ToolbarProps {
  readonly placed: Placed;
  /** EMU per stage pixel. */
  readonly unit: number;
  readonly stage: { readonly width: number; readonly height: number };
  readonly fill: Colour;
  readonly text: Colour;
  readonly textRefused: string | null;
  readonly onText: () => void;
  readonly onDelete: () => void;
}

const HEIGHT = 36;
const BUTTON =
  'inline-flex h-7 items-center gap-1.5 rounded-control px-2 text-[12px] text-fg hover:bg-sunken disabled:cursor-not-allowed disabled:opacity-40';

interface SwatchesProps {
  readonly what: string;
  readonly name: string;
  readonly colour: Colour;
  readonly onClose: () => void;
}

/** Eight swatches and the native picker; a drag in the picker is one gesture. */
function Swatches({ what, name, colour, onClose }: SwatchesProps) {
  const picking = useRef(false);
  const current = colour.value?.toUpperCase() ?? null;
  return (
    <div
      role="group"
      aria-label={`${what} colours`}
      className="absolute top-full left-0 mt-1.5 flex w-52 flex-wrap items-center gap-1.5 rounded-control border border-line-strong bg-surface p-2 shadow-panel"
    >
      {SWATCHES.map((hex) => (
        <button
          key={hex}
          type="button"
          title={hex}
          aria-label={`${what} ${name} ${hex}`}
          aria-pressed={current === hex}
          onClick={() => {
            colour.onPick(hex, false);
            onClose();
          }}
          className={`h-7 w-7 rounded-control border ${current === hex ? 'border-accent ring-2 ring-accent/40' : 'border-line-strong'}`}
          style={{ background: hex }}
        />
      ))}
      <label className="ml-auto inline-flex h-7 cursor-pointer items-center gap-1.5 rounded-control border border-line-strong px-2 text-[12px] text-fg-muted hover:text-fg">
        Any
        <input
          type="color"
          aria-label={`Any ${what.toLowerCase()} colour for ${name}`}
          defaultValue={current ?? '#0B62B0'}
          className="h-5 w-6 cursor-pointer border-0 bg-transparent p-0"
          onInput={(event) => {
            picking.current = true;
            colour.onPick(event.currentTarget.value, true);
          }}
          onChange={() => {
            if (!picking.current) return;
            picking.current = false;
            colour.onEnd();
          }}
        />
      </label>
    </div>
  );
}

interface ColourButtonProps {
  readonly what: string;
  readonly name: string;
  readonly colour: Colour;
  readonly icon: ReactNode;
  readonly open: boolean;
  readonly onToggle: () => void;
  readonly onClose: () => void;
}

/** A labelled button carrying the colour as drawn, opening its swatches. */
function ColourButton({ what, name, colour, icon, open, onToggle, onClose }: ColourButtonProps) {
  const label = colour.refused ?? `${what} colour of ${name}`;
  return (
    <div className="relative">
      <button
        type="button"
        title={colour.refused ?? `${what} colour`}
        aria-label={label}
        aria-expanded={open}
        disabled={colour.refused !== null}
        onClick={onToggle}
        className={BUTTON}
      >
        {icon}
        {what}
        <span
          className="h-3.5 w-3.5 rounded-sm border border-line-strong"
          style={{ background: colour.value ?? 'transparent' }}
        />
      </button>
      {open ? <Swatches what={what} name={name} colour={colour} onClose={onClose} /> : null}
    </div>
  );
}

/** Four controls over the selected shape: its fill, its text colour, its text, and delete. */
export function Toolbar({
  placed,
  unit,
  stage,
  fill,
  text,
  textRefused,
  onText,
  onDelete,
}: ToolbarProps) {
  const [open, setOpen] = useState<'fill' | 'text' | null>(null);
  const root = useRef<HTMLDivElement>(null);
  const { top, bottom, centre } = above(placed.frame);
  const gap = 10;
  const wanted = top / unit - HEIGHT - gap;
  const y = wanted >= 4 ? wanted : Math.min(bottom / unit + gap, stage.height - HEIGHT - 4);
  const x = Math.min(Math.max(centre / unit, 140), stage.width - 140);

  useEffect(() => {
    if (open === null) return;
    const away = (event: PointerEvent): void => {
      if (root.current?.contains(event.target as Node) === false) setOpen(null);
    };
    window.addEventListener('pointerdown', away);
    return () => window.removeEventListener('pointerdown', away);
  }, [open]);

  const name = placed.shape.name === '' ? 'this shape' : placed.shape.name;
  const toggle = (which: 'fill' | 'text') => () => setOpen((was) => (was === which ? null : which));
  // A picked swatch unmounts under the pointer; focus goes back to the slide so the keys still work.
  const picked = (): void => {
    setOpen(null);
    root.current?.closest<HTMLElement>('.stage-surface')?.focus();
  };

  return (
    <div
      ref={root}
      role="toolbar"
      aria-label={`Edit ${name}`}
      className="absolute z-10 flex -translate-x-1/2 items-center gap-0.5 rounded-control border border-line-strong bg-surface p-1 shadow-panel"
      style={{ left: x, top: y, height: HEIGHT }}
      onPointerDown={(event) => event.stopPropagation()}
    >
      <ColourButton
        what="Fill"
        name={name}
        colour={fill}
        icon={<PaintBucket size={15} aria-hidden />}
        open={open === 'fill'}
        onToggle={toggle('fill')}
        onClose={picked}
      />
      <ColourButton
        what="Text"
        name={name}
        colour={text}
        icon={<Baseline size={15} aria-hidden />}
        open={open === 'text'}
        onToggle={toggle('text')}
        onClose={picked}
      />
      <button
        type="button"
        title={textRefused ?? 'Edit the text (or double-click the shape)'}
        aria-label={textRefused ?? `Edit the text of ${name}`}
        disabled={textRefused !== null}
        onClick={onText}
        className={BUTTON}
      >
        <Pencil size={15} aria-hidden />
        Edit
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
    </div>
  );
}

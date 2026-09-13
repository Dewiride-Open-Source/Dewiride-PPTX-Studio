'use client';

import {
  dragHandle,
  type PresetShape,
  type ResolvedGeometry,
  type ResolvedHandle,
} from '@pptx-studio/geometry';

import { Button } from '@/design/button';
import { Mono } from '@/design/code';

const PATH_COLOURS = ['#E8453C', '#1B9E4B', '#F09300', '#9C27B0', '#00838F'];
const MARGIN = 20;

interface StageProps {
  readonly preset: PresetShape;
  readonly size: { readonly w: number; readonly h: number };
  readonly geometry: ResolvedGeometry;
  readonly handles: readonly ResolvedHandle[];
  readonly adjust: Readonly<Record<string, number>>;
  readonly onAdjust: (adjust: Record<string, number>) => void;
}

/** The shape at one size, with draggable handles and a keyboard proxy for each. */
export function Stage({ preset, size, geometry, handles, adjust, onAdjust }: StageProps) {
  const drag = (at: number, to: { x: number; y: number }) => {
    const declared = preset.ahLst[at];
    if (declared === undefined) return;
    try {
      onAdjust(dragHandle(preset, size, declared, to, { adjust }));
    } catch {
      // A point outside the handle's declared range has no answer; the last good one stands.
    }
  };

  return (
    <div>
      <div className="flex justify-center p-4">
        <svg
          viewBox={`${String(-MARGIN)} ${String(-MARGIN)} ${String(size.w + 2 * MARGIN)} ${String(size.h + 2 * MARGIN)}`}
          className="h-auto w-full max-w-lg touch-none"
          role="img"
          aria-label={`${preset.name} at ${String(size.w)} by ${String(size.h)}`}
        >
          <rect
            x={0}
            y={0}
            width={size.w}
            height={size.h}
            fill="none"
            stroke="var(--line-strong)"
            strokeDasharray="4 4"
          />
          {geometry.paths.map((path, at) => (
            <path
              key={at}
              d={path.d}
              fill={path.fill === 'none' ? 'none' : 'rgb(59 143 212 / 0.2)'}
              stroke={PATH_COLOURS[at % PATH_COLOURS.length]}
              strokeWidth={2}
            />
          ))}
          {geometry.textRect === null ? null : (
            <rect
              x={geometry.textRect.l}
              y={geometry.textRect.t}
              width={geometry.textRect.r - geometry.textRect.l}
              height={geometry.textRect.b - geometry.textRect.t}
              fill="none"
              stroke="#a855c9"
              strokeDasharray="3 3"
            />
          )}
          {geometry.connectionSites.map((site, at) => (
            <circle key={at} cx={site.pos.x} cy={site.pos.y} r={3.5} fill="#00a5b5" />
          ))}
          {handles.map((handle, at) =>
            !handle.finite ? null : (
              <rect
                key={at}
                x={handle.pos.x - 7}
                y={handle.pos.y - 7}
                width={14}
                height={14}
                fill="#E8453C"
                className="cursor-grab"
                onPointerDown={(event) => {
                  event.currentTarget.setPointerCapture(event.pointerId);
                  const svg = event.currentTarget.ownerSVGElement;
                  if (svg === null) return;
                  const move = (moveEvent: PointerEvent): void => {
                    const box = svg.getBoundingClientRect();
                    const scale = (size.w + 2 * MARGIN) / box.width;
                    drag(at, {
                      x: (moveEvent.clientX - box.left) * scale - MARGIN,
                      y: (moveEvent.clientY - box.top) * scale - MARGIN,
                    });
                  };
                  const up = (): void => {
                    window.removeEventListener('pointermove', move);
                    window.removeEventListener('pointerup', up);
                  };
                  window.addEventListener('pointermove', move);
                  window.addEventListener('pointerup', up);
                }}
              />
            ),
          )}
        </svg>
      </div>
      <div className="flex flex-wrap items-center gap-3 border-t border-line px-4 py-2.5 text-[11px] text-fg-muted">
        <span className="flex items-center gap-1.5">
          <span className="inline-block h-2 w-2 bg-bad" /> adjust handle
        </span>
        <span className="flex items-center gap-1.5">
          <span className="inline-block h-2 w-2 rounded-full bg-site" /> connection site
        </span>
        <span className="flex items-center gap-1.5">
          <span className="inline-block h-2 w-2 border border-info" /> a:rect
        </span>
      </div>
      {handles.length === 0 ? null : (
        <ul className="divide-y divide-line border-t border-line" aria-label="Handles">
          {handles.map((handle, at) => (
            <li key={at} className="flex flex-wrap items-center gap-2 px-4 py-2 text-[12px]">
              <span className="font-mono text-fg-muted">
                handle {at + 1} · {handle.kind}
              </span>
              {handle.axes.map((axis) => (
                <span key={axis.guide} className="inline-flex items-center gap-1">
                  <Mono>{axis.guide}</Mono>
                  <span className="tabular text-fg">{axis.value}</span>
                  <span className="text-fg-faint">
                    [{axis.min}, {axis.max}]{axis.widened ? ' widened' : ''}
                  </span>
                </span>
              ))}
              <span className="ml-auto inline-flex gap-1">
                {(
                  [
                    ['←', -8, 0],
                    ['→', 8, 0],
                    ['↑', 0, -8],
                    ['↓', 0, 8],
                  ] as const
                ).map(([label, dx, dy]) => (
                  <Button
                    key={label}
                    size="sm"
                    variant="ghost"
                    aria-label={`Move handle ${String(at + 1)} ${label}`}
                    onClick={() => drag(at, { x: handle.pos.x + dx, y: handle.pos.y + dy })}
                    disabled={!handle.finite}
                  >
                    {label}
                  </Button>
                ))}
              </span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

'use client';

import { useMemo, useState } from 'react';

import {
  PRESET_SOURCE,
  dragHandle,
  getPreset,
  presetNames,
  resolveGeometry,
  resolveHandles,
} from '@pptx-studio/geometry';

import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';

const HOW = `import { getPreset, resolveGeometry, resolveHandles, dragHandle } from '@pptx-studio/geometry';

const preset = getPreset('leftArrow');
const geom = resolveGeometry(preset, { w: 400, h: 240 }, { adjust });

geom.paths[0].d;        // SVG path data, empty when a guide went non-finite
geom.guides;            // every built-in and computed guide, at this size
geom.textRect;          // the a:rect, which is not the shape's own box
geom.connectionSites;   // where a connector may attach, with its angle

// Dragging inverts the formula rather than guessing.
const next = dragHandle(preset, size, preset.ahLst[0], { x, y });`;

const SIZE = { w: 400, h: 260 };
const SWATCH = { w: 96, h: 64 };

/** One preset drawn at a given size, or null if its guides went non-finite. */
function useDrawn(name: string, size: { w: number; h: number }, adjust: Record<string, number>) {
  return useMemo(() => {
    const preset = getPreset(name);
    if (preset === undefined) return null;
    try {
      const geometry = resolveGeometry(preset, size, { adjust });
      const handles = resolveHandles(preset, size, { adjust });
      return { preset, geometry, handles };
    } catch {
      return null;
    }
  }, [name, size, adjust]);
}

function Swatch({ name }: { name: string }) {
  const drawn = useDrawn(name, SWATCH, EMPTY_ADJUST);
  if (drawn === null) return null;
  return (
    <svg viewBox={`0 0 ${SWATCH.w} ${SWATCH.h}`} className="h-12 w-full">
      {drawn.geometry.paths.map((path, at) => (
        <path
          key={at}
          d={path.d}
          fill={path.fill === 'none' ? 'none' : 'rgb(59 143 212 / 0.25)'}
          stroke="#3b8fd4"
          strokeWidth={1.2}
        />
      ))}
    </svg>
  );
}

const EMPTY_ADJUST: Record<string, number> = {};

export default function GeometryPage() {
  const names = useMemo(() => presetNames(), []);
  const [query, setQuery] = useState('');
  const [chosen, setChosen] = useState('leftArrow');
  const [adjust, setAdjust] = useState<Record<string, number>>(EMPTY_ADJUST);
  const [showGuides, setShowGuides] = useState(false);

  const shown = useMemo(() => {
    const needle = query.trim().toLowerCase();
    return needle === '' ? names : names.filter((one) => one.toLowerCase().includes(needle));
  }, [names, query]);

  const drawn = useDrawn(chosen, SIZE, adjust);

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Geometry</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/geometry</Mono> - the 187 presets, their guides and their adjust
          handles.
        </p>
      </header>

      <Panel>
        <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-4">
          <Stat label="presets" value={names.length} />
          <Stat label="showing" value={shown.length} />
          <Stat label="handles here" value={drawn?.handles.length ?? 0} />
          <Stat label="guides here" value={drawn?.geometry.guides.size ?? 0} />
        </div>
        <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-ink-500">
          Transcoded at build time from {PRESET_SOURCE.file} ({PRESET_SOURCE.bytes.toLocaleString()}{' '}
          bytes), whose root element is misspelled <Mono tone="dim">{PRESET_SOURCE.root}</Mono>.
        </p>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[minmax(0,22rem)_1fr]">
        <Panel title="Presets">
          <div className="border-b border-ink-700 p-3">
            <input
              type="search"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Search 187 shapes"
              className="w-full rounded border border-ink-600 bg-ink-800 px-2 py-1.5 text-[13px] text-ink-100"
            />
          </div>
          <div className="grid max-h-[30rem] grid-cols-3 gap-1 overflow-y-auto p-2">
            {shown.map((name) => (
              <button
                key={name}
                type="button"
                onClick={() => {
                  setChosen(name);
                  setAdjust(EMPTY_ADJUST);
                }}
                title={name}
                className={`rounded border p-1.5 ${
                  name === chosen
                    ? 'border-chrome bg-chrome/10'
                    : 'border-transparent hover:border-ink-600'
                }`}
              >
                <Swatch name={name} />
                <span className="mt-1 block truncate font-mono text-[10px] text-ink-400">
                  {name}
                </span>
              </button>
            ))}
          </div>
        </Panel>

        <div className="flex flex-col gap-4">
          <Panel
            title={chosen}
            hint="Drag a red handle. The formula is inverted, not guessed."
            aside={
              drawn === null || drawn.preset.ahLst.length === 0 ? (
                <Badge tone="plain">no handles</Badge>
              ) : (
                <Badge tone="info">{drawn.preset.ahLst.length} handle(s)</Badge>
              )
            }
          >
            {drawn === null ? (
              <p className="p-4 text-[13px] text-handle">This preset would not resolve.</p>
            ) : (
              <div className="flex justify-center p-4">
                <svg
                  viewBox={`-20 -20 ${SIZE.w + 40} ${SIZE.h + 40}`}
                  className="h-auto w-full max-w-lg touch-none"
                >
                  <rect
                    x={0}
                    y={0}
                    width={SIZE.w}
                    height={SIZE.h}
                    fill="none"
                    stroke="#262d37"
                    strokeDasharray="4 4"
                  />
                  {drawn.geometry.paths.map((path, at) => (
                    <path
                      key={at}
                      d={path.d}
                      fill={path.fill === 'none' ? 'none' : 'rgb(59 143 212 / 0.2)'}
                      stroke={['#E8453C', '#1B9E4B', '#F09300', '#9C27B0', '#00838F'][at % 5]}
                      strokeWidth={2}
                    />
                  ))}
                  {drawn.geometry.textRect === null ? null : (
                    <rect
                      x={drawn.geometry.textRect.l}
                      y={drawn.geometry.textRect.t}
                      width={drawn.geometry.textRect.r - drawn.geometry.textRect.l}
                      height={drawn.geometry.textRect.b - drawn.geometry.textRect.t}
                      fill="none"
                      stroke="#a855c9"
                      strokeDasharray="3 3"
                    />
                  )}
                  {drawn.geometry.connectionSites.map((site, at) => (
                    <circle key={at} cx={site.pos.x} cy={site.pos.y} r={3.5} fill="#00a5b5" />
                  ))}
                  {drawn.handles.map((handle, at) => {
                    const declared = drawn.preset.ahLst[at];
                    if (declared === undefined || !handle.finite) return null;
                    return (
                      <rect
                        key={at}
                        x={handle.pos.x - 6}
                        y={handle.pos.y - 6}
                        width={12}
                        height={12}
                        fill="#E8453C"
                        className="cursor-grab"
                        onPointerDown={(event) => {
                          event.currentTarget.setPointerCapture(event.pointerId);
                          const svg = event.currentTarget.ownerSVGElement;
                          if (svg === null) return;
                          const move = (moveEvent: PointerEvent): void => {
                            const box = svg.getBoundingClientRect();
                            const scale = (SIZE.w + 40) / box.width;
                            const to = {
                              x: (moveEvent.clientX - box.left) * scale - 20,
                              y: (moveEvent.clientY - box.top) * scale - 20,
                            };
                            try {
                              setAdjust(dragHandle(drawn.preset, SIZE, declared, to, { adjust }));
                            } catch {
                              // A drag outside the handle's declared range has no
                              // answer; keeping the last good one is the right refusal.
                            }
                          };
                          const up = (): void => {
                            window.removeEventListener('pointermove', move);
                            window.removeEventListener('pointerup', up);
                          };
                          window.addEventListener('pointermove', move);
                          window.addEventListener('pointerup', up);
                        }}
                      />
                    );
                  })}
                </svg>
              </div>
            )}
            <div className="flex flex-wrap items-center gap-3 border-t border-ink-700 px-4 py-2.5 text-[11px]">
              <span className="flex items-center gap-1.5 text-ink-400">
                <span className="inline-block h-2 w-2 bg-handle" /> adjust handle
              </span>
              <span className="flex items-center gap-1.5 text-ink-400">
                <span className="inline-block h-2 w-2 rounded-full bg-site" /> connection site
              </span>
              <span className="flex items-center gap-1.5 text-ink-400">
                <span className="inline-block h-2 w-2 border border-textrect" /> a:rect
              </span>
              {adjust === EMPTY_ADJUST ? null : (
                <button
                  type="button"
                  onClick={() => setAdjust(EMPTY_ADJUST)}
                  className="ml-auto rounded border border-ink-600 px-2 py-0.5 text-ink-300"
                >
                  Reset adjusts
                </button>
              )}
            </div>
          </Panel>

          {drawn === null ? null : (
            <Panel
              title="Guides"
              hint="Every built-in and computed guide, at the size on screen."
              aside={
                <button
                  type="button"
                  onClick={() => setShowGuides((on) => !on)}
                  className="rounded border border-ink-600 px-2 py-0.5 text-[11px] text-ink-300"
                >
                  {showGuides ? 'Hide' : 'Show'}
                </button>
              }
            >
              {showGuides ? (
                <div className="grid max-h-64 grid-cols-2 gap-x-6 gap-y-1 overflow-y-auto p-4 sm:grid-cols-3">
                  {[...drawn.geometry.guides.entries()].map(([name, value]) => (
                    <div key={name} className="flex justify-between gap-2 text-[12px]">
                      <Mono tone="dim">{name}</Mono>
                      <span
                        className={`tabular ${Number.isFinite(value) ? 'text-ink-200' : 'text-handle'}`}
                      >
                        {Number.isFinite(value) ? value.toFixed(1) : 'non-finite'}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="p-4 text-[13px] text-ink-400">
                  {drawn.geometry.guides.size} guides resolved for this shape at {SIZE.w} x {SIZE.h}
                  .
                </p>
              )}
            </Panel>
          )}

          <Panel title="How">
            <div className="p-4">
              <Snippet code={HOW} />
            </div>
          </Panel>
        </div>
      </div>
    </div>
  );
}

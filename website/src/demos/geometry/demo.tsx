'use client';

import { useMemo, useState } from 'react';

import { PRESET_SOURCE, getPreset, resolveGeometry, resolveHandles } from '@pptx-studio/geometry';

import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Labelled, Slider } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';
import type { DemoProps } from '../registry';
import { Gallery } from './gallery';
import { Stage } from './stage';

const HOW = `import { getPreset, resolveGeometry, resolveHandles, dragHandle } from '@pptx-studio/geometry';

const preset = getPreset('leftArrow');
const geom = resolveGeometry(preset, { w: 400, h: 240 }, { adjust });

geom.paths[0].d;        // SVG path data, empty when a guide went non-finite
geom.guides;            // every built-in and computed guide, at this size
geom.textRect;          // the a:rect, which is not the shape's own box
geom.connectionSites;   // where a connector may attach, with its angle

// Dragging inverts the formula rather than guessing, and returns what a:avLst should hold.
const next = dragHandle(preset, size, preset.ahLst[0], { x, y });`;

const EMPTY_ADJUST: Record<string, number> = {};

/** The `a:avLst` a save would write for the adjusts the handles set. */
function avLst(prst: string, adjust: Readonly<Record<string, number>>): string {
  const entries = Object.entries(adjust);
  if (entries.length === 0) return `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
  const guides = entries
    .map(([name, value]) => `  <a:gd name="${name}" fmla="val ${String(Math.round(value))}"/>`)
    .join('\n');
  return `<a:prstGeom prst="${prst}">\n  <a:avLst>\n${guides.replace(/^/gm, '  ')}\n  </a:avLst>\n</a:prstGeom>`;
}

export default function GeometryDemo({ full }: DemoProps) {
  const [chosen, setChosen] = useState('leftArrow');
  const [adjust, setAdjust] = useState<Record<string, number>>(EMPTY_ADJUST);
  const [size, setSize] = useState({ w: 400, h: 260 });
  const [showGuides, setShowGuides] = useState(false);

  const drawn = useMemo(() => {
    const preset = getPreset(chosen);
    if (preset === undefined) return null;
    return {
      preset,
      geometry: resolveGeometry(preset, size, { adjust }),
      handles: resolveHandles(preset, size, { adjust }),
    };
  }, [chosen, size, adjust]);

  const nonFinite =
    drawn === null
      ? 0
      : [...drawn.geometry.guides.values()].filter((v) => !Number.isFinite(v)).length;

  return (
    <div className="flex flex-col gap-4">
      <Panel>
        <StatRow>
          <Stat label="presets" value={187} />
          <Stat label="handles here" value={drawn?.handles.length ?? 0} />
          <Stat label="guides here" value={drawn?.geometry.guides.size ?? 0} />
          <Stat label="non-finite" value={nonFinite} tone={nonFinite === 0 ? 'good' : 'warn'} />
          <Stat label="paths" value={drawn?.geometry.paths.length ?? 0} />
        </StatRow>
        <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
          Transcoded at build time from {PRESET_SOURCE.file} ({PRESET_SOURCE.bytes.toLocaleString()}{' '}
          bytes), whose root element is misspelled <Mono tone="dim">{PRESET_SOURCE.root}</Mono>.
        </p>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[minmax(0,22rem)_1fr]">
        <Panel title="Presets">
          <Gallery
            chosen={chosen}
            onChoose={(name) => {
              setChosen(name);
              setAdjust(EMPTY_ADJUST);
            }}
          />
        </Panel>

        <div className="flex flex-col gap-4">
          <Panel
            title={chosen}
            hint="Drag a red handle, or nudge it from the list beneath. The formula is inverted, not guessed."
            aside={
              drawn === null || drawn.preset.ahLst.length === 0 ? (
                <Badge>no handles</Badge>
              ) : (
                <Badge tone="info">{drawn.preset.ahLst.length} handle(s)</Badge>
              )
            }
          >
            {drawn === null ? (
              <p className="p-4 text-[13px] text-bad">This preset would not resolve.</p>
            ) : (
              <Stage
                preset={drawn.preset}
                size={size}
                geometry={drawn.geometry}
                handles={drawn.handles}
                adjust={adjust}
                onAdjust={setAdjust}
              />
            )}
            <div className="flex flex-wrap items-center gap-4 border-t border-line px-4 py-2.5">
              <Labelled label="width" inline>
                {(id) => (
                  <Slider
                    id={id}
                    min={0}
                    max={800}
                    value={size.w}
                    onChange={(w) => setSize({ ...size, w })}
                  />
                )}
              </Labelled>
              <Labelled label="height" inline>
                {(id) => (
                  <Slider
                    id={id}
                    min={0}
                    max={520}
                    value={size.h}
                    onChange={(h) => setSize({ ...size, h })}
                  />
                )}
              </Labelled>
              {adjust === EMPTY_ADJUST ? null : (
                <Button size="sm" className="ml-auto" onClick={() => setAdjust(EMPTY_ADJUST)}>
                  Reset adjusts
                </Button>
              )}
            </div>
          </Panel>

          {drawn === null ? null : (
            <div className="grid gap-4 md:grid-cols-2">
              <Panel
                title="What a save would write"
                hint="The adjusts, as a:avLst; unchanged handles write nothing."
              >
                <div className="p-4">
                  <Code lang="xml" code={avLst(chosen, adjust)} />
                </div>
              </Panel>
              <Panel
                title="Guides"
                hint="Every built-in and computed guide, at the size on screen."
                aside={
                  <Button
                    size="sm"
                    variant="ghost"
                    onClick={() => setShowGuides((on) => !on)}
                    aria-expanded={showGuides}
                  >
                    {showGuides ? 'Hide' : 'Show'}
                  </Button>
                }
              >
                {showGuides ? (
                  <div className="grid max-h-64 grid-cols-2 gap-x-6 gap-y-1 overflow-y-auto p-4">
                    {[...drawn.geometry.guides.entries()].map(([name, value]) => (
                      <div key={name} className="flex justify-between gap-2 text-[12px]">
                        <Mono tone="dim">{name}</Mono>
                        <span
                          className={`tabular ${Number.isFinite(value) ? 'text-fg' : 'text-bad'}`}
                        >
                          {Number.isFinite(value) ? value.toFixed(1) : 'non-finite'}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="p-4 text-[13px] text-fg-muted">
                    {drawn.geometry.guides.size} guides resolved at {size.w} × {size.h}. Division by
                    zero is reachable at zero width; it does not throw, it marks the guide.
                  </p>
                )}
              </Panel>
            </div>
          )}
        </div>
      </div>

      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Nothing here is painted: fills and strokes are the paint package&apos;s job, and the
              colours on this stage are the debug overlay&apos;s. A deck&apos;s custom geometry is
              drawn on the editor page, not here.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

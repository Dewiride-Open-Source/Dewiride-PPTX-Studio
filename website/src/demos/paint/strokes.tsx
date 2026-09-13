'use client';

import { useState } from 'react';

import {
  EMU_PER_POINT,
  MARKER_TYPE_NAMES,
  PRESET_DASHES,
  PRESET_DASH_NAMES,
  dashArray,
  markerGeometry,
  type LineCap,
  type LineEndSize,
} from '@pptx-studio/paint';

import { Mono } from '@/design/code';
import { Segmented } from '@/design/segmented';
import { Slider } from '@/design/field';
import { Panel } from '@/design/panel';

const CAPS: readonly { value: LineCap; label: string }[] = [
  { value: 'flat', label: 'flat' },
  { value: 'sq', label: 'square' },
  { value: 'rnd', label: 'round' },
];
const SVG_CAP: Record<LineCap, 'butt' | 'square' | 'round'> = {
  flat: 'butt',
  sq: 'square',
  rnd: 'round',
};
const SIZES: readonly LineEndSize[] = ['sm', 'med', 'lg'];

/** The eleven preset dashes under each cap, compensated the way PowerPoint paints them. */
function Dashes() {
  const [cap, setCap] = useState<LineCap>('flat');
  const [width, setWidth] = useState(3);
  return (
    <Panel
      title="Preset dashes"
      hint="A round cap would lengthen every dash by one width, so the array is shortened to match what PowerPoint paints."
      aside={<Segmented label="Line cap" size="sm" options={CAPS} value={cap} onChange={setCap} />}
    >
      <div className="flex flex-col gap-2 p-4">
        <div className="flex items-center gap-2 text-[12px] text-fg-muted">
          width
          <Slider
            ariaLabel="Stroke width"
            min={1}
            max={8}
            value={width}
            onChange={setWidth}
            format={(v) => `${String(v)} px`}
          />
        </div>
        {PRESET_DASH_NAMES.map((name) => {
          const segments = PRESET_DASHES[name] ?? [];
          const array = dashArray(segments, width, cap);
          return (
            <div key={name} className="flex items-center gap-3">
              <span className="w-24 shrink-0 font-mono text-[11px] text-fg-muted">{name}</span>
              <svg viewBox="0 0 240 12" className="h-3 flex-1" aria-hidden="true">
                <line
                  x1={4}
                  y1={6}
                  x2={236}
                  y2={6}
                  stroke="#3b8fd4"
                  strokeWidth={width}
                  strokeLinecap={SVG_CAP[cap]}
                  strokeDasharray={array.join(' ')}
                />
              </svg>
              <Mono tone="dim">{array.map((n) => n.toFixed(1)).join(' ')}</Mono>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

/** Every arrowhead at every size, from markerGeometry, drawn in the same box. */
function Arrowheads() {
  const [size, setSize] = useState<LineEndSize>('med');
  const strokeWidth = 2 * EMU_PER_POINT;
  return (
    <Panel
      title="Arrowheads"
      hint="markerGeometry gives the outline in a unit box, plus how far the tip overshoots the line's end."
      aside={
        <Segmented
          label="Arrowhead size"
          size="sm"
          options={SIZES.map((value) => ({ value, label: value }))}
          value={size}
          onChange={setSize}
        />
      }
    >
      <div className="grid grid-cols-3 gap-3 p-4 sm:grid-cols-6">
        {MARKER_TYPE_NAMES.map((type) => {
          const marker = markerGeometry({ type, w: size, len: size }, strokeWidth);
          const scale = 1 / EMU_PER_POINT;
          const length = marker === null ? 0 : marker.length * scale;
          const breadth = marker === null ? 0 : marker.width * scale;
          return (
            <div key={type} className="flex flex-col items-center gap-1">
              <svg viewBox="0 0 60 24" className="h-8 w-full" aria-hidden="true">
                <line
                  x1={2}
                  y1={12}
                  x2={marker === null ? 58 : 58 - length}
                  y2={12}
                  stroke="#3b8fd4"
                  strokeWidth={2}
                />
                {marker === null ? null : (
                  <g
                    transform={`translate(${String(58 - length)} 12) scale(${String(length)} ${String(breadth)}) translate(0 -0.5)`}
                  >
                    <path
                      d={marker.path}
                      fill={marker.filled ? '#3b8fd4' : 'none'}
                      stroke="#3b8fd4"
                      strokeWidth={2 / Math.max(length, 1)}
                      vectorEffect="non-scaling-stroke"
                    />
                  </g>
                )}
              </svg>
              <span className="font-mono text-[10px] text-fg-muted">{type}</span>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

export function Strokes() {
  return (
    <div className="grid gap-4 lg:grid-cols-2">
      <Dashes />
      <Arrowheads />
    </div>
  );
}

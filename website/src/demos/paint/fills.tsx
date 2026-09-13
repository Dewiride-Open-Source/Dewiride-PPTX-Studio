'use client';

import { useId, useMemo, useState } from 'react';

import {
  PATTERN_TILES,
  PRESET_PATTERN_NAMES,
  RAMP_BLEND_GAMMA,
  TWO_STOP_RAMP,
  fromRampSpace,
  linearGradientVector,
  svgStops,
  toCss,
  toRampSpace,
  type GradientFill,
  type Rgba,
} from '@pptx-studio/paint';

import { Badge } from '@/design/badge';
import { Mono } from '@/design/code';
import { Slider, Switch } from '@/design/field';
import { Panel } from '@/design/panel';

const FROM: Rgba = { r: 68, g: 114, b: 196, a: 1 };
const TO: Rgba = { r: 255, g: 255, b: 255, a: 1 };
const BOX = { w: 320, h: 160 };

/** The measured two-stop curve: each knot mixes the stops in the 2.2 power space by its weight. */
function Ramp() {
  const ramp = useMemo(() => {
    const mix = (from: number, onto: number, weight: number): number =>
      fromRampSpace(toRampSpace(from) + (toRampSpace(onto) - toRampSpace(from)) * weight);
    return TWO_STOP_RAMP.map(([position, weight]) => ({
      position,
      css: toCss({
        r: mix(FROM.r, TO.r, weight),
        g: mix(FROM.g, TO.g, weight),
        b: mix(FROM.b, TO.b, weight),
        a: 1,
      }),
    }));
  }, []);
  return (
    <Panel
      title="The two-stop ramp"
      hint={`Pre-sampled through a gamma of ${String(RAMP_BLEND_GAMMA)}, not blended linearly.`}
    >
      <div className="p-4">
        <div className="flex h-12 overflow-hidden rounded-control border border-line-strong">
          {ramp.map((stop) => (
            <div key={stop.position} className="flex-1" style={{ background: stop.css }} />
          ))}
        </div>
        <div
          className="mt-2 h-12 rounded-control border border-line-strong"
          style={{ background: `linear-gradient(90deg, ${toCss(FROM)}, ${toCss(TO)})` }}
        />
        <p className="mt-2 text-[12px] text-fg-faint">
          Above, PowerPoint&apos;s curve in {ramp.length} knots; below, a plain linear blend of the
          same two colours, which is what a naive gradient renders and why it looks duller.
        </p>
      </div>
    </Panel>
  );
}

/** A linear gradient built the way the renderer builds one: stops from svgStops, a vector from the angle. */
function GradientBuilder() {
  const id = useId();
  const [ang, setAng] = useState(90);
  const [scaled, setScaled] = useState(false);
  const [middle, setMiddle] = useState(50);

  const fill: GradientFill = {
    type: 'gradient',
    stops: [
      { pos: 0, color: { space: 'srgb', hex: '4472C4', transforms: [] } },
      { pos: middle * 1000, color: { space: 'srgb', hex: 'A855C9', transforms: [] } },
      { pos: 100000, color: { space: 'srgb', hex: '00A5B5', transforms: [] } },
    ],
    shade: { kind: 'linear', ang: ang * 60000, scaled },
    tileRect: null,
    flip: 'none',
    rotWithShape: true,
  };
  const stops = svgStops(fill);
  const vector = linearGradientVector({ kind: 'linear', ang: ang * 60000, scaled }, BOX.w, BOX.h);

  return (
    <Panel
      title="A linear gradient"
      hint="svgStops and linearGradientVector are exactly what the renderer writes into <defs>."
    >
      <div className="flex flex-col gap-3 p-4">
        <svg
          viewBox={`0 0 ${String(BOX.w)} ${String(BOX.h)}`}
          className="h-auto w-full max-w-md rounded-control border border-line-strong"
          role="img"
          aria-label="Gradient preview"
        >
          <defs>
            <linearGradient
              id={id}
              gradientUnits="userSpaceOnUse"
              x1={vector.x1}
              y1={vector.y1}
              x2={vector.x2}
              y2={vector.y2}
            >
              {stops.map((stop, at) => (
                <stop
                  key={at}
                  offset={stop.offset}
                  stopColor={stop.color}
                  stopOpacity={stop.opacity}
                />
              ))}
            </linearGradient>
          </defs>
          <rect x={0} y={0} width={BOX.w} height={BOX.h} fill={`url(#${id})`} />
          <line
            x1={vector.x1}
            y1={vector.y1}
            x2={vector.x2}
            y2={vector.y2}
            stroke="#fff"
            strokeWidth={1.5}
            strokeDasharray="4 3"
          />
        </svg>
        <div className="flex flex-wrap items-center gap-4 text-[13px] text-fg-muted">
          <span className="inline-flex items-center gap-2">
            angle
            <Slider
              ariaLabel="Gradient angle"
              min={0}
              max={359}
              value={ang}
              onChange={setAng}
              format={(v) => `${String(v)}°`}
            />
          </span>
          <span className="inline-flex items-center gap-2">
            middle stop
            <Slider
              ariaLabel="Middle stop position"
              min={1}
              max={99}
              value={middle}
              onChange={setMiddle}
              format={(v) => `${String(v)}%`}
            />
          </span>
          <Switch checked={scaled} onChange={setScaled}>
            <Mono>scaled</Mono>
          </Switch>
        </div>
        <p className="text-[12px] text-fg-faint">
          The dashed line is the vector:{' '}
          <Mono tone="dim">{`${vector.x1.toFixed(0)},${vector.y1.toFixed(0)} → ${vector.x2.toFixed(0)},${vector.y2.toFixed(0)}`}</Mono>
          . With <Mono tone="dim">scaled</Mono> the angle is measured in a space where the box is a
          square, which is how PowerPoint bends it on a wide shape.
        </p>
      </div>
    </Panel>
  );
}

/** One 8x8 pattern tile, drawn from its own coverage string. */
function Tile({ name }: { name: string }) {
  const tile = PATTERN_TILES[name];
  if (tile === undefined) return null;
  const cells: string[] = [];
  for (let at = 0; at < tile.cover.length; at += 2) cells.push(tile.cover.slice(at, at + 2));
  return (
    <svg
      viewBox={`0 0 ${String(tile.w)} ${String(tile.h)}`}
      className="h-10 w-10"
      shapeRendering="crispEdges"
      aria-hidden="true"
    >
      <rect x={0} y={0} width={tile.w} height={tile.h} className="fill-sunken" />
      {cells.map((cell, at) => {
        const alpha = parseInt(cell, 16) / 255;
        if (alpha === 0) return null;
        return (
          <rect
            key={at}
            x={at % tile.w}
            y={Math.floor(at / tile.w)}
            width={1}
            height={1}
            fill="#3b8fd4"
            fillOpacity={alpha}
          />
        );
      })}
    </svg>
  );
}

export function Fills() {
  return (
    <>
      <div className="grid gap-4 lg:grid-cols-2">
        <Ramp />
        <GradientBuilder />
      </div>
      <Panel
        title="Pattern tiles"
        hint="Physical 8 by 8 pixel tiles, transcoded from Mono's libgdiplus; a pattern does not scale with the shape."
        aside={<Badge tone="info">{PRESET_PATTERN_NAMES.length}</Badge>}
      >
        <div className="grid grid-cols-4 gap-3 p-4 sm:grid-cols-6 lg:grid-cols-9">
          {PRESET_PATTERN_NAMES.map((name) => (
            <div key={name} className="flex flex-col items-center gap-1">
              <Tile name={name} />
              <span className="w-full truncate text-center font-mono text-[9px] text-fg-faint">
                {name}
              </span>
            </div>
          ))}
        </div>
      </Panel>
    </>
  );
}

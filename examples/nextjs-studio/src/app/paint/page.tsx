'use client';

import { useMemo, useState } from 'react';

import {
  PATTERN_TILES,
  TWO_STOP_RAMP,
  fromRampSpace,
  toRampSpace,
  PRESET_DASHES,
  PRESET_DASH_NAMES,
  PRESET_PATTERN_NAMES,
  RAMP_BLEND_GAMMA,
  applyTransforms,
  parseHex,
  toCss,
  type ColorTransform,
  type PercentTransformOp,
  type Rgba,
} from '@pptx-studio/paint';

import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';

const HOW = `import { applyTransforms, toCss } from '@pptx-studio/paint';

// Document order is load-bearing. On #4472C4:
//   lumMod 60% then lumOff 40%  ->  8FAADC   (PowerPoint's "Lighter 40%")
//   lumOff 40% then lumMod 60%  ->  517CC8   (wrong, and plausible)
const out = applyTransforms(base, [
  { op: 'lumMod', val: 60000 },   // hundred-thousandths, as the file writes it
  { op: 'lumOff', val: 40000 },
]);

toCss(out);   // 'rgb(143 170 220)'`;

/** The transforms worth putting on a slider, with the range PowerPoint uses. */
const OPS: readonly PercentTransformOp[] = ['tint', 'shade', 'lumMod', 'lumOff', 'satMod', 'alpha'];

const START = '#4472C4';

function swatchOf(base: Rgba, applied: readonly ColorTransform[]): string {
  try {
    return toCss(applyTransforms(base, applied));
  } catch {
    return 'transparent';
  }
}

/** One 8x8 pattern tile, drawn from its own coverage string. */
function Tile({ name }: { name: string }) {
  const tile = PATTERN_TILES[name];
  if (tile === undefined) return null;
  const cells: string[] = [];
  for (let at = 0; at < tile.cover.length; at += 2) {
    cells.push(tile.cover.slice(at, at + 2));
  }
  return (
    <svg viewBox={`0 0 ${tile.w} ${tile.h}`} className="h-10 w-10" shapeRendering="crispEdges">
      <rect x={0} y={0} width={tile.w} height={tile.h} fill="#12151a" />
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

export default function PaintPage() {
  const [hex, setHex] = useState(START);
  const [applied, setApplied] = useState<readonly ColorTransform[]>([
    { op: 'lumMod', val: 60000 },
    { op: 'lumOff', val: 40000 },
  ]);

  const base = useMemo<Rgba>(() => {
    const parsed = parseHex(hex);
    return parsed === null
      ? { r: 68, g: 114, b: 196, a: 1 }
      : { r: parsed[0], g: parsed[1], b: parsed[2], a: 1 };
  }, [hex]);

  const forward = swatchOf(base, applied);
  const reversed = swatchOf(base, [...applied].reverse());
  // The measured curve, sampled: at each position the colour is the two stops
  // mixed in the 2.2 power space by that position's weight.
  const ramp = useMemo(() => {
    const to: Rgba = { r: 255, g: 255, b: 255, a: 1 };
    const mix = (from: number, onto: number, weight: number): number =>
      fromRampSpace(toRampSpace(from) + (toRampSpace(onto) - toRampSpace(from)) * weight);
    return TWO_STOP_RAMP.map(([position, weight]) => ({
      position,
      color: {
        r: mix(base.r, to.r, weight),
        g: mix(base.g, to.g, weight),
        b: mix(base.b, to.b, weight),
        a: 1,
      } satisfies Rgba,
    }));
  }, [base]);

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Paint</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/paint</Mono> - colour transforms, gradients, patterns and dashes.
        </p>
      </header>

      <Panel
        title="Transforms apply in document order"
        hint="Reorder them and the answer changes. This is the classic wrong-shade bug."
      >
        <div className="flex flex-col gap-4 p-4">
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-[13px] text-ink-400">
              base
              <input
                type="color"
                value={hex}
                onChange={(event) => setHex(event.target.value)}
                className="h-8 w-12 rounded border border-ink-600 bg-ink-800"
              />
              <Mono>{hex.toUpperCase()}</Mono>
            </label>
            <div className="flex flex-wrap gap-1.5">
              {OPS.map((op) => (
                <button
                  key={op}
                  type="button"
                  onClick={() =>
                    setApplied((was) =>
                      was.some((one) => one.op === op)
                        ? was.filter((one) => one.op !== op)
                        : [...was, { op, val: 40000 }],
                    )
                  }
                  className={`rounded border px-2 py-1 font-mono text-[11px] ${
                    applied.some((one) => one.op === op)
                      ? 'border-chrome bg-chrome/10 text-chrome'
                      : 'border-ink-600 bg-ink-800 text-ink-300'
                  }`}
                >
                  {op}
                </button>
              ))}
            </div>
          </div>

          {applied.length === 0 ? (
            <p className="text-[13px] text-ink-400">
              No transform - the base colour is the answer.
            </p>
          ) : (
            <ol className="flex flex-col gap-2">
              {applied.map((one, at) => (
                <li key={`${one.op}-${String(at)}`} className="flex items-center gap-3">
                  <span className="w-5 text-right font-mono text-[11px] text-ink-500">
                    {at + 1}
                  </span>
                  <Mono>{one.op}</Mono>
                  {'val' in one ? (
                    <>
                      <input
                        type="range"
                        min={0}
                        max={100000}
                        step={1000}
                        value={one.val}
                        onChange={(event) =>
                          setApplied((was) =>
                            was.map((each, index) =>
                              index === at ? { ...each, val: Number(event.target.value) } : each,
                            ),
                          )
                        }
                        className="w-40 accent-chrome"
                      />
                      <span className="tabular text-[12px] text-ink-400">
                        {(one.val / 1000).toFixed(0)}%
                      </span>
                    </>
                  ) : null}
                  <button
                    type="button"
                    onClick={() => setApplied((was) => was.filter((_, index) => index !== at))}
                    className="ml-auto text-[11px] text-ink-500 hover:text-handle"
                  >
                    remove
                  </button>
                </li>
              ))}
            </ol>
          )}

          <div className="grid gap-3 sm:grid-cols-3">
            <div>
              <div
                className="h-16 rounded border border-ink-600"
                style={{ background: toCss(base) }}
              />
              <p className="mt-1.5 text-[12px] text-ink-400">base</p>
            </div>
            <div>
              <div className="h-16 rounded border border-chrome" style={{ background: forward }} />
              <p className="mt-1.5 text-[12px] text-ink-200">
                in order <Mono>{forward}</Mono>
              </p>
            </div>
            <div>
              <div
                className="h-16 rounded border border-ink-600"
                style={{ background: reversed }}
              />
              <p className="mt-1.5 text-[12px] text-ink-400">
                reversed <Mono tone="dim">{reversed}</Mono>
              </p>
            </div>
          </div>
          {forward === reversed && applied.length > 1 ? (
            <p className="text-[12px] text-ink-500">
              These two commute. Swap in <Mono tone="dim">lumMod</Mono> and{' '}
              <Mono tone="dim">lumOff</Mono>, or <Mono tone="dim">tint</Mono> and{' '}
              <Mono tone="dim">shade</Mono>, and they do not.
            </p>
          ) : null}
        </div>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-2">
        <Panel
          title="The two-stop ramp"
          hint={`Pre-sampled through a gamma of ${String(RAMP_BLEND_GAMMA)}, not blended linearly.`}
        >
          <div className="p-4">
            <div className="flex h-16 overflow-hidden rounded border border-ink-600">
              {ramp.map((stop) => (
                <div
                  key={stop.position}
                  className="flex-1"
                  style={{ background: toCss(stop.color) }}
                />
              ))}
            </div>
            <p className="mt-2 text-[12px] text-ink-500">
              {ramp.length} stops. A plain two-stop SVG gradient between the same colours renders
              noticeably duller than PowerPoint, and every Office <Mono tone="dim">fmtScheme</Mono>{' '}
              gradient is exactly this shape.
            </p>
          </div>
        </Panel>

        <Panel title="Preset dashes" hint="As line-width multiples, with cap compensation.">
          <div className="flex flex-col gap-2 p-4">
            {PRESET_DASH_NAMES.map((name) => {
              const segments = PRESET_DASHES[name] ?? [];
              return (
                <div key={name} className="flex items-center gap-3">
                  <span className="w-24 shrink-0 font-mono text-[11px] text-ink-400">{name}</span>
                  <svg viewBox="0 0 120 6" className="h-2 flex-1">
                    <line
                      x1={0}
                      y1={3}
                      x2={120}
                      y2={3}
                      stroke="#3b8fd4"
                      strokeWidth={3}
                      strokeDasharray={segments.flatMap((one) => [one.d * 3, one.sp * 3]).join(' ')}
                    />
                  </svg>
                </div>
              );
            })}
          </div>
        </Panel>
      </div>

      <Panel
        title="Pattern tiles"
        hint="Codegen'd from Mono libgdiplus (MIT), not Wine (LGPL)."
        aside={<Badge tone="info">{PRESET_PATTERN_NAMES.length}</Badge>}
      >
        <div className="grid grid-cols-4 gap-3 p-4 sm:grid-cols-6 lg:grid-cols-9">
          {PRESET_PATTERN_NAMES.map((name) => (
            <div key={name} className="flex flex-col items-center gap-1">
              <Tile name={name} />
              <span className="w-full truncate text-center font-mono text-[9px] text-ink-500">
                {name}
              </span>
            </div>
          ))}
        </div>
      </Panel>

      <Panel>
        <div className="grid grid-cols-3 gap-4 p-4">
          <Stat label="pattern tiles" value={PRESET_PATTERN_NAMES.length} />
          <Stat label="preset dashes" value={PRESET_DASH_NAMES.length} />
          <Stat label="ramp stops" value={ramp.length} />
        </div>
      </Panel>

      <Panel title="How">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

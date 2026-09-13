'use client';

import { useMemo, useState } from 'react';

import { colorMapOf, schemeOf } from '@pptx-studio/model';
import {
  SCHEME_SLOTS,
  applyTransforms,
  parseHex,
  resolveColor,
  toCss,
  toHexColor,
  type Color,
  type ColorTransform,
  type PercentTransformOp,
  type Rgba,
  type SchemeSlot,
} from '@pptx-studio/paint';

import { useDocument } from '@/deck/document';
import { Chip } from '@/design/badge';
import { Mono } from '@/design/code';
import { Select, Slider } from '@/design/field';
import { Panel } from '@/design/panel';

/** The transforms worth putting on a slider, with the range PowerPoint uses. */
const OPS: readonly PercentTransformOp[] = ['tint', 'shade', 'lumMod', 'lumOff', 'satMod', 'alpha'];
const FALLBACK: Rgba = { r: 68, g: 114, b: 196, a: 1 };

function swatchOf(base: Rgba, applied: readonly ColorTransform[]): string {
  return toCss(applyTransforms(base, applied));
}

/** A base colour - typed, or a scheme slot of the open deck's theme - through transforms in order. */
export function Colours() {
  const { parsed } = useDocument();
  const [hex, setHex] = useState('#4472C4');
  const [slot, setSlot] = useState<SchemeSlot | ''>('');
  const [applied, setApplied] = useState<readonly ColorTransform[]>([
    { op: 'lumMod', val: 60000 },
    { op: 'lumOff', val: 40000 },
  ]);

  const theme = useMemo(() => {
    const slide = parsed?.document.slides[0];
    if (slide === undefined) return null;
    const scheme = schemeOf(slide);
    const map = colorMapOf(slide);
    return scheme === undefined ? null : { scheme, map, name: slide.theme?.name ?? 'theme' };
  }, [parsed]);

  const base = useMemo<Rgba>(() => {
    if (slot !== '' && theme !== null) {
      const color: Color = { space: 'scheme', name: slot, transforms: [] };
      return resolveColor(color, { scheme: theme.scheme, map: theme.map });
    }
    const typed = parseHex(hex);
    return typed === null ? FALLBACK : { r: typed[0], g: typed[1], b: typed[2], a: 1 };
  }, [hex, slot, theme]);

  const forward = swatchOf(base, applied);
  const reversed = swatchOf(base, [...applied].reverse());

  return (
    <Panel
      title="Transforms apply in document order"
      hint="Reorder them and the answer changes. This is the classic wrong-shade bug."
    >
      <div className="flex flex-col gap-4 p-4">
        <div className="flex flex-wrap items-center gap-3">
          <label className="flex items-center gap-2 text-[13px] text-fg-muted">
            base
            <input
              type="color"
              value={hex}
              onChange={(event) => {
                setHex(event.target.value);
                setSlot('');
              }}
              className="h-8 w-12 rounded-control border border-line-strong bg-surface"
            />
            <Mono>{toHexColor(base).toUpperCase()}</Mono>
          </label>
          {theme === null ? null : (
            <label className="flex items-center gap-2 text-[13px] text-fg-muted">
              or a slot of <span className="text-fg">{theme.name}</span>
              <Select
                value={slot}
                onChange={(event) => setSlot(event.target.value as SchemeSlot | '')}
              >
                <option value="">typed colour</option>
                {SCHEME_SLOTS.map((one) => (
                  <option key={one} value={one}>
                    {one}
                  </option>
                ))}
              </Select>
            </label>
          )}
          <div className="flex flex-wrap gap-1.5">
            {OPS.map((op) => (
              <Chip
                key={op}
                selected={applied.some((one) => one.op === op)}
                onClick={() =>
                  setApplied((was) =>
                    was.some((one) => one.op === op)
                      ? was.filter((one) => one.op !== op)
                      : [...was, { op, val: 40000 }],
                  )
                }
              >
                {op}
              </Chip>
            ))}
          </div>
        </div>

        {applied.length === 0 ? (
          <p className="text-[13px] text-fg-muted">No transform - the base colour is the answer.</p>
        ) : (
          <ol className="flex flex-col gap-2">
            {applied.map((one, at) => (
              <li key={`${one.op}-${String(at)}`} className="flex flex-wrap items-center gap-3">
                <span className="w-5 text-right font-mono text-[11px] text-fg-faint">{at + 1}</span>
                <Mono>{one.op}</Mono>
                {'val' in one ? (
                  <Slider
                    ariaLabel={`${one.op} value`}
                    min={0}
                    max={100000}
                    step={1000}
                    value={one.val}
                    format={(value) => `${(value / 1000).toFixed(0)}%`}
                    onChange={(value) =>
                      setApplied((was) =>
                        was.map((each, index) => (index === at ? { ...each, val: value } : each)),
                      )
                    }
                  />
                ) : null}
                <span className="inline-flex gap-1">
                  <button
                    type="button"
                    aria-label="Move up"
                    disabled={at === 0}
                    onClick={() =>
                      setApplied((was) => {
                        const next = [...was];
                        [next[at - 1], next[at]] = [next[at]!, next[at - 1]!];
                        return next;
                      })
                    }
                    className="text-[11px] text-fg-faint hover:text-fg disabled:opacity-30"
                  >
                    ↑
                  </button>
                  <button
                    type="button"
                    aria-label="Move down"
                    disabled={at === applied.length - 1}
                    onClick={() =>
                      setApplied((was) => {
                        const next = [...was];
                        [next[at], next[at + 1]] = [next[at + 1]!, next[at]!];
                        return next;
                      })
                    }
                    className="text-[11px] text-fg-faint hover:text-fg disabled:opacity-30"
                  >
                    ↓
                  </button>
                </span>
                <button
                  type="button"
                  onClick={() => setApplied((was) => was.filter((_, index) => index !== at))}
                  className="ml-auto text-[11px] text-fg-faint hover:text-bad"
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
              className="h-16 rounded-control border border-line-strong"
              style={{ background: toCss(base) }}
            />
            <p className="mt-1.5 text-[12px] text-fg-muted">base</p>
          </div>
          <div>
            <div
              className="h-16 rounded-control border border-accent"
              style={{ background: forward }}
            />
            <p className="mt-1.5 text-[12px] text-fg">
              in order <Mono>{forward}</Mono>
            </p>
          </div>
          <div>
            <div
              className="h-16 rounded-control border border-line-strong"
              style={{ background: reversed }}
            />
            <p className="mt-1.5 text-[12px] text-fg-muted">
              reversed <Mono tone="dim">{reversed}</Mono>
            </p>
          </div>
        </div>
        {forward === reversed && applied.length > 1 ? (
          <p className="text-[12px] text-fg-faint">
            These two commute. <Mono tone="dim">lumMod</Mono> then <Mono tone="dim">lumOff</Mono>,
            or <Mono tone="dim">tint</Mono> then <Mono tone="dim">shade</Mono>, do not.
          </p>
        ) : null}
      </div>
    </Panel>
  );
}

'use client';

import { useMemo, useState } from 'react';

import {
  inheritanceChain,
  resolve,
  resolveAppearance,
  resolveBackground,
  resolveXfrm,
  type Origin,
  type Shape,
  type Sheet,
} from '@pptx-studio/model';
import { flatten, layoutSheet } from '@pptx-studio/render-svg';

import { useDocument } from '@/deck/document';
import { Badge, type Tone } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel } from '@/shell/panel';

const HOW = `import { resolve, inheritanceChain } from '@pptx-studio/model';

// Not "what is the fill" but "what is the fill, and who said so".
const fill = resolve(shape, sheet, (s) => s.fill);
fill?.value;      // the fill
fill?.origin;     // 'shape' | 'layoutPh' | 'masterPh' | 'theme' | ...
fill?.explicit;   // did the thing it came from declare it, or inherit it too

// undefined means the file did not say - and that is the fact
// Change Layout, provenance and theme swapping are all built on.
shape.xfrm === undefined;`;

/** Where a value came from, said in words rather than in the enum's spelling. */
const ORIGINS: Record<Origin, { label: string; tone: Tone }> = {
  run: { label: 'the run', tone: 'good' },
  paragraph: { label: 'the paragraph', tone: 'good' },
  shape: { label: 'the shape', tone: 'good' },
  layoutPh: { label: 'the layout placeholder', tone: 'info' },
  masterPh: { label: 'the master placeholder', tone: 'info' },
  txStyles: { label: 'the master text styles', tone: 'info' },
  defaultTextStyle: { label: 'p:defaultTextStyle', tone: 'warn' },
  builtin: { label: "PowerPoint's built-in styles", tone: 'warn' },
  theme: { label: 'the theme', tone: 'warn' },
  schemaDefault: { label: 'the schema default', tone: 'plain' },
};

function OriginChip({ origin, explicit }: { origin: Origin; explicit: boolean }) {
  const found = ORIGINS[origin];
  return (
    <span className="inline-flex items-center gap-1.5">
      <span
        aria-hidden
        className={`inline-block h-2 w-2 rounded-full ${
          explicit ? 'bg-chrome' : 'border border-ink-400'
        }`}
      />
      <Badge tone={found.tone}>{found.label}</Badge>
    </span>
  );
}

function Property({
  name,
  origin,
  explicit,
  value,
}: {
  name: string;
  origin: Origin | null;
  explicit: boolean;
  value: string;
}) {
  return (
    <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1 border-b border-ink-800 px-4 py-2.5 last:border-0">
      <span className="w-28 shrink-0 font-mono text-[12px] text-ink-400">{name}</span>
      <span className="min-w-0 flex-1 text-[13px] text-ink-100">{value}</span>
      {origin === null ? (
        <Badge tone="plain">nothing declares it</Badge>
      ) : (
        <OriginChip origin={origin} explicit={explicit} />
      )}
    </div>
  );
}

export default function InheritancePage() {
  const { parsed, error, loading } = useDocument();
  const [at, setAt] = useState(0);
  const [selected, setSelected] = useState<number | null>(null);

  const slide: Sheet | undefined = parsed?.document.slides[at];

  const shapes = useMemo(() => {
    if (slide === undefined) return [];
    // The sheet's own shapes, not the inherited furniture: inheritance is a
    // question about what this slide declares.
    return flatten(layoutSheet(slide)).map((placed) => placed.shape);
  }, [slide]);

  if (loading) return <p className="text-sm text-ink-400">Reading the deck…</p>;
  if (error !== null) return <p className="text-sm text-handle">{error}</p>;
  if (parsed === null || slide === undefined) return null;

  const shape: Shape | undefined = shapes.find((one) => one.cNvPrId === selected) ?? shapes[0];

  const chain = shape === undefined ? [] : inheritanceChain(shape, slide);
  const xfrm = shape === undefined ? undefined : resolveXfrm(shape, slide);
  const appearance = shape === undefined ? null : resolveAppearance(shape, slide);
  const background = resolveBackground(slide);
  const geometry = shape === undefined ? undefined : resolve(shape, slide, (one) => one.prstGeom);

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Inheritance</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/model</Mono> - not what a property is, but where it came from.
        </p>
      </header>

      <div className="flex flex-wrap gap-2">
        {parsed.document.slides.map((one, index) => (
          <button
            key={one.partName}
            type="button"
            onClick={() => {
              setAt(index);
              setSelected(null);
            }}
            className={`rounded border px-2 py-1 text-[12px] ${
              index === at
                ? 'border-chrome bg-chrome/10 text-chrome'
                : 'border-ink-700 bg-ink-850 text-ink-300 hover:border-ink-500'
            }`}
          >
            {index + 1}
          </button>
        ))}
      </div>

      <div className="grid gap-4 lg:grid-cols-[minmax(0,18rem)_1fr]">
        <Panel title="Shapes on this slide">
          {shapes.length === 0 ? (
            <p className="p-4 text-[13px] text-ink-400">
              This slide declares no shapes of its own. Everything visible on it comes from the
              layout and the master.
            </p>
          ) : (
            <div className="max-h-96 overflow-y-auto p-2">
              {shapes.map((one) => (
                <button
                  key={one.cNvPrId}
                  type="button"
                  onClick={() => setSelected(one.cNvPrId)}
                  className={`block w-full rounded px-2 py-1.5 text-left ${
                    one.cNvPrId === shape?.cNvPrId ? 'bg-chrome/10' : 'hover:bg-ink-800'
                  }`}
                >
                  <span className="block text-[13px] text-ink-100">
                    {one.name === '' ? `#${String(one.cNvPrId)}` : one.name}
                  </span>
                  <span className="block font-mono text-[11px] text-ink-500">
                    {one.kind}
                    {one.placeholder === null
                      ? ''
                      : ` · ph ${one.placeholder.type}/${String(one.placeholder.idx)}`}
                  </span>
                </button>
              ))}
            </div>
          )}
        </Panel>

        <div className="flex flex-col gap-4">
          <Panel
            title="The chain"
            hint="Slide to layout matches on (type, idx); layout to master on the type alone."
          >
            {chain.length === 0 ? (
              <p className="p-4 text-[13px] text-ink-400">Select a shape.</p>
            ) : (
              <ol className="p-4">
                {chain.map((link, index) => (
                  <li key={index} className="flex items-baseline gap-3 py-1.5">
                    <Badge tone={index === 0 ? 'good' : 'info'}>{link.origin}</Badge>
                    <span className="text-[13px] text-ink-200">
                      {link.shape.name === '' ? `#${String(link.shape.cNvPrId)}` : link.shape.name}
                    </span>
                    <Mono tone="dim">{link.sheet.partName}</Mono>
                  </li>
                ))}
                {chain.length === 1 ? (
                  <li className="pt-2 text-[12px] text-ink-500">
                    One link only - this shape is not a placeholder, or it matched nothing in the
                    layout.
                  </li>
                ) : null}
              </ol>
            )}
          </Panel>

          <Panel
            title="Resolved properties"
            hint="A filled dot means the thing it came from declared it; hollow means it inherited it too."
          >
            {shape === undefined ? (
              <p className="p-4 text-[13px] text-ink-400">Select a shape.</p>
            ) : (
              <div>
                <Property
                  name="xfrm"
                  origin={xfrm?.origin ?? null}
                  explicit={xfrm?.explicit ?? false}
                  value={
                    xfrm === undefined
                      ? 'none'
                      : `${String(Math.round(xfrm.value.x / 12700))}, ` +
                        `${String(Math.round(xfrm.value.y / 12700))} pt   ` +
                        `${String(Math.round(xfrm.value.cx / 12700))} x ` +
                        `${String(Math.round(xfrm.value.cy / 12700))} pt`
                  }
                />
                <Property
                  name="prstGeom"
                  origin={geometry?.origin ?? null}
                  explicit={geometry?.explicit ?? false}
                  value={geometry?.value ?? 'none or a:custGeom'}
                />
                <Property
                  name="fill"
                  origin={appearance?.fillOrigin ?? null}
                  explicit={appearance?.fillOrigin === 'shape'}
                  value={
                    appearance?.fill === null ? 'nothing painted' : (appearance?.fill.type ?? '-')
                  }
                />
                <Property
                  name="line"
                  origin={appearance?.lineOrigin ?? null}
                  explicit={appearance?.lineOrigin === 'shape'}
                  value={appearance?.line === null ? 'no outline' : 'declared'}
                />
                <Property
                  name="background"
                  origin={background === null ? null : 'shape'}
                  explicit={background !== null}
                  value={
                    background === null
                      ? 'no sheet in the chain declares one'
                      : `${background.fill.type}, from ${background.sheet.kind}`
                  }
                />
              </div>
            )}
          </Panel>

          {shape !== undefined && shape.xfrm === undefined ? (
            <div className="rounded-lg border border-chrome/40 bg-chrome/10 p-4 text-[13px] leading-relaxed text-ink-200">
              This shape has <Mono>xfrm === undefined</Mono>. That is not missing data - it is the
              fact that makes Change Layout possible. Flatten it to a number now and the information
              that it never had geometry is gone, along with the ability to rebind it to a different
              layout later.
            </div>
          ) : null}

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

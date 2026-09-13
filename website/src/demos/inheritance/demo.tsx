'use client';

import { useMemo, useState } from 'react';

import {
  inheritanceChain,
  resolve,
  resolveAppearance,
  resolveBackground,
  resolveBulletKind,
  resolveLatinTypeface,
  resolveSize,
  resolveXfrm,
  sheetChain,
  type Origin,
  type Shape,
  type Sheet,
} from '@pptx-studio/model';
import { flatten, layoutSheet, layoutSlide } from '@pptx-studio/render-svg';

import { useDocument } from '@/deck/document';
import { Badge, type Tone } from '@/design/badge';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Switch } from '@/design/field';
import { Panel } from '@/design/panel';
import { Segmented } from '@/design/segmented';
import { ErrorState, Status } from '@/design/state';
import type { DemoProps } from '../registry';

const HOW = `import { resolve, resolveXfrm, inheritanceChain, resolveSize } from '@pptx-studio/model';

// Not "what is the fill" but "what is the fill, and who said so".
const fill = resolve(shape, sheet, (s) => s.fill);
fill?.value;      // the fill
fill?.origin;     // 'shape' | 'layoutPh' | 'masterPh' | 'theme' | ...
fill?.explicit;   // did the thing it came from declare it, or inherit it too

// The text cascade has ten sources; the answer names the one that spoke.
resolveSize({ sheet, shape, defaultTextStyle }, paragraph, run).origin;   // 'masterPh', say

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
        aria-hidden="true"
        className={`inline-block h-2 w-2 rounded-full ${explicit ? 'bg-accent' : 'border border-fg-muted'}`}
      />
      <Badge tone={found.tone}>{found.label}</Badge>
    </span>
  );
}

interface PropertyProps {
  readonly name: string;
  readonly origin: Origin | null;
  readonly explicit: boolean;
  readonly value: string;
}

function Property({ name, origin, explicit, value }: PropertyProps) {
  return (
    <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1 border-b border-line/60 px-4 py-2.5 last:border-0">
      <span className="w-28 shrink-0 font-mono text-[12px] text-fg-muted">{name}</span>
      <span className="min-w-0 flex-1 text-[13px] text-fg">{value}</span>
      {origin === null ? (
        <Badge>nothing declares it</Badge>
      ) : (
        <OriginChip origin={origin} explicit={explicit} />
      )}
    </div>
  );
}

const pt = (emu: number): string => String(Math.round(emu / 12700));

export default function InheritanceDemo({ full }: DemoProps) {
  const { parsed, failure, loading } = useDocument();
  const [at, setAt] = useState(0);
  const [selected, setSelected] = useState<number | null>(null);
  const [inherited, setInherited] = useState(false);

  const slide: Sheet | undefined = parsed?.document.slides[at];

  const shapes = useMemo(() => {
    if (slide === undefined) return [];
    // The sheet's own shapes by default: inheritance is a question about what this slide declares.
    return flatten(inherited ? layoutSlide(slide) : layoutSheet(slide)).map((placed) => ({
      shape: placed.shape,
      sheet: placed.sheet,
    }));
  }, [slide, inherited]);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (parsed === null || slide === undefined) return null;

  const found = shapes.find((one) => one.shape.cNvPrId === selected) ?? shapes[0];
  const shape: Shape | undefined = found?.shape;
  const on: Sheet = found?.sheet ?? slide;

  const chain = shape === undefined ? [] : inheritanceChain(shape, on);
  const sheets = sheetChain(slide);
  const xfrm = shape === undefined ? undefined : resolveXfrm(shape, on);
  const appearance = shape === undefined ? null : resolveAppearance(shape, on);
  const background = resolveBackground(slide);
  const geometry = shape === undefined ? undefined : resolve(shape, on, (one) => one.prstGeom);
  const paragraph = shape?.text?.paragraphs[0];
  const context =
    shape === undefined
      ? null
      : { sheet: on, shape, defaultTextStyle: parsed.document.defaultTextStyle };
  const run = paragraph?.content.find((one) => one.kind === 'run');
  const size =
    context === null || paragraph === undefined ? undefined : resolveSize(context, paragraph, run);
  const face =
    context === null || paragraph === undefined
      ? undefined
      : resolveLatinTypeface(context, paragraph, run);
  const bullet =
    context === null || paragraph === undefined ? undefined : resolveBulletKind(context, paragraph);

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center gap-3">
        <Segmented
          label="Slide"
          size="sm"
          options={parsed.document.slides.map((one, index) => ({
            value: String(index),
            label: String(index + 1),
          }))}
          value={String(at)}
          onChange={(value) => {
            setAt(Number(value));
            setSelected(null);
          }}
        />
        <Switch checked={inherited} onChange={setInherited}>
          show the layout and master furniture too
        </Switch>
      </div>

      <Panel
        title="The sheets this slide resolves through"
        hint="Each binding comes from the part's own .rels, never from a name."
      >
        <ol className="flex flex-wrap items-center gap-2 p-4 text-[12px]">
          {sheets.map((sheet, index) => (
            <li key={sheet.partName} className="flex items-center gap-2">
              {index === 0 ? null : <span className="text-fg-faint">→</span>}
              <span className="rounded-control border border-line bg-sunken px-2 py-1">
                <Badge tone={sheet.kind === 'slide' ? 'good' : 'info'}>{sheet.kind}</Badge>{' '}
                <Mono tone="dim">{sheet.partName}</Mono>
              </span>
            </li>
          ))}
          {slide.theme === null ? null : (
            <li className="flex items-center gap-2">
              <span className="text-fg-faint">→</span>
              <span className="rounded-control border border-line bg-sunken px-2 py-1">
                <Badge tone="warn">theme</Badge> <span className="text-fg">{slide.theme.name}</span>
              </span>
            </li>
          )}
        </ol>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[minmax(0,18rem)_1fr]">
        <Panel title={inherited ? 'Shapes drawn on this slide' : 'Shapes this slide declares'}>
          {shapes.length === 0 ? (
            <p className="p-4 text-[13px] text-fg-muted">
              This slide declares no shapes of its own. Everything visible on it comes from the
              layout and the master.
            </p>
          ) : (
            <div className="max-h-96 overflow-y-auto p-2" role="listbox" aria-label="Shapes">
              {shapes.map((one) => (
                <button
                  key={`${one.sheet.partName}-${String(one.shape.cNvPrId)}`}
                  type="button"
                  role="option"
                  aria-selected={one.shape === shape}
                  onClick={() => setSelected(one.shape.cNvPrId)}
                  className={`block w-full rounded-control px-2 py-1.5 text-left ${
                    one.shape === shape ? 'bg-accent-soft' : 'hover:bg-sunken'
                  }`}
                >
                  <span className="block text-[13px] text-fg">
                    {one.shape.name === '' ? `#${String(one.shape.cNvPrId)}` : one.shape.name}
                  </span>
                  <span className="block font-mono text-[11px] text-fg-faint">
                    {one.shape.kind}
                    {one.shape.placeholder === null
                      ? ''
                      : ` · ph ${one.shape.placeholder.type}/${String(one.shape.placeholder.idx)}`}
                    {one.sheet.kind === 'slide' ? '' : ` · from the ${one.sheet.kind}`}
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
              <p className="p-4 text-[13px] text-fg-muted">Select a shape.</p>
            ) : (
              <ol className="p-4">
                {chain.map((link, index) => (
                  <li key={index} className="flex flex-wrap items-baseline gap-3 py-1.5">
                    <Badge tone={index === 0 ? 'good' : 'info'}>{link.origin}</Badge>
                    <span className="text-[13px] text-fg">
                      {link.shape.name === '' ? `#${String(link.shape.cNvPrId)}` : link.shape.name}
                    </span>
                    <Mono tone="dim">{link.sheet.partName}</Mono>
                    {link.shape.placeholder === null ? null : (
                      <span className="text-[11px] text-fg-faint">
                        matched by{' '}
                        {index === 1
                          ? `idx ${String(link.shape.placeholder.idx)}`
                          : index === 2
                            ? `type ${link.shape.placeholder.type}`
                            : 'itself'}
                      </span>
                    )}
                  </li>
                ))}
                {chain.length === 1 ? (
                  <li className="pt-2 text-[12px] text-fg-faint">
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
              <p className="p-4 text-[13px] text-fg-muted">Select a shape.</p>
            ) : (
              <div>
                <Property
                  name="xfrm"
                  origin={xfrm?.origin ?? null}
                  explicit={xfrm?.explicit ?? false}
                  value={
                    xfrm === undefined
                      ? 'none'
                      : `${pt(xfrm.value.x)}, ${pt(xfrm.value.y)} pt · ${pt(xfrm.value.cx)} × ${pt(xfrm.value.cy)} pt`
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
                  explicit={background?.explicit ?? false}
                  value={
                    background === null
                      ? 'no sheet in the chain declares one'
                      : `${background.fill.type}, from the ${background.sheet.kind} (${background.form})`
                  }
                />
                {size === undefined || face === undefined ? null : (
                  <>
                    <Property
                      name="text size"
                      origin={size.origin}
                      explicit={size.explicit}
                      value={`${(size.value / 100).toFixed(0)} pt`}
                    />
                    <Property
                      name="typeface"
                      origin={face.origin}
                      explicit={face.explicit}
                      value={face.value}
                    />
                    <Property
                      name="bullet"
                      origin={bullet?.origin ?? null}
                      explicit={bullet?.explicit ?? false}
                      value={bullet?.value ?? 'none'}
                    />
                  </>
                )}
              </div>
            )}
          </Panel>

          {shape !== undefined && shape.xfrm === undefined ? (
            <Callout kind="note" title="xfrm === undefined">
              <p>
                That is not missing data - it is the fact that makes Change Layout possible. Flatten
                it to a number now and the information that it never had geometry is gone, along
                with the ability to rebind it to a different layout later.
              </p>
            </Callout>
          ) : null}
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
              This page reads the chain; it does not change layouts or swap themes - those verbs are
              phase 7. The text rows show the first paragraph&apos;s first run only.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

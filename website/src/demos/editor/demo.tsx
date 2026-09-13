'use client';

import { useMemo, useRef, useState } from 'react';

import { flatten, layoutSlide, type Placed } from '@pptx-studio/render-svg';

import { useDocument } from '@/deck/document';
import { DeckWorkerError, useDeckWorker } from '@/deck/worker/client';
import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Slider, Switch, TextInput } from '@/design/field';
import { Panel } from '@/design/panel';
import { ErrorState, Status, type Failure } from '@/design/state';
import type { DemoProps } from '../registry';
import { Stage, useStage } from './stage';
import { useEditor } from './use-editor';

const EMU_PER_POINT = 12700;
const pt = (emu: number): string => (emu / EMU_PER_POINT).toFixed(1);
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const SWATCHES = ['#E8453C', '#D98A2B', '#2FA869', '#0B62B0', '#7B1FA2', '#111111', '#FFFFFF'];

const HOW = `import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { slideNode, serializeSvg, flatten, mediaFromStore, createTextEngine } from '@pptx-studio/render-svg';

const store = PartStore.open(bytes);
const doc = loadDocument(store);
const text = createTextEngine({ defaultTextStyle: doc.defaultTextStyle });   // one measurer for the deck

const { node, placed } = slideNode(doc.slides[0], doc.slideSize, { media: mediaFromStore(store), text });
const svg = serializeSvg(node);   // the picture, as a string
const shapes = flatten(placed);   // the same shapes, for hit-testing

// Every edit is an XmlEdit on the slide's own part; applyEdit hands back the inverse.
// Move: two setAttribute on a:off. Recolour: a:srgbClr/@val. Retype: setValue on the first a:t.`;

function SvgSource({ svg, name }: { svg: string; name: string }) {
  const [shown, setShown] = useState(false);
  const download = () => {
    const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${name}.svg`;
    anchor.click();
    URL.revokeObjectURL(url);
  };
  return (
    <Panel
      title="The SVG it emitted"
      hint={`${(svg.length / 1024).toFixed(1)} kB · ${String((svg.match(/<defs/g) ?? []).length)} <defs> · real <text>, never foreignObject`}
      aside={
        <div className="flex gap-2">
          <Button size="sm" onClick={() => setShown((on) => !on)} aria-expanded={shown}>
            {shown ? 'Hide' : 'View'}
          </Button>
          <Button size="sm" onClick={download}>
            Download .svg
          </Button>
        </div>
      }
    >
      {shown ? (
        <div className="max-h-96 overflow-auto p-4">
          <Code
            lang="xml"
            code={
              svg.length > 60_000
                ? `${svg.slice(0, 60_000)}\n<!-- … ${String(svg.length - 60_000)} more characters -->`
                : svg
            }
          />
        </div>
      ) : null}
    </Panel>
  );
}

export default function EditorDemo({ full }: DemoProps) {
  const { parsed, failure, loading, bytes, name } = useDocument();
  const editor = useEditor(bytes);
  const worker = useDeckWorker();
  const [at, setAt] = useState(0);
  const [selected, setSelected] = useState<number | null>(null);
  const [overlay, setOverlay] = useState(true);
  const [withText, setWithText] = useState(true);
  const [zoom, setZoom] = useState(820);
  const [exported, setExported] = useState<{
    url: string;
    rewritten: number;
    streamed: number;
  } | null>(null);
  const [exportFailure, setExportFailure] = useState<Failure | null>(null);
  const drag = useRef<{ id: number; x: number; y: number } | null>(null);

  const document = editor.view?.document ?? parsed?.document ?? null;
  const version = editor.view?.version ?? 0;
  const sheet = document?.slides[Math.min(at, (document?.slides.length ?? 1) - 1)];
  const size = document?.slideSize;
  const placed: readonly Placed[] = useMemo(
    () => (sheet === undefined ? [] : flatten(layoutSlide(sheet))),
    [sheet],
  );
  const text = useMemo(
    () => (parsed === null ? null : withText ? parsed.text : false),
    [parsed, withText],
  );
  const drawn = useStage(sheet, size, text, `e${String(at)}v${String(version)}`, parsed?.media);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (
    parsed === null ||
    document === null ||
    sheet === undefined ||
    size === undefined ||
    text === null
  ) {
    return <p className="text-[13px] text-fg-muted">This package has no slides.</p>;
  }

  const shape = placed.find((one) => one.shape.cNvPrId === selected) ?? null;
  const emuPerPixel = size.cx / zoom;

  const exportEdit = () => {
    const current = worker.current;
    if (current === null || bytes === null) return;
    setExportFailure(null);
    current
      .export(bytes, editor.replaced())
      .then((done) => {
        setExported((was) => {
          if (was !== null) URL.revokeObjectURL(was.url);
          return {
            url: URL.createObjectURL(new Blob([done.bytes], { type: MIME })),
            rewritten: done.rewritten.length,
            streamed: done.streamed,
          };
        });
      })
      .catch((cause: unknown) =>
        setExportFailure(
          cause instanceof DeckWorkerError ? cause.failure : { message: String(cause) },
        ),
      );
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center gap-3">
        <div className="flex gap-1 overflow-x-auto pb-1" role="tablist" aria-label="Slides">
          {document.slides.map((one, index) => (
            <button
              key={one.partName}
              type="button"
              role="tab"
              aria-selected={index === at}
              onClick={() => {
                setAt(index);
                setSelected(null);
              }}
              className={`shrink-0 rounded-control border px-2 py-1 font-mono text-[12px] ${
                index === at
                  ? 'border-accent bg-accent-soft text-fg'
                  : 'border-line text-fg-muted hover:border-line-strong'
              }`}
            >
              {index + 1}
            </button>
          ))}
        </div>
        <span className="ml-auto inline-flex items-center gap-2 text-[12px] text-fg-muted">
          zoom
          <Slider
            ariaLabel="Zoom"
            min={360}
            max={1400}
            step={20}
            value={zoom}
            onChange={setZoom}
            format={(v) => `${String(Math.round((v / 960) * 100))}%`}
          />
        </span>
        <Switch checked={overlay} onChange={setOverlay}>
          overlay
        </Switch>
        <Switch checked={withText} onChange={setWithText}>
          text
        </Switch>
        <Button size="sm" disabled={editor.view?.canUndo !== true} onClick={editor.undo}>
          Undo
        </Button>
        <Button size="sm" disabled={editor.view?.canRedo !== true} onClick={editor.redo}>
          Redo
        </Button>
      </div>

      {editor.view?.refused == null ? null : (
        <Callout kind="warn" title="Refused">
          <p>{editor.view.refused}</p>
        </Callout>
      )}

      <div className="flex flex-wrap items-start gap-4">
        <div className="max-w-full overflow-auto rounded-frame border border-line bg-desk p-3">
          <Stage
            key={`${sheet.partName}-${String(version)}`}
            content={drawn}
            size={size}
            width={zoom}
            selected={selected}
            overlay={overlay}
            onSelect={setSelected}
            onNudge={(id, dx, dy) => editor.move(sheet, id, dx, dy)}
            onShapePointerDown={(one, event) => {
              drag.current = { id: one.shape.cNvPrId, x: event.clientX, y: event.clientY };
              const move = (moveEvent: PointerEvent): void => {
                const from = drag.current;
                if (from === null) return;
                const dx = (moveEvent.clientX - from.x) * emuPerPixel;
                const dy = (moveEvent.clientY - from.y) * emuPerPixel;
                // A drag under three pixels is a click, and turning every click
                // into a no-op edit would fill the undo stack with nothing.
                if (Math.abs(dx) + Math.abs(dy) < emuPerPixel * 3) return;
                drag.current = { id: from.id, x: moveEvent.clientX, y: moveEvent.clientY };
                editor.move(sheet, from.id, dx, dy);
              };
              const up = (): void => {
                drag.current = null;
                window.removeEventListener('pointermove', move);
                window.removeEventListener('pointerup', up);
              };
              window.addEventListener('pointermove', move);
              window.addEventListener('pointerup', up);
            }}
          />
        </div>

        <div className="flex min-w-72 flex-1 flex-col gap-4">
          <Panel
            title="Selection"
            hint={shape === null ? 'Click a shape on the slide.' : undefined}
          >
            {shape === null ? (
              <p className="p-4 text-[13px] text-fg-muted">
                Nothing selected. Clicking picks the innermost shape, so a click inside a group
                selects the child rather than the group.
              </p>
            ) : (
              <dl className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-1.5 p-4 text-[13px]">
                <dt className="text-fg-muted">name</dt>
                <dd className="text-fg">
                  {shape.shape.name === '' ? `#${String(shape.shape.cNvPrId)}` : shape.shape.name}
                </dd>
                <dt className="text-fg-muted">kind</dt>
                <dd className="text-fg">{shape.shape.kind}</dd>
                <dt className="text-fg-muted">position</dt>
                <dd className="tabular text-fg">
                  {pt(shape.frame.x)}, {pt(shape.frame.y)} pt
                </dd>
                <dt className="text-fg-muted">size</dt>
                <dd className="tabular text-fg">
                  {pt(shape.frame.cx)} × {pt(shape.frame.cy)} pt
                </dd>
                {shape.frame.rot === 0 ? null : (
                  <>
                    <dt className="text-fg-muted">rotation</dt>
                    <dd className="tabular text-fg">{shape.frame.rot.toFixed(2)}°</dd>
                  </>
                )}
                {!shape.frame.flipH && !shape.frame.flipV ? null : (
                  <>
                    <dt className="text-fg-muted">mirrored</dt>
                    <dd className="text-fg">
                      {[shape.frame.flipH ? 'H' : '', shape.frame.flipV ? 'V' : '']
                        .filter(Boolean)
                        .join(' + ')}
                    </dd>
                  </>
                )}
                <dt className="text-fg-muted">geometry</dt>
                <dd className="text-fg">{shape.geometry?.name ?? 'custom or none'}</dd>
                <dt className="text-fg-muted">declared on</dt>
                <dd className="text-fg">{shape.sheet.kind}</dd>
              </dl>
            )}
          </Panel>

          {shape === null ? null : (
            <Panel title="Edit" hint="Each one is an XmlEdit that hands back its own inverse.">
              <div className="flex flex-col gap-3 p-4">
                <div>
                  <p className="mb-1.5 text-[12px] text-fg-muted">Fill</p>
                  <div className="flex flex-wrap gap-1.5">
                    {SWATCHES.map((hex) => (
                      <button
                        key={hex}
                        type="button"
                        title={hex}
                        aria-label={`Fill ${hex}`}
                        onClick={() => editor.recolour(sheet, shape.shape.cNvPrId, hex)}
                        style={{ background: hex }}
                        className="h-7 w-7 rounded-control border border-line-strong"
                      />
                    ))}
                  </div>
                </div>
                <label className="block">
                  <span className="mb-1.5 block text-[12px] text-fg-muted">First run</span>
                  <TextInput
                    placeholder="type, then press Enter"
                    className="w-full"
                    onKeyDown={(event) => {
                      if (event.key !== 'Enter') return;
                      editor.retext(sheet, shape.shape.cNvPrId, event.currentTarget.value);
                    }}
                  />
                </label>
                <p className="text-[12px] text-fg-faint">
                  Drag the shape on the slide to move it, or nudge with the arrow keys once it is
                  selected. A placeholder with no <Mono tone="dim">a:xfrm</Mono> refuses, and says
                  why.
                </p>
              </div>
            </Panel>
          )}

          <Panel
            title="Session"
            aside={
              editor.view === null ? null : (
                <Button size="sm" variant="primary" onClick={exportEdit}>
                  Download this edit
                </Button>
              )
            }
          >
            <div className="p-4 text-[13px] text-fg-muted">
              {editor.view === null || editor.view.editedParts.length === 0 ? (
                <p>No part has been rewritten yet.</p>
              ) : (
                <div className="flex flex-wrap items-center gap-1.5">
                  <span>Rewritten:</span>
                  {editor.view.editedParts.map((part) => (
                    <Badge key={part}>{part}</Badge>
                  ))}
                </div>
              )}
              {editor.view?.lastStep == null ? null : (
                <p className="mt-2">Last step: {editor.view.lastStep}</p>
              )}
              {exported === null ? null : (
                <p className="mt-2 text-fg">
                  Exported: {exported.rewritten} part(s) rewritten, {exported.streamed} streamed.{' '}
                  <a
                    href={exported.url}
                    download={name ?? 'deck.pptx'}
                    className="text-accent hover:underline"
                  >
                    Download
                  </a>
                </p>
              )}
              {exportFailure === null ? null : <ErrorState {...exportFailure} />}
              <p className="mt-2 text-[12px] text-fg-faint">
                Every other part is still the bytes that were read; the export streams them back out
                untouched.
              </p>
            </div>
          </Panel>
        </div>
      </div>

      {drawn.error === null ? (
        <SvgSource
          svg={drawn.svg}
          name={`${(name ?? 'deck').replace(/\.ppt[xm]$/, '')}-slide-${String(at + 1)}`}
        />
      ) : null}

      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Three edits with exact inverses, not an editor: no snapping, no resize, no rotate, no
              multi-select and no text editing - those are phases 5 and 6. Charts, tables, SmartArt
              and OLE objects are drawn as frames until their phase draws them.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

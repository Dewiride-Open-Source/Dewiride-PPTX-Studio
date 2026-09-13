'use client';

import { useMemo, useRef, useState } from 'react';

import { flatten, layoutSlide, type Placed } from '@pptx-studio/render-svg';

import { useDeck } from '@/deck/provider';
import { useEditor } from '@/editing/use-editor';
import { mediaFrom } from '@/drawing/media';
import { Stage } from '@/drawing/stage';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel } from '@/shell/panel';

const EMU_PER_POINT = 12700;
const pt = (emu: number): string => (emu / EMU_PER_POINT).toFixed(1);

const SWATCHES = ['#E8453C', '#D98A2B', '#2FA869', '#0B62B0', '#7B1FA2', '#111111', '#FFFFFF'];

const HOW = `import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { slideNode, serializeSvg, flatten } from '@pptx-studio/render-svg';

const doc = loadDocument(PartStore.open(bytes));
const { node, placed } = slideNode(doc.slides[0], doc.slideSize, {
  text: { defaultTextStyle: doc.defaultTextStyle },
  media,
});

const svg = serializeSvg(node);   // the picture
const shapes = flatten(placed);   // the same shapes, for hit-testing`;

export default function SlidesPage() {
  const { deck } = useDeck();
  const editor = useEditor(deck?.bytes ?? null);
  const [at, setAt] = useState(0);
  const [selected, setSelected] = useState<number | null>(null);
  const [overlay, setOverlay] = useState(true);
  const [zoom, setZoom] = useState(820);
  const drag = useRef<{ id: number; x: number; y: number } | null>(null);

  const media = useMemo(() => (editor === null ? undefined : mediaFrom(editor.store)), [editor]);

  if (editor === null) return <p className="text-sm text-ink-400">Reading the deck…</p>;

  const loaded = editor.document;
  const sheet = loaded.slides[Math.min(at, loaded.slides.length - 1)];
  if (sheet === undefined) {
    return <p className="text-sm text-ink-400">This package has no slides.</p>;
  }

  const size = loaded.slideSize;
  const placed: readonly Placed[] = flatten(layoutSlide(sheet));
  const shape = placed.find((one) => one.shape.cNvPrId === selected) ?? null;
  const emuPerPixel = size.cx / zoom;

  return (
    <div className="flex flex-col gap-4">
      <header className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-lg font-semibold text-ink-100">Slides</h1>
          <p className="text-[13px] text-ink-400">
            <Mono>@pptx-studio/render-svg</Mono> draws it; three XML edits change it.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setOverlay((on) => !on)}
            className={`rounded border px-2 py-1 text-[12px] ${
              overlay
                ? 'border-chrome bg-chrome/10 text-chrome'
                : 'border-ink-600 bg-ink-800 text-ink-300'
            }`}
          >
            Debug overlay
          </button>
          <input
            type="range"
            min={360}
            max={1400}
            step={20}
            value={zoom}
            onChange={(event) => setZoom(Number(event.target.value))}
            className="w-32 accent-chrome"
            aria-label="Zoom"
          />
          <button
            type="button"
            disabled={!editor.canUndo}
            onClick={editor.undo}
            className="rounded border border-ink-600 bg-ink-800 px-2 py-1 text-[12px] text-ink-200 disabled:opacity-40"
          >
            Undo
          </button>
          <button
            type="button"
            disabled={!editor.canRedo}
            onClick={editor.redo}
            className="rounded border border-ink-600 bg-ink-800 px-2 py-1 text-[12px] text-ink-200 disabled:opacity-40"
          >
            Redo
          </button>
        </div>
      </header>

      {editor.refused === null ? null : (
        <div className="rounded-lg border border-widened/40 bg-widened/10 p-3 text-[13px] text-widened">
          {editor.refused}
        </div>
      )}

      <div className="flex gap-2 overflow-x-auto pb-1">
        {loaded.slides.map((one, index) => (
          <button
            key={one.partName}
            type="button"
            onClick={() => {
              setAt(index);
              setSelected(null);
            }}
            className={`shrink-0 rounded border px-2 py-1 text-[12px] ${
              index === at
                ? 'border-chrome bg-chrome/10 text-chrome'
                : 'border-ink-700 bg-ink-850 text-ink-300 hover:border-ink-500'
            }`}
          >
            {index + 1}
          </button>
        ))}
      </div>

      <div className="flex flex-wrap items-start gap-4">
        <Stage
          key={`${sheet.partName}-${String(editor.version)}`}
          sheet={sheet}
          size={size}
          defaultTextStyle={loaded.defaultTextStyle}
          media={media}
          width={zoom}
          idPrefix={`s${String(at)}v${String(editor.version)}`}
          selected={selected}
          overlay={overlay}
          onSelect={setSelected}
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

        <div className="flex min-w-72 flex-1 flex-col gap-4">
          <Panel
            title="Selection"
            hint={shape === null ? 'Click a shape on the slide.' : undefined}
          >
            {shape === null ? (
              <p className="p-4 text-[13px] text-ink-400">
                Nothing selected. Clicking picks the innermost shape, so a click inside a group
                selects the child rather than the group.
              </p>
            ) : (
              <dl className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-1.5 p-4 text-[13px]">
                <dt className="text-ink-400">name</dt>
                <dd className="text-ink-100">
                  {shape.shape.name === '' ? `#${String(shape.shape.cNvPrId)}` : shape.shape.name}
                </dd>
                <dt className="text-ink-400">kind</dt>
                <dd className="text-ink-200">{shape.shape.kind}</dd>
                <dt className="text-ink-400">position</dt>
                <dd className="tabular text-ink-200">
                  {pt(shape.frame.x)}, {pt(shape.frame.y)} pt
                </dd>
                <dt className="text-ink-400">size</dt>
                <dd className="tabular text-ink-200">
                  {pt(shape.frame.cx)} x {pt(shape.frame.cy)} pt
                </dd>
                {shape.frame.rot === 0 ? null : (
                  <>
                    <dt className="text-ink-400">rotation</dt>
                    <dd className="tabular text-ink-200">{shape.frame.rot.toFixed(2)} deg</dd>
                  </>
                )}
                {!shape.frame.flipH && !shape.frame.flipV ? null : (
                  <>
                    <dt className="text-ink-400">mirrored</dt>
                    <dd className="text-ink-200">
                      {[shape.frame.flipH ? 'H' : '', shape.frame.flipV ? 'V' : '']
                        .filter(Boolean)
                        .join(' + ')}
                    </dd>
                  </>
                )}
                <dt className="text-ink-400">declared on</dt>
                <dd className="text-ink-200">{shape.sheet.kind}</dd>
              </dl>
            )}
          </Panel>

          {shape === null ? null : (
            <Panel title="Edit" hint="Each one is an XmlEdit that hands back its own inverse.">
              <div className="flex flex-col gap-3 p-4">
                <div>
                  <p className="mb-1.5 text-[12px] text-ink-400">Fill</p>
                  <div className="flex flex-wrap gap-1.5">
                    {SWATCHES.map((hex) => (
                      <button
                        key={hex}
                        type="button"
                        title={hex}
                        aria-label={`Fill ${hex}`}
                        onClick={() => editor.recolour(sheet, shape.shape.cNvPrId, hex)}
                        style={{ background: hex }}
                        className="h-7 w-7 rounded border border-ink-500"
                      />
                    ))}
                  </div>
                </div>
                <label className="block">
                  <span className="mb-1.5 block text-[12px] text-ink-400">First run</span>
                  <input
                    type="text"
                    placeholder="type, then press Enter"
                    onKeyDown={(event) => {
                      if (event.key !== 'Enter') return;
                      editor.retext(sheet, shape.shape.cNvPrId, event.currentTarget.value);
                    }}
                    className="w-full rounded border border-ink-600 bg-ink-800 px-2 py-1 text-[13px] text-ink-100"
                  />
                </label>
                <p className="text-[12px] text-ink-500">
                  Drag the shape on the slide to move it. A placeholder with no{' '}
                  <Mono tone="dim">a:xfrm</Mono> refuses, and says why.
                </p>
              </div>
            </Panel>
          )}

          <Panel title="Session">
            <div className="p-4 text-[13px] text-ink-300">
              {editor.editedParts.length === 0 ? (
                <p>No part has been rewritten yet.</p>
              ) : (
                <div className="flex flex-wrap items-center gap-1.5">
                  <span>Rewritten:</span>
                  {editor.editedParts.map((part) => (
                    <Badge key={part}>{part}</Badge>
                  ))}
                </div>
              )}
              {editor.lastStep === null ? null : (
                <p className="mt-2 text-ink-400">Last step: {editor.lastStep}</p>
              )}
              <p className="mt-2 text-[12px] text-ink-500">
                Every other part is still the bytes that were read. The Round trip tool proves that
                and hands the file back.
              </p>
            </div>
          </Panel>
        </div>
      </div>

      <Panel title="How it is drawn">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

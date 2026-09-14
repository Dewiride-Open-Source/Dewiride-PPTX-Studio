'use client';

import { Download, Redo2, Undo2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react';

import {
  resolveBackgroundColor,
  resolveLatinTypeface,
  resolveRun,
  resolveSize,
  resolveSolidFill,
  resolveOnSheet,
  type ListStyle,
  type Sheet,
} from '@pptx-studio/model';
import { toHexColor, type Rgba } from '@pptx-studio/paint';
import type { Placed } from '@pptx-studio/render-svg';

import { useDocument } from '@/deck/document';
import { DeckWorkerError, useDeckWorker } from '@/deck/worker/client';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Kbd } from '@/design/code';
import { Slider, Switch } from '@/design/field';
import { Panel } from '@/design/panel';
import { ErrorState, Status, type Failure } from '@/design/state';
import type { DemoProps } from '../registry';
import { Stage, useStage, type Hit } from './stage';
import { SvgSource } from './svg-source';
import { drawnText, TextBox, type DrawnText, type Paragraphs } from './text-box';
import { Toolbar } from './toolbar';
import { useEditor } from './use-editor';

const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const HOW = `import { PartStore } from '@pptx-studio/opc';
import { loadDocument, parseSheet } from '@pptx-studio/model';
import { slideNode, serializeSvg, flatten, mediaFromStore, createTextEngine } from '@pptx-studio/render-svg';
import { parseXml, serializeXml, applyEdits, insertInOrder, newElement, newAttribute } from '@pptx-studio/xml';

const store = PartStore.open(bytes);
const doc = loadDocument(store);
const text = createTextEngine({ defaultTextStyle: doc.defaultTextStyle });

// Draw a slide, and get every shape's frame back for hit-testing.
const { node, placed } = slideNode(doc.slides[0], doc.slideSize, { media: mediaFromStore(store), text });
const svg = serializeSvg(node);
const shapes = flatten(placed);

// Edit the slide's own XML. insertInOrder puts a new element where the schema does;
// applyEdits returns the inverses, which is undo.
const tree = parseXml(store.read(doc.slides[0].partName));
const spPr = /* the selected shape's p:spPr */ null;
const inverses = applyEdits([
  insertInOrder(spPr, newElement('a:solidFill', [], [newElement('a:srgbClr', [newAttribute('val', 'E8453C')])])),
]);

// Write the part back and re-read that one sheet; every other part is untouched bytes.
store.replacePart(doc.slides[0].partName, serializeXml(tree));
const sheet = { ...parseSheet(tree.root, doc.slides[0].partName), parent: doc.slides[0].parent, theme: doc.slides[0].theme };`;

const hexOf = (rgba: Rgba): string => `#${toHexColor(rgba)}`;

/** The shape's text as paragraphs of segments, from the model. */
function paragraphsOf(placed: Placed): Paragraphs {
  const text = placed.shape.text;
  if (text === undefined) return [['']];
  return text.paragraphs.map((paragraph) => {
    const segments = [''];
    for (const one of paragraph.content) {
      if (one.kind === 'br') segments.push('');
      else segments[segments.length - 1] += one.text;
    }
    return segments;
  });
}

/** What the box shows when nothing is drawn yet: the resolved face and size, in a colour the background can take. */
function resolvedText(
  placed: Placed,
  sheet: Sheet,
  defaultTextStyle: ListStyle | undefined,
): DrawnText {
  const background = resolveBackgroundColor(sheet);
  const light =
    background === null
      ? true
      : 0.2126 * background.r + 0.7152 * background.g + 0.0722 * background.b > 0.55;
  const paragraph = placed.shape.text?.paragraphs[0];
  if (paragraph === undefined) {
    return {
      fontFamily: 'Calibri',
      fontSize: 18,
      fontWeight: 'normal',
      fontStyle: 'normal',
      fill: light ? '#111111' : '#FFFFFF',
    };
  }
  const context = { sheet, shape: placed.shape, defaultTextStyle };
  const fill = resolveRun(context, paragraph, undefined, (props) => props.fill)?.value;
  const colour =
    fill?.type === 'solid'
      ? hexOf(resolveOnSheet(fill.color, sheet))
      : light
        ? '#111111'
        : '#FFFFFF';
  return {
    fontFamily: resolveLatinTypeface(context, paragraph, undefined).value,
    fontSize: resolveSize(context, paragraph, undefined).value / 100,
    fontWeight:
      resolveRun(context, paragraph, undefined, (props) => props.b)?.value === true
        ? 'bold'
        : 'normal',
    fontStyle: 'normal',
    fill: colour,
  };
}

/** The colour the shape's first run is drawn in, when that is one solid colour. */
function textColourOf(
  placed: Placed,
  sheet: Sheet,
  defaultTextStyle: ListStyle | undefined,
): string | null {
  const paragraph = placed.shape.text?.paragraphs[0];
  if (paragraph === undefined) return null;
  const run = paragraph.content.find((one) => one.kind !== 'br');
  const context = { sheet, shape: placed.shape, defaultTextStyle };
  const fill = resolveRun(context, paragraph, run, (props) => props.fill)?.value;
  return fill?.type === 'solid' ? hexOf(resolveOnSheet(fill.color, sheet)) : null;
}

const FILL_REFUSED: Partial<Record<string, string>> = {
  pic: 'A picture keeps its image',
  grpSp: 'A group has no colour of its own; colour the shapes inside it',
  graphicFrame: 'Charts, tables and diagrams are preserved as they are',
  contentPart: 'Ink is preserved as it is',
};

interface Editing {
  readonly id: number;
  readonly drawn: DrawnText;
  readonly initial: Paragraphs;
}

const KEYS = new WeakMap<Uint8Array, number>();
let next = 0;

/** A key per deck, so a new deck mounts a fresh editor with nothing carried over. */
function keyOf(bytes: Uint8Array | null): number {
  if (bytes === null) return -1;
  let key = KEYS.get(bytes);
  if (key === undefined) {
    key = next;
    next += 1;
    KEYS.set(bytes, key);
  }
  return key;
}

export default function EditorDemo({ full }: DemoProps) {
  const { bytes } = useDocument();
  return <Editor key={keyOf(bytes)} full={full} />;
}

function Editor({ full }: DemoProps) {
  const { parsed, failure, loading, bytes, name } = useDocument();
  const editor = useEditor(bytes, parsed?.document ?? null);
  const worker = useDeckWorker();
  const wrapper = useRef<HTMLDivElement>(null);
  const [at, setAt] = useState(0);
  const [selected, setSelected] = useState<number | null>(null);
  const [editing, setEditing] = useState<Editing | null>(null);
  const [debug, setDebug] = useState(false);
  const [zoom, setZoom] = useState<number | null>(null);
  const [fit, setFit] = useState(820);
  const frame = useRef<HTMLDivElement>(null);
  const [note, setNote] = useState<string | null>(null);
  const [exporting, setExporting] = useState(false);
  const [exported, setExported] = useState<string | null>(null);
  const [exportFailure, setExportFailure] = useState<Failure | null>(null);
  const textOpen = useRef(false);
  const picking = useRef<'fill' | 'text' | null>(null);

  const document = editor.view?.document ?? parsed?.document ?? null;
  const sheet = document?.slides[Math.min(at, (document?.slides.length ?? 1) - 1)];
  const size = document?.slideSize;
  const drawn = useStage(sheet, size, parsed?.text ?? null, `e${String(at)}`, parsed?.media);
  // A drag or a typing session outlives the render it started in, and edits the sheet as it is now.
  const latest = useRef(sheet);
  useEffect(() => {
    latest.current = sheet;
  }, [sheet]);

  // The slide fills the width it has until the zoom slider says otherwise.
  useEffect(() => {
    const element = frame.current;
    if (element === null) return;
    const observer = new ResizeObserver((entries) => {
      const width = entries[0]?.contentRect.width;
      if (width !== undefined && width > 0) setFit(Math.max(320, Math.floor(width)));
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  const chosen = useMemo(
    () => drawn.placed.find((one) => one.shape.cNvPrId === selected && one.sheet === sheet) ?? null,
    [drawn.placed, selected, sheet],
  );

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (parsed === null || document === null || sheet === undefined || size === undefined) {
    return <p className="text-[13px] text-fg-muted">This package has no slides.</p>;
  }

  const width = zoom ?? fit;
  const unit = size.cx / width;
  const stageHeight = Math.round((width * size.cy) / size.cx);
  const shapeName = (placed: Placed): string =>
    placed.shape.name === '' ? `shape #${String(placed.shape.cNvPrId)}` : placed.shape.name;

  // The keys are the slide's; a control that unmounts under the focus hands it back.
  const focusStage = (): void =>
    wrapper.current?.querySelector<HTMLElement>('.stage-surface')?.focus();

  const closeText = (commit: boolean): void => {
    if (!textOpen.current) return;
    textOpen.current = false;
    if (commit) editor.end();
    else editor.cancel();
    setEditing(null);
    focusStage();
  };

  const openText = (placed: Placed): void => {
    if (textOpen.current) return;
    if (placed.shape.text === undefined) {
      setNote(`${shapeName(placed)} holds no text.`);
      return;
    }
    const surface = wrapper.current?.querySelector<HTMLElement>('.stage-surface') ?? null;
    const style =
      (surface === null ? null : drawnText(surface, placed.shape.cNvPrId)) ??
      resolvedText(placed, sheet, document.defaultTextStyle);
    textOpen.current = true;
    editor.begin(`Retyped ${shapeName(placed)}`);
    setNote(null);
    setSelected(placed.shape.cNvPrId);
    setEditing({ id: placed.shape.cNvPrId, drawn: style, initial: paragraphsOf(placed) });
  };

  const select = (hit: Hit | null): void => {
    closeText(true);
    setNote(
      hit !== null && !hit.own
        ? `${shapeName(hit.placed)} comes from the ${hit.placed.sheet.kind}; editing it there is phase 7.`
        : null,
    );
    setSelected(hit?.own === true ? hit.placed.shape.cNvPrId : null);
  };

  const dragStart = (placed: Placed, event: { clientX: number; clientY: number }): void => {
    const id = placed.shape.cNvPrId;
    const { frame } = placed;
    let last = { x: event.clientX, y: event.clientY };
    let began = false;
    const move = (moveEvent: PointerEvent): void => {
      const dx = (moveEvent.clientX - last.x) * unit;
      const dy = (moveEvent.clientY - last.y) * unit;
      // A drag under three pixels is a click, and turning every click into a no-op
      // edit would fill the undo stack with nothing.
      if (!began && Math.abs(dx) + Math.abs(dy) < unit * 3) return;
      if (!began) {
        began = true;
        editor.begin(`Moved ${shapeName(placed)}`);
      }
      last = { x: moveEvent.clientX, y: moveEvent.clientY };
      if (latest.current !== undefined) editor.move(latest.current, id, frame, dx, dy);
    };
    const up = (): void => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (began) editor.end();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const colours = {
    fill: { apply: editor.recolour, label: (name: string) => `Filled ${name}` },
    text: { apply: editor.recolourText, label: (name: string) => `Coloured the text of ${name}` },
  } as const;

  // The native picker fires an input per drag; one gesture takes them all.
  const pick = (which: 'fill' | 'text', hex: string, live: boolean): void => {
    if (chosen === null) return;
    if (live && picking.current === null) {
      picking.current = which;
      editor.begin(colours[which].label(shapeName(chosen)));
    }
    colours[which].apply(sheet, chosen.shape.cNvPrId, hex);
  };

  const pickEnd = (which: 'fill' | 'text'): void => {
    if (picking.current !== which) return;
    picking.current = null;
    editor.end();
  };

  const remove = (id: number): void => {
    closeText(true);
    if (editor.remove(sheet, id)) setSelected(null);
    focusStage();
  };

  const keys = (event: KeyboardEvent<HTMLDivElement>): void => {
    if (editing !== null || !(event.ctrlKey || event.metaKey)) return;
    if (event.key === 'z' && !event.shiftKey) {
      event.preventDefault();
      editor.undo();
    } else if (event.key === 'y' || (event.key === 'z' && event.shiftKey)) {
      event.preventDefault();
      editor.redo();
    }
  };

  const download = (): void => {
    const current = worker.current;
    if (current === null || bytes === null) return;
    closeText(true);
    setExporting(true);
    setExportFailure(null);
    const file = name ?? 'deck.pptx';
    current
      .export(bytes, editor.replaced())
      .then((done) => {
        const url = URL.createObjectURL(new Blob([done.bytes], { type: MIME }));
        const anchor = window.document.createElement('a');
        anchor.href = url;
        anchor.download = file;
        anchor.click();
        URL.revokeObjectURL(url);
        setExported(
          `Downloaded ${file}: ${String(done.rewritten.length)} part${done.rewritten.length === 1 ? '' : 's'} rewritten, ${String(done.streamed)} streamed out untouched.`,
        );
      })
      .catch((cause: unknown) =>
        setExportFailure(
          cause instanceof DeckWorkerError ? cause.failure : { message: String(cause) },
        ),
      )
      .finally(() => setExporting(false));
  };

  const fill = chosen === null ? null : resolveSolidFill(chosen.shape, sheet);
  const textRefused =
    chosen === null || chosen.shape.text === undefined ? 'This shape holds no text' : null;
  const status =
    editor.view?.refused ??
    note ??
    exported ??
    editor.view?.said ??
    'Click a shape to select it. Drag it, double-click to type in it, or pick its fill or text colour from the toolbar.';

  return (
    <div ref={wrapper} className="flex flex-col gap-3" onKeyDown={keys}>
      <div className="flex flex-wrap items-center gap-2">
        <div className="flex gap-1 overflow-x-auto pb-1" role="tablist" aria-label="Slides">
          {document.slides.map((one, index) => (
            <button
              key={one.partName}
              type="button"
              role="tab"
              aria-selected={index === at}
              onClick={() => {
                closeText(true);
                setAt(index);
                setSelected(null);
                setNote(null);
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
        <div className="ml-auto flex flex-wrap items-center gap-2">
          <Button
            size="sm"
            variant="ghost"
            disabled={editor.view?.canUndo !== true}
            onClick={editor.undo}
            title="Undo (Ctrl+Z)"
          >
            <Undo2 size={14} aria-hidden /> Undo
          </Button>
          <Button
            size="sm"
            variant="ghost"
            disabled={editor.view?.canRedo !== true}
            onClick={editor.redo}
            title="Redo (Ctrl+Y)"
          >
            <Redo2 size={14} aria-hidden /> Redo
          </Button>
          <Button size="sm" variant="primary" disabled={exporting} onClick={download}>
            <Download size={14} aria-hidden /> {exporting ? 'Exporting…' : 'Download .pptx'}
          </Button>
        </div>
      </div>

      <div
        ref={frame}
        className="max-w-full overflow-auto rounded-frame border border-line bg-desk p-3"
      >
        <Stage
          content={drawn}
          size={size}
          width={width}
          sheet={sheet}
          selected={selected}
          debug={debug}
          editing={editing?.id ?? null}
          onSelect={select}
          onDragStart={dragStart}
          onOpenText={openText}
          onNudge={(id, dx, dy) => chosen !== null && editor.move(sheet, id, chosen.frame, dx, dy)}
          onDelete={remove}
          inSlide={
            editing === null || chosen === null ? null : (
              <TextBox
                key={`${String(editing.id)}-${String(at)}`}
                placed={chosen}
                sheet={sheet}
                drawn={editing.drawn}
                initial={editing.initial}
                onInput={(paragraphs) => {
                  if (latest.current !== undefined)
                    editor.retext(latest.current, editing.id, paragraphs);
                }}
                onCommit={() => closeText(true)}
                onCancel={() => closeText(false)}
              />
            )
          }
        >
          {chosen === null || editing !== null ? null : (
            <Toolbar
              key={chosen.shape.cNvPrId}
              placed={chosen}
              unit={unit}
              stage={{ width, height: stageHeight }}
              fill={{
                value: fill === null ? null : hexOf(fill),
                refused: FILL_REFUSED[chosen.shape.kind] ?? null,
                onPick: (hex, live) => pick('fill', hex, live),
                onEnd: () => pickEnd('fill'),
              }}
              text={{
                value: textColourOf(chosen, sheet, document.defaultTextStyle),
                refused: textRefused,
                onPick: (hex, live) => pick('text', hex, live),
                onEnd: () => pickEnd('text'),
              }}
              textRefused={textRefused}
              onText={() => openText(chosen)}
              onDelete={() => remove(chosen.shape.cNvPrId)}
            />
          )}
        </Stage>
      </div>

      <div className="flex flex-wrap items-center gap-3 text-[12px]">
        <p
          role="status"
          aria-live="polite"
          className={`min-w-0 basis-full sm:flex-1 sm:basis-auto ${editor.view?.refused == null ? 'text-fg-muted' : 'text-warn'}`}
        >
          {status}
        </p>
        <span className="inline-flex items-center gap-2 text-fg-muted">
          zoom
          <Slider
            ariaLabel="Zoom"
            min={360}
            max={1400}
            step={20}
            value={width}
            onChange={setZoom}
            format={(v) => `${String(Math.round((v / 960) * 100))}%`}
          />
        </span>
        <Switch checked={debug} onChange={setDebug}>
          overlay
        </Switch>
      </div>
      {exportFailure === null ? null : <ErrorState {...exportFailure} />}

      {full ? (
        <>
          {drawn.error === null ? (
            <SvgSource
              svg={drawn.svg}
              name={`${(name ?? 'deck').replace(/.ppt[xm]$/, '')}-slide-${String(at + 1)}`}
            />
          ) : null}
          <Panel title="Keys">
            <dl className="grid grid-cols-[auto_1fr] gap-x-4 gap-y-1.5 p-4 text-[13px] text-fg-muted">
              <dt>
                <Kbd>drag</Kbd>
              </dt>
              <dd>
                move the shape; the arrow keys nudge it a point, with <Kbd>Shift</Kbd> ten
              </dd>
              <dt>
                <Kbd>double-click</Kbd> / <Kbd>Enter</Kbd>
              </dt>
              <dd>
                type in the shape; <Kbd>Enter</Kbd> is a new paragraph, <Kbd>Shift</Kbd>+
                <Kbd>Enter</Kbd> a line break, <Kbd>Esc</Kbd> puts the text back
              </dd>
              <dt>
                <Kbd>toolbar</Kbd>
              </dt>
              <dd>
                Fill and Text each open eight swatches and a picker; a colour is one undo step
              </dd>
              <dt>
                <Kbd>Delete</Kbd>
              </dt>
              <dd>remove the shape</dd>
              <dt>
                <Kbd>Ctrl</Kbd>+<Kbd>Z</Kbd> / <Kbd>Ctrl</Kbd>+<Kbd>Y</Kbd>
              </dt>
              <dd>undo and redo, each step the exact inverse of the XML edit it undoes</dd>
            </dl>
          </Panel>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Move, fill, colour the text, type and delete - not yet resize, rotate, add a shape,
              pick a font, a size or a bullet, which are phases 5 and 6. Typing over a paragraph
              keeps its first run's formatting and drops the others'. A dragged placeholder gets a
              position of its own, as PowerPoint writes it; Reset and Change Layout, which take it
              back, are phase 7. Charts, tables and SmartArt can be moved and deleted but are drawn
              as frames until their phases draw them.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

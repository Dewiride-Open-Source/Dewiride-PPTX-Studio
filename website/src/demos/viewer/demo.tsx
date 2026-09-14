'use client';

import { useCallback, useEffect, useRef, useState } from 'react';

import type { Placed } from '@pptx-studio/render-svg';

import { useDocument } from '@/deck/document';
import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Kbd, Mono } from '@/design/code';
import { Switch } from '@/design/field';
import { Panel } from '@/design/panel';
import { Segmented } from '@/design/segmented';
import { ErrorState, Status } from '@/design/state';
import type { DemoProps } from '../registry';
import { Stage } from './stage';
import { Strip } from './strip';
import { ZOOM_CHOICES, zoomOf, type ZoomChoice } from './zoom';

const HOW = `import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { createTextEngine, mediaFromStore } from '@pptx-studio/render-svg';
import { mountSlide, mountOverlay, mountTextLayer } from '@pptx-studio/render-dom';

const store = PartStore.open(bytes);
const doc = loadDocument(store);
const text = createTextEngine({ defaultTextStyle: doc.defaultTextStyle });   // one measurer, every mount

const mounted = mountSlide(host, doc.slides[0], doc.slideSize, {
  width: 960,                                    // 100%: one CSS pixel per point
  devicePixelRatio: window.devicePixelRatio,    // strokes round to device pixels
  media: mediaFromStore(store),
  text,
});
mountOverlay(mounted.root, mounted.placed[0], { unit: doc.slideSize.cx / 960 });
mountTextLayer(textHost, blocks, { widthPt: 960, heightPt: 540 }, { scale: 1 });

// Zoom: unmount and mount again at the new width. The line breaks do not move.
mounted.unmount();`;

const pt = (emu: number): string => (emu / 12700).toFixed(1);

export default function ViewerDemo({ full }: DemoProps) {
  const { parsed, failure, loading } = useDocument();
  const [at, setAt] = useState(0);
  const [choice, setChoice] = useState<ZoomChoice>('fit');
  const [overlay, setOverlay] = useState(true);
  const [textLayer, setTextLayer] = useState(false);
  const [selected, setSelected] = useState<number | null>(null);
  const [ratio, setRatio] = useState(1);
  const [viewport, setViewport] = useState(960);
  const [mountedIn, setMountedIn] = useState<{ shapes: number; ms: number } | null>(null);
  const [placed, setPlaced] = useState<readonly Placed[]>([]);
  const frame = useRef<HTMLDivElement>(null);

  // The display's ratio and the frame's width, watched rather than read once.
  useEffect(() => {
    const element = frame.current;
    if (element === null) return;
    const measure = () => {
      setViewport(element.clientWidth - 2);
      setRatio(window.devicePixelRatio);
    };
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    let stopped = false;
    const listen = () => {
      window.matchMedia(`(resolution: ${String(window.devicePixelRatio)}dppx)`).addEventListener(
        'change',
        () => {
          if (stopped) return;
          measure();
          listen();
        },
        { once: true },
      );
    };
    listen();
    return () => {
      stopped = true;
      observer.disconnect();
    };
  }, []);

  const onMounted = useCallback((shapes: readonly Placed[], ms: number) => {
    setPlaced(shapes);
    setMountedIn({ shapes: shapes.length, ms });
  }, []);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (parsed === null) return null;

  const { document } = parsed;
  const size = document.slideSize;
  const index = Math.min(at, document.slides.length - 1);
  const sheet = document.slides[index];
  if (sheet === undefined)
    return <p className="text-[13px] text-fg-muted">This package has no slides.</p>;
  const zoom = zoomOf(choice, viewport, size);
  const shape = placed.find((one) => one.shape.cNvPrId === selected) ?? null;

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center gap-3">
        <Segmented
          label="Zoom"
          size="sm"
          options={ZOOM_CHOICES}
          value={choice}
          onChange={setChoice}
        />
        <Switch checked={overlay} onChange={setOverlay}>
          overlay
        </Switch>
        <Switch checked={textLayer} onChange={setTextLayer}>
          selectable text
        </Switch>
        <span className="ml-auto flex items-center gap-2 text-[12px] text-fg-muted">
          <Button
            size="sm"
            variant="ghost"
            onClick={() => setAt(Math.max(0, index - 1))}
            disabled={index === 0}
            aria-label="Previous slide"
          >
            ←
          </Button>
          <span className="tabular font-mono">
            {index + 1} / {document.slides.length}
          </span>
          <Button
            size="sm"
            variant="ghost"
            onClick={() => setAt(Math.min(document.slides.length - 1, index + 1))}
            disabled={index === document.slides.length - 1}
            aria-label="Next slide"
          >
            →
          </Button>
        </span>
      </div>

      <div className="flex flex-col gap-4 lg:flex-row lg:items-start">
        <aside className="lg:w-40 lg:shrink-0" aria-label="Slide strip">
          <div className="lg:hidden">
            <Strip deck={parsed} size={size} current={index} onSelect={setAt} horizontal />
          </div>
          <div className="hidden lg:block">
            <Strip deck={parsed} size={size} current={index} onSelect={setAt} />
          </div>
        </aside>

        <div className="flex min-w-0 flex-1 flex-col gap-4">
          <div ref={frame} className="overflow-auto rounded-frame border border-line bg-desk p-3">
            <Stage
              key={sheet.partName}
              deck={parsed}
              sheet={sheet}
              size={size}
              zoom={zoom}
              devicePixelRatio={ratio}
              overlay={overlay}
              textLayer={textLayer}
              selected={selected}
              onSelect={setSelected}
              onMounted={onMounted}
            />
          </div>
          <div className="flex flex-wrap items-center gap-2 text-[12px] text-fg-muted">
            <Badge>{`${String(Math.round(zoom * 100))}% · ${String(Math.round((zoom * size.cx) / 12700))} px wide`}</Badge>
            <Badge>{`${String(ratio)}× display`}</Badge>
            {mountedIn === null ? null : (
              <Badge tone="good">{`${String(mountedIn.shapes)} shapes mounted in ${mountedIn.ms.toFixed(0)} ms`}</Badge>
            )}
            <span className="text-fg-faint">
              Strokes are rounded at {String(ratio)}×; change the zoom and the SVG changes with it.
              Click a shape, then <Kbd>←</Kbd> <Kbd>→</Kbd> to walk the z-order, <Kbd>Esc</Kbd> to
              clear.
            </span>
          </div>
        </div>

        <Panel
          title="Shape"
          hint={shape === null ? 'Click a shape on the slide.' : undefined}
          className="lg:w-72 lg:shrink-0"
        >
          {shape === null ? (
            <p className="p-4 text-[13px] text-fg-muted">
              Nothing selected. The overlay draws the selected shape&apos;s paths, its text
              rectangle, its connection sites and its adjust handles.
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
              <dt className="text-fg-muted">rotation</dt>
              <dd className="tabular text-fg">{shape.frame.rot.toFixed(1)}°</dd>
              <dt className="text-fg-muted">geometry</dt>
              <dd className="text-fg">
                {shape.geometry?.name ?? 'custom or none'}
                {shape.geometry === null ? '' : ` · ${String(shape.geometry.paths.length)} path(s)`}
              </dd>
              <dt className="text-fg-muted">fill</dt>
              <dd className="text-fg">{shape.fill === null ? 'none' : shape.fill.type}</dd>
              <dt className="text-fg-muted">declared on</dt>
              <dd className="text-fg">{shape.sheet.kind}</dd>
            </dl>
          )}
        </Panel>
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
              No virtualization: every thumbnail of a long deck is mounted, in frame-sized chunks as
              it scrolls into view - slide virtualization is 12.1. No gestures: selection here is a
              click; dragging, colouring and typing are the editor&apos;s. There is no{' '}
              <Mono tone="dim">resize()</Mono> on a mounted slide - a zoom is a fresh mount.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

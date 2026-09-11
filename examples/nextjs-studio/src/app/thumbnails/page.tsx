'use client';

import { useState } from 'react';

import { useDeck } from '@/deck/provider';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `// app/api/thumbnails/route.ts
import { renderDeck } from '@pptx-studio/cli';

export const runtime = 'nodejs';   // it reads font files off the disk

export async function POST(request: Request) {
  const result = renderDeck(new Uint8Array(await request.arrayBuffer()), { width: 640 });
  return Response.json({ slides: result.slides, fonts: result.fonts });
}`;

interface FaceUse {
  readonly asked: string;
  readonly drawn: string;
  readonly substituted: boolean;
}

interface Rendered {
  readonly slides: readonly { number: number; svg: string }[];
  readonly width: number;
  readonly height: number;
  readonly fonts: readonly FaceUse[];
  readonly missing: readonly string[];
  readonly facesIndexed: number;
  readonly fontDirectories: readonly string[];
  readonly ms: number;
}

export default function ThumbnailsPage() {
  const { deck } = useDeck();
  const [result, setResult] = useState<Rendered | null>(null);
  const [failed, setFailed] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const send = (): void => {
    if (deck === null) return;
    setBusy(true);
    setFailed(null);
    setResult(null);
    fetch('/api/thumbnails?width=640', {
      method: 'POST',
      body: new Uint8Array(deck.bytes).slice(),
      headers: { 'content-type': 'application/octet-stream' },
    })
      .then(async (response) => {
        const body: unknown = await response.json();
        if (!response.ok) {
          const message =
            typeof body === 'object' && body !== null && 'error' in body
              ? String((body as { error: unknown }).error)
              : `${String(response.status)}`;
          throw new Error(message);
        }
        setResult(body as Rendered);
      })
      .catch((cause: unknown) => setFailed(cause instanceof Error ? cause.message : String(cause)))
      .finally(() => setBusy(false));
  };

  const substituted = (result?.fonts ?? []).filter((one) => one.substituted);

  return (
    <div className="flex flex-col gap-4">
      <header className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-lg font-semibold text-ink-100">Thumbnails</h1>
          <p className="text-[13px] text-ink-400">
            <Mono>@pptx-studio/cli</Mono> - rendered on the server, with no LibreOffice and no
            headless browser.
          </p>
        </div>
        <button
          type="button"
          onClick={send}
          disabled={busy || deck === null}
          className="rounded border border-widened bg-widened/10 px-3 py-1.5 text-[13px] text-widened disabled:opacity-40"
        >
          {busy ? 'Rendering…' : 'Render on the server'}
        </button>
      </header>

      <div className="rounded-lg border border-widened/30 bg-widened/5 p-3 text-[12px] text-ink-300">
        This is the one page that sends the deck anywhere. It posts the bytes to a route handler in
        this same app, which renders and returns SVG. Every other tool on this site works entirely
        in the tab.
      </div>

      {failed === null ? null : (
        <p className="rounded-lg border border-handle/40 bg-handle/10 p-4 text-[13px] text-handle">
          {failed}
        </p>
      )}

      {result === null ? (
        <Panel title="What the server does that the browser cannot">
          <div className="space-y-2 p-4 text-[13px] leading-relaxed text-ink-300">
            <p>
              Text was the only part of a slide that needed a canvas. On the server there is none,
              so the CLI indexes the machine&apos;s real font files, reads their metrics out of the
              SFNT tables directly, and measures from those.
            </p>
            <p>
              Because the choice of face is ours rather than the browser&apos;s, it is also
              reportable: the response says which typeface was asked for and which was actually
              drawn. A container with no fonts installed substitutes everything, and this is where
              you would see that rather than in a wrong-looking picture.
            </p>
          </div>
        </Panel>
      ) : (
        <>
          <Panel
            aside={
              substituted.length === 0 ? (
                <Badge tone="good">no substitutions</Badge>
              ) : (
                <Badge tone="warn">{substituted.length} substituted</Badge>
              )
            }
          >
            <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-5">
              <Stat label="slides" value={result.slides.length} />
              <Stat label="at" value={`${String(result.width)}x${String(result.height)}`} />
              <Stat label="faces indexed" value={result.facesIndexed} />
              <Stat label="typefaces" value={result.fonts.length} />
              <Stat label="took" value={`${result.ms.toFixed(0)} ms`} />
            </div>
          </Panel>

          <Panel title="Slides">
            <div className="grid gap-4 p-4 sm:grid-cols-2 xl:grid-cols-3">
              {result.slides.map((slide) => (
                <figure key={slide.number}>
                  <div
                    className="stage-surface overflow-hidden [&>svg]:h-auto [&>svg]:w-full"
                    dangerouslySetInnerHTML={{ __html: slide.svg }}
                  />
                  <figcaption className="mt-1.5 text-[11px] text-ink-500">
                    slide {slide.number}
                  </figcaption>
                </figure>
              ))}
            </div>
          </Panel>

          <Panel
            title="Fonts"
            hint="Which typeface the deck asked for, and which one this machine actually drew."
          >
            {result.fonts.length === 0 ? (
              <p className="p-4 text-[13px] text-ink-400">No typeface was asked for.</p>
            ) : (
              <Table head={['asked for', 'drawn in', '']}>
                {result.fonts.map((one, at) => (
                  <Row key={`${one.asked}-${String(at)}`}>
                    <Cell>{one.asked}</Cell>
                    <Cell>{one.drawn}</Cell>
                    <Cell>
                      {one.substituted ? (
                        <Badge tone="warn">substituted</Badge>
                      ) : (
                        <Badge tone="good">as asked</Badge>
                      )}
                    </Cell>
                  </Row>
                ))}
              </Table>
            )}
            {result.missing.length === 0 ? null : (
              <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-widened">
                No indexed face can draw {result.missing.length} code point(s):{' '}
                <Mono>{result.missing.slice(0, 12).join(' ')}</Mono>
              </p>
            )}
            {result.fontDirectories.length === 0 ? null : (
              <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-ink-500">
                Scanned:{' '}
                {result.fontDirectories.map((one) => (
                  <Mono key={one}>{one} </Mono>
                ))}
              </p>
            )}
          </Panel>
        </>
      )}

      <Panel title="How">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

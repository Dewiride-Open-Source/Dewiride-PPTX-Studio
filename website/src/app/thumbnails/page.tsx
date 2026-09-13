'use client';

import { useEffect, useState } from 'react';

import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';
import { publicUrl } from '@/site/base-path';
import { SiteError } from '@/site/errors';

const HOW = `import { renderDeck } from '@pptx-studio/cli';

const result = renderDeck(bytes, { width: 1280 });
result.slides[0]?.svg;                           // stands alone: pictures inline, defs inline
result.fonts.filter((face) => face.substituted); // what this machine could not draw as asked`;

interface FaceUse {
  readonly asked: string;
  readonly drawn: string;
  readonly file: string;
  readonly substituted: boolean;
}

interface RenderedIndex {
  readonly width: number;
  readonly command: string;
  readonly machine: { readonly node: string; readonly platform: string; readonly arch: string };
  readonly decks: readonly { id: string; file: string; slides: number; ms: number }[];
}

interface RenderedDeck {
  readonly id: string;
  readonly file: string;
  readonly width: number;
  readonly height: number;
  readonly slides: readonly { number: number; file: string; bytes: number }[];
  readonly fonts: readonly FaceUse[];
  readonly missing: readonly string[];
  readonly fontDirectories: readonly string[];
  readonly facesIndexed: number;
  readonly ms: number;
  readonly validateExitCode: number;
}

async function fetchJson<T>(path: string): Promise<T> {
  const response = await fetch(publicUrl(path));
  if (!response.ok) {
    throw new SiteError('SITE_RENDERED_MISSING', `${String(response.status)} fetching ${path}`);
  }
  return (await response.json()) as T;
}

export default function ThumbnailsPage() {
  const [index, setIndex] = useState<RenderedIndex | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [deck, setDeck] = useState<RenderedDeck | null>(null);
  const [failed, setFailed] = useState<string | null>(null);

  useEffect(() => {
    fetchJson<RenderedIndex>('/rendered/index.json')
      .then((found) => {
        setIndex(found);
        setSelected(found.decks[0]?.id ?? null);
      })
      .catch((cause: unknown) => setFailed(cause instanceof Error ? cause.message : String(cause)));
  }, []);

  useEffect(() => {
    if (selected === null) return;
    setDeck(null);
    fetchJson<RenderedDeck>(`/rendered/${selected}/render.json`)
      .then(setDeck)
      .catch((cause: unknown) => setFailed(cause instanceof Error ? cause.message : String(cause)));
  }, [selected]);

  const substituted = (deck?.fonts ?? []).filter((one) => one.substituted);

  return (
    <div className="flex flex-col gap-4">
      <header className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-lg font-semibold text-ink-100">Thumbnails</h1>
          <p className="text-[13px] text-ink-400">
            <Mono>@pptx-studio/cli</Mono> - rendered in Node, with no LibreOffice and no headless
            browser.
          </p>
        </div>
        {index === null ? null : (
          <select
            value={selected ?? ''}
            onChange={(event) => setSelected(event.target.value)}
            aria-label="Rendered deck"
            className="rounded border border-ink-600 bg-ink-800 px-2 py-1 font-mono text-[12px] text-ink-100"
          >
            {index.decks.map((one) => (
              <option key={one.id} value={one.id}>
                {one.file} - {one.slides} slide(s)
              </option>
            ))}
          </select>
        )}
      </header>

      <div className="rounded-lg border border-widened/30 bg-widened/5 p-3 text-[12px] text-ink-300">
        The CLI is the one Node package, so nothing here runs in your tab. These slides were
        rendered when this site was built
        {index === null
          ? ''
          : `, on ${index.machine.platform}/${index.machine.arch} under Node ${index.machine.node}`}
        , from the sample decks; a deck you open is not sent anywhere.
      </div>

      {failed === null ? null : (
        <p className="rounded-lg border border-handle/40 bg-handle/10 p-4 text-[13px] text-handle">
          {failed}
        </p>
      )}

      {deck === null ? null : (
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
              <Stat label="slides" value={deck.slides.length} />
              <Stat label="at" value={`${String(deck.width)}x${String(deck.height)}`} />
              <Stat label="faces indexed" value={deck.facesIndexed} />
              <Stat label="typefaces" value={deck.fonts.length} />
              <Stat label="took" value={`${String(deck.ms)} ms`} />
            </div>
          </Panel>

          <Panel title="Slides">
            <div className="grid gap-4 p-4 sm:grid-cols-2 xl:grid-cols-3">
              {deck.slides.map((slide) => (
                <figure key={slide.number}>
                  <img
                    src={publicUrl(`/rendered/${deck.id}/${slide.file}`)}
                    alt={`Slide ${String(slide.number)} of ${deck.file}`}
                    width={deck.width}
                    height={deck.height}
                    loading="lazy"
                    className="stage-surface h-auto w-full"
                  />
                  <figcaption className="mt-1.5 text-[11px] text-ink-500">
                    slide {slide.number} - {(slide.bytes / 1024).toFixed(1)} kB
                  </figcaption>
                </figure>
              ))}
            </div>
          </Panel>

          <Panel
            title="Fonts"
            hint="Which typeface the deck asked for, and which one the build machine actually drew."
          >
            {deck.fonts.length === 0 ? (
              <p className="p-4 text-[13px] text-ink-400">No typeface was asked for.</p>
            ) : (
              <Table head={['asked for', 'drawn in', '']}>
                {deck.fonts.map((one, at) => (
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
            {deck.missing.length === 0 ? null : (
              <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-widened">
                No indexed face can draw {deck.missing.length} code point(s):{' '}
                <Mono>{deck.missing.slice(0, 12).join(' ')}</Mono>
              </p>
            )}
            {deck.fontDirectories.length === 0 ? null : (
              <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-ink-500">
                Scanned:{' '}
                {deck.fontDirectories.map((one) => (
                  <Mono key={one}>{one} </Mono>
                ))}
              </p>
            )}
          </Panel>
        </>
      )}

      <Panel title="How" hint={index?.command}>
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

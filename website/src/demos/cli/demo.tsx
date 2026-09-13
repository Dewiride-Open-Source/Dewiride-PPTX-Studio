'use client';

import { useEffect, useState } from 'react';

import { Badge } from '@/design/badge';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Select } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState } from '@/design/state';
import { Cell, Row, Table } from '@/design/table';
import type { DemoProps } from '../registry';
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

export default function CliDemo({ full }: DemoProps) {
  const [index, setIndex] = useState<RenderedIndex | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [loaded, setLoaded] = useState<RenderedDeck | null>(null);
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
    fetchJson<RenderedDeck>(`/rendered/${selected}/render.json`)
      .then(setLoaded)
      .catch((cause: unknown) => setFailed(cause instanceof Error ? cause.message : String(cause)));
  }, [selected]);

  const deck = loaded !== null && loaded.id === selected ? loaded : null;
  const substituted = (deck?.fonts ?? []).filter((one) => one.substituted);

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <p className="text-[13px] text-fg-muted">
          The CLI is the one Node package, so nothing here runs in your tab. These slides were
          rendered when this site was built
          {index === null
            ? ''
            : `, on ${index.machine.platform}/${index.machine.arch} under Node ${index.machine.node}`}
          , from the sample decks; a deck you open is not sent anywhere.
        </p>
        {index === null ? null : (
          <Select
            value={selected ?? ''}
            onChange={(event) => setSelected(event.target.value)}
            aria-label="Rendered deck"
            className="font-mono text-[12px]"
          >
            {index.decks.map((one) => (
              <option key={one.id} value={one.id}>
                {one.file} · {one.slides} slide(s)
              </option>
            ))}
          </Select>
        )}
      </div>

      {failed === null ? null : <ErrorState message={failed} />}

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
            <StatRow>
              <Stat label="slides" value={deck.slides.length} />
              <Stat label="at" value={`${String(deck.width)}x${String(deck.height)}`} />
              <Stat label="faces indexed" value={deck.facesIndexed} />
              <Stat label="typefaces" value={deck.fonts.length} />
              <Stat label="took" value={`${String(deck.ms)} ms`} />
            </StatRow>
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
                  <figcaption className="mt-1.5 text-[11px] text-fg-faint">
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
              <p className="p-4 text-[13px] text-fg-muted">No typeface was asked for.</p>
            ) : (
              <Table
                head={['asked for', 'drawn in', '']}
                caption="Typefaces the deck asked for, and what drew them"
              >
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
              <p className="border-t border-line px-4 py-3 text-[12px] text-warn">
                No indexed face can draw {deck.missing.length} code point(s):{' '}
                <Mono>{deck.missing.slice(0, 12).join(' ')}</Mono>
              </p>
            )}
            {deck.fontDirectories.length === 0 ? null : (
              <p className="border-t border-line px-4 py-3 text-[12px] text-fg-faint">
                Scanned:{' '}
                {deck.fontDirectories.map((one) => (
                  <Mono key={one}>{one} </Mono>
                ))}
              </p>
            )}
          </Panel>
        </>
      )}

      {full ? (
        <>
          <Panel title="How this page does it" hint={index?.command}>
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Nothing is rendered on your machine: the CLI is Node-only, so the grid above is what
              the build runner drew with the fonts it had. No PNG, no shaping - Arabic and
              Devanagari measure wide - and the face box follows the machine that renders.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

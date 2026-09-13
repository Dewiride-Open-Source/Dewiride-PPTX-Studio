'use client';

import Link from 'next/link';
import { useEffect, useState } from 'react';

import { formatCensus, type PackageCensus } from '@pptx-studio/census';

import { useDeck } from '@/deck/provider';
import { DeckWorkerError, useDeckWorker } from '@/deck/worker/client';
import type { WorkerEnvironment } from '@/deck/worker/protocol';
import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState, Status, type Failure } from '@/design/state';
import { Cell, Row, Table } from '@/design/table';
import type { DemoProps } from '../registry';

const HOW = `// census.worker.ts
import { censusPackage } from '@pptx-studio/census';

self.onmessage = ({ data }: MessageEvent<ArrayBuffer>) => {
  const census = censusPackage(new Uint8Array(data), {
    onProgress: (done, total) => self.postMessage({ progress: { done, total } }),
  });
  self.postMessage({ census }); // plain JSON by design, so it survives postMessage
};

// main.ts
const worker = new Worker(new URL('./census.worker.ts', import.meta.url), { type: 'module' });
worker.postMessage(buffer, [buffer]); // transferred, not copied`;

interface Progress {
  readonly done: number;
  readonly total: number;
}

export default function CensusDemo({ full }: DemoProps) {
  const { deck } = useDeck();
  const worker = useDeckWorker();
  const [environment, setEnvironment] = useState<WorkerEnvironment | null>(null);
  const [progress, setProgress] = useState<Progress | null>(null);
  const [result, setResult] = useState<{ census: PackageCensus; ms: number } | null>(null);
  const [failure, setFailure] = useState<Failure | null>(null);
  const [asText, setAsText] = useState(false);
  const [run, setRun] = useState(0);

  useEffect(() => {
    const current = worker.current;
    if (deck === null || current === null) return;
    let live = true;
    void current.ready.then((found) => {
      if (live) setEnvironment(found);
    });
    current
      .census(deck.bytes, (done, total) => {
        if (live) setProgress({ done, total });
      })
      .then((done) => {
        if (!live) return;
        setResult(done);
        setFailure(null);
        setProgress(null);
      })
      .catch((cause: unknown) => {
        if (!live) return;
        setFailure(cause instanceof DeckWorkerError ? cause.failure : { message: String(cause) });
        setProgress(null);
      });
    return () => {
      live = false;
    };
  }, [deck, worker, run]);

  const census = result?.census ?? null;

  const download = () => {
    if (census === null || deck === null) return;
    const blob = new Blob([JSON.stringify(census, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${deck.name}.census.json`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="flex flex-col gap-4">
      {environment === null ? null : (
        <Panel title="The Worker it ran in">
          <div className="flex flex-wrap items-center gap-2 p-4 text-[13px] text-fg-muted">
            <Badge tone={environment.hasDOMParser ? 'warn' : 'good'}>
              DOMParser: {environment.hasDOMParser ? 'present' : 'absent'}
            </Badge>
            <Badge>OffscreenCanvas: {environment.hasOffscreenCanvas ? 'yes' : 'no'}</Badge>
            <Badge>{environment.hardwareConcurrency} cores</Badge>
            <span className="text-[12px] text-fg-faint">
              {environment.hasDOMParser
                ? 'This browser exposes DOMParser here.'
                : 'No DOMParser in a Worker - which is why the tokenizer is hand-rolled.'}
            </span>
          </div>
        </Panel>
      )}

      {failure !== null ? <ErrorState {...failure} /> : null}
      {progress !== null ? (
        <div className="flex flex-col gap-1">
          <Status busy>
            Scanning {progress.done} / {progress.total} parts…
          </Status>
          <progress
            value={progress.done}
            max={progress.total}
            aria-label="Parts scanned"
            className="h-1.5 w-full max-w-md accent-accent"
          />
        </div>
      ) : null}
      {census === null && failure === null && progress === null ? (
        <Status busy>Handing the bytes to the Worker…</Status>
      ) : null}

      {census === null || result === null ? null : (
        <>
          <Panel
            aside={
              <div className="flex gap-2">
                <Button size="sm" onClick={() => setAsText((on) => !on)} aria-pressed={asText}>
                  {asText ? 'As tables' : 'As the CLI prints it'}
                </Button>
                <Button size="sm" onClick={download}>
                  Download census.json
                </Button>
                <Button size="sm" variant="ghost" onClick={() => setRun((n) => n + 1)}>
                  Run again
                </Button>
              </div>
            }
          >
            <StatRow>
              <Stat label="parts" value={census.parts.length} />
              <Stat label="relationships" value={census.relationships.total} />
              <Stat label="namespaces" value={census.namespaces.length} />
              <Stat label="typefaces" value={census.typefaces.length} />
              <Stat label="scan" value={`${result.ms.toFixed(0)} ms`} />
            </StatRow>
          </Panel>

          {asText ? (
            <Code lang="txt" code={formatCensus(census)} title="pptx-studio inspect" />
          ) : (
            <div className="grid gap-4 lg:grid-cols-2">
              <Panel title="Features" hint="Each keyed to the phase of the plan that renders it.">
                {census.features.length === 0 ? (
                  <p className="p-4 text-[13px] text-fg-muted">No tracked feature in this deck.</p>
                ) : (
                  <Table head={['feature', 'phase', 'count']} caption="Features found in the deck">
                    {census.features.map((feature) => (
                      <Row key={feature.key}>
                        <Cell>{feature.label}</Cell>
                        <Cell>
                          <Link href={`/docs/status#phase-${feature.phase.split('.')[0] ?? ''}`}>
                            <Badge tone="info">{feature.phase}</Badge>
                          </Link>
                        </Cell>
                        <Cell className="tabular">{feature.count}</Cell>
                      </Row>
                    ))}
                  </Table>
                )}
              </Panel>

              <Panel title="Namespaces" hint="With the prefixes they were actually spelled with.">
                <div className="max-h-96 overflow-y-auto">
                  <Table head={['uri', 'as', 'count']} caption="Namespaces in the deck">
                    {census.namespaces.map((ns) => (
                      <Row key={ns.uri}>
                        <Cell>
                          <Mono tone="dim">{ns.uri.replace(/^https?:\/\//, '')}</Mono>{' '}
                          {ns.extension ? <Badge tone="warn">extension</Badge> : null}
                        </Cell>
                        <Cell>
                          <Mono>{ns.prefixes.join(' ')}</Mono>
                        </Cell>
                        <Cell className="tabular">{ns.count}</Cell>
                      </Row>
                    ))}
                  </Table>
                </div>
              </Panel>

              <Panel title="Preset geometry" hint="Which of the 187 presets this deck uses.">
                {census.presetGeometry.length === 0 ? (
                  <p className="p-4 text-[13px] text-fg-muted">No preset geometry.</p>
                ) : (
                  <div className="flex flex-wrap gap-1.5 p-4">
                    {census.presetGeometry.map((preset) => (
                      <Badge key={preset.name}>
                        {preset.name} ×{preset.count}
                      </Badge>
                    ))}
                  </div>
                )}
              </Panel>

              <Panel title="Typefaces">
                {census.typefaces.length === 0 ? (
                  <p className="p-4 text-[13px] text-fg-muted">No typeface named.</p>
                ) : (
                  <div className="flex flex-wrap gap-1.5 p-4">
                    {census.typefaces.map((face) => (
                      <Badge key={face.name}>
                        {face.name} ×{face.count}
                      </Badge>
                    ))}
                  </div>
                )}
              </Panel>
            </div>
          )}

          {census.problems.length === 0 ? null : (
            <Panel title="Problems" hint="A census does not refuse; it records and carries on.">
              <Table
                head={['severity', 'code', 'part', 'message']}
                caption="Problems the census recorded"
              >
                {census.problems.map((problem, at) => (
                  <Row key={at}>
                    <Cell>
                      <Badge tone={problem.severity === 'error' ? 'bad' : 'warn'}>
                        {problem.severity}
                      </Badge>
                    </Cell>
                    <Cell>
                      <Mono>{problem.code}</Mono>
                    </Cell>
                    <Cell>
                      <Mono tone="dim">{problem.part ?? '-'}</Mono>
                    </Cell>
                    <Cell>{problem.message}</Cell>
                  </Row>
                ))}
              </Table>
            </Panel>
          )}
        </>
      )}

      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              A census is not the validator: its problems are observations about what is there, not
              rules a file broke. It refuses nothing, so a damaged deck still gets a count.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

'use client';

import { useEffect, useState } from 'react';

import type { PackageCensus } from '@pptx-studio/census';

import type { CensusResponse, WorkerEnvironment } from '@/census/worker';
import { useDeck } from '@/deck/provider';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `// census.worker.ts
import { censusPackage } from '@pptx-studio/census';

self.addEventListener('message', (event) => {
  const census = censusPackage(new Uint8Array(event.data.bytes), {
    onProgress: (done, total) => postMessage({ done, total }),
  });
  postMessage({ census });   // plain JSON, so it survives postMessage
});`;

export default function CensusPage() {
  const { deck } = useDeck();
  const [census, setCensus] = useState<PackageCensus | null>(null);
  const [environment, setEnvironment] = useState<WorkerEnvironment | null>(null);
  const [ms, setMs] = useState<number | null>(null);
  const [progress, setProgress] = useState<{ done: number; total: number } | null>(null);
  const [failed, setFailed] = useState<string | null>(null);

  useEffect(() => {
    if (deck === null) return;
    setCensus(null);
    setFailed(null);
    setProgress(null);

    const worker = new Worker(new URL('../../census/worker.ts', import.meta.url), {
      type: 'module',
    });
    worker.addEventListener('message', (event: MessageEvent<CensusResponse>) => {
      const message = event.data;
      if (message.type === 'ready') setEnvironment(message.environment);
      else if (message.type === 'progress')
        setProgress({ done: message.done, total: message.total });
      else if (message.type === 'done') {
        setCensus(message.census);
        setEnvironment(message.environment);
        setMs(message.ms);
        setProgress(null);
      } else setFailed(message.message);
    });

    // The buffer is copied rather than transferred: the deck is shared with
    // every other tool on this site and a transfer would detach it from them.
    worker.postMessage({ bytes: deck.bytes.slice().buffer });
    return () => worker.terminate();
  }, [deck]);

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Census</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/census</Mono> - what is inside, counted in a Web Worker.
        </p>
      </header>

      {environment === null ? null : (
        <Panel title="The Worker it ran in">
          <div className="flex flex-wrap items-center gap-2 p-4 text-[13px] text-ink-300">
            <Badge tone={environment.hasDOMParser ? 'warn' : 'good'}>
              DOMParser: {environment.hasDOMParser ? 'present' : 'absent'}
            </Badge>
            <Badge tone="plain">
              OffscreenCanvas: {environment.hasOffscreenCanvas ? 'yes' : 'no'}
            </Badge>
            <Badge tone="plain">{environment.hardwareConcurrency} cores</Badge>
            <span className="text-[12px] text-ink-500">
              {environment.hasDOMParser
                ? 'This browser exposes DOMParser here.'
                : 'No DOMParser in a Worker - which is why the tokenizer is hand-rolled.'}
            </span>
          </div>
        </Panel>
      )}

      {failed !== null ? <p className="text-sm text-handle">{failed}</p> : null}
      {progress !== null ? (
        <p className="text-sm text-ink-400">
          Scanning {progress.done} / {progress.total} parts…
        </p>
      ) : null}
      {census === null && failed === null && progress === null ? (
        <p className="text-sm text-ink-400">Handing the bytes to the Worker…</p>
      ) : null}

      {census === null ? null : (
        <>
          <Panel>
            <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-5">
              <Stat label="parts" value={census.parts.length} />
              <Stat label="relationships" value={census.relationships.total} />
              <Stat label="namespaces" value={census.namespaces.length} />
              <Stat label="typefaces" value={census.typefaces.length} />
              <Stat label="scan" value={ms === null ? '-' : `${ms.toFixed(0)} ms`} />
            </div>
          </Panel>

          <div className="grid gap-4 lg:grid-cols-2">
            <Panel title="Features" hint="Each keyed to the phase of the plan that renders it.">
              {census.features.length === 0 ? (
                <p className="p-4 text-[13px] text-ink-400">No tracked feature in this deck.</p>
              ) : (
                <Table head={['feature', 'phase', 'count']}>
                  {census.features.map((feature) => (
                    <Row key={feature.key}>
                      <Cell>{feature.label}</Cell>
                      <Cell>
                        <Badge tone="info">{feature.phase}</Badge>
                      </Cell>
                      <Cell className="tabular">{feature.count}</Cell>
                    </Row>
                  ))}
                </Table>
              )}
            </Panel>

            <Panel title="Namespaces" hint="With the prefixes they were actually spelled with.">
              <div className="max-h-96 overflow-y-auto">
                <Table head={['uri', 'as', 'count']}>
                  {census.namespaces.map((ns) => (
                    <Row key={ns.uri}>
                      <Cell>
                        <Mono tone="dim">{ns.uri.replace(/^https?:\/\//, '')}</Mono>
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
                <p className="p-4 text-[13px] text-ink-400">No preset geometry.</p>
              ) : (
                <div className="flex flex-wrap gap-1.5 p-4">
                  {census.presetGeometry.map((preset) => (
                    <Badge key={preset.name}>
                      {preset.name} x{preset.count}
                    </Badge>
                  ))}
                </div>
              )}
            </Panel>

            <Panel title="Typefaces">
              {census.typefaces.length === 0 ? (
                <p className="p-4 text-[13px] text-ink-400">No typeface named.</p>
              ) : (
                <div className="flex flex-wrap gap-1.5 p-4">
                  {census.typefaces.map((face) => (
                    <Badge key={face.name}>
                      {face.name} x{face.count}
                    </Badge>
                  ))}
                </div>
              )}
            </Panel>
          </div>

          {census.problems.length === 0 ? null : (
            <Panel title="Problems" hint="A census does not refuse; it records and carries on.">
              <Table head={['severity', 'code', 'part', 'message']}>
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

      <Panel title="How">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

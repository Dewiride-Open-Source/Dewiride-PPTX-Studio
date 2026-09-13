'use client';

import { useState } from 'react';

import { PartStore } from '@pptx-studio/opc';
import { comparePackages, exportPackage, openPackage } from '@pptx-studio/writer';

import { useDeck } from '@/deck/provider';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `import { openPackage, exportPackage, comparePackages } from '@pptx-studio/writer';

// Two stores over the same bytes: the one you edit, and the untouched twin
// every check compares against. Without the baseline, three rules, the media
// sweep and the preservation check all quietly skip themselves.
const { store, baseline, baselineBytes } = openPackage(bytes);

const result = exportPackage({ store, baseline, baselineBytes });
result.rewritten;   // parts serialised afresh - empty on a no-op
result.streamed;    // parts copied across still compressed`;

interface Outcome {
  readonly rewritten: readonly string[];
  readonly streamed: number;
  readonly preservation: { checked: number; rewritten: number; skipped: string | null };
  readonly counts: { same: number; xml: number; binary: number; relationships: number };
  readonly differences: readonly { kind: string; part: string; detail: string }[];
  readonly ok: boolean;
  readonly bytesIn: number;
  readonly bytesOut: number;
  readonly ms: number;
  readonly blob: string;
}

export default function RoundTripPage() {
  const { deck } = useDeck();
  const [outcome, setOutcome] = useState<Outcome | null>(null);
  const [failed, setFailed] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const run = (): void => {
    if (deck === null) return;
    setBusy(true);
    setFailed(null);
    setOutcome(null);
    // A frame, so the button paints its busy state before the main thread is
    // taken for the length of an inflate, a parse and two archive reads.
    requestAnimationFrame(() => {
      try {
        const startedAt = performance.now();
        const { store, baseline, baselineBytes } = openPackage(deck.bytes);
        const result = exportPackage({ store, baseline, baselineBytes });
        const comparison = comparePackages(
          PartStore.open(baselineBytes),
          PartStore.open(result.bytes),
        );
        setOutcome({
          rewritten: [...result.rewritten],
          streamed: result.streamed,
          preservation: {
            checked: result.preservation.checked,
            rewritten: result.preservation.rewritten,
            skipped: result.preservation.skipped,
          },
          counts: comparison.counts,
          differences: comparison.differences.map((one) => ({
            kind: one.kind,
            part: one.part,
            detail: one.detail,
          })),
          ok: comparison.ok,
          bytesIn: deck.bytes.byteLength,
          bytesOut: result.bytes.byteLength,
          ms: performance.now() - startedAt,
          blob: URL.createObjectURL(
            new Blob([result.bytes.slice()], {
              type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            }),
          ),
        });
      } catch (cause) {
        setFailed(cause instanceof Error ? cause.message : String(cause));
      } finally {
        setBusy(false);
      }
    });
  };

  return (
    <div className="flex flex-col gap-4">
      <header className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-lg font-semibold text-ink-100">Round trip</h1>
          <p className="text-[13px] text-ink-400">
            <Mono>@pptx-studio/writer</Mono> - hand the file back, and prove nothing moved.
          </p>
        </div>
        <button
          type="button"
          onClick={run}
          disabled={busy || deck === null}
          className="rounded border border-chrome bg-chrome/10 px-3 py-1.5 text-[13px] text-chrome disabled:opacity-40"
        >
          {busy ? 'Exporting…' : 'Export this deck'}
        </button>
      </header>

      {failed === null ? null : (
        <div className="rounded-lg border border-handle/40 bg-handle/10 p-4 text-[13px] text-handle">
          <p className="font-medium">The firewall refused to hand over bytes.</p>
          <p className="mt-1">{failed}</p>
        </div>
      )}

      {outcome === null ? (
        <Panel title="What happens when you press it">
          <div className="space-y-2 p-4 text-[13px] leading-relaxed text-ink-300">
            <p>
              Four things, in order: the package is opened twice - once to edit and once as the
              untouched baseline; the prepare hooks and the media sweep run; the archive is written;
              and then the preservation check and the twenty-nine rules run over what was written.
              Bytes are returned only if both pass.
            </p>
            <p>
              On a deck nothing has edited, <Mono>rewritten</Mono> should be empty and every part
              should be streamed out of the source archive still compressed. That is not an
              optimisation - it is why charts, SmartArt, animations and OLE objects survive a save
              from a tool that cannot draw them.
            </p>
          </div>
        </Panel>
      ) : (
        <>
          <Panel
            aside={
              outcome.ok ? (
                <Badge tone="good">identical</Badge>
              ) : (
                <Badge tone="warn">{outcome.differences.length} difference(s)</Badge>
              )
            }
          >
            <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-5">
              <Stat
                label="rewritten"
                value={outcome.rewritten.length}
                tone={outcome.rewritten.length === 0 ? 'good' : 'warn'}
              />
              <Stat label="streamed" value={outcome.streamed} tone="good" />
              <Stat label="parts same" value={outcome.counts.same} />
              <Stat label="bytes out" value={outcome.bytesOut.toLocaleString()} />
              <Stat label="took" value={`${outcome.ms.toFixed(0)} ms`} />
            </div>
            <div className="flex flex-wrap items-center gap-3 border-t border-ink-700 px-4 py-3">
              <a
                href={outcome.blob}
                download={deck?.name ?? 'export.pptx'}
                className="rounded border border-good/50 bg-good/10 px-3 py-1.5 text-[13px] text-good"
              >
                Download the .pptx
              </a>
              <span className="text-[12px] text-ink-500">
                Open it in PowerPoint. There should be no repair prompt and no visible difference.
              </span>
            </div>
          </Panel>

          <Panel title="Preservation check" hint="The bytes, not the store's opinion of them.">
            <div className="grid grid-cols-3 gap-4 p-4">
              <Stat label="entries checked" value={outcome.preservation.checked} />
              <Stat label="skipped as ours" value={outcome.preservation.rewritten} />
              <Stat
                label="ran"
                value={outcome.preservation.skipped === null ? 'yes' : 'no'}
                tone={outcome.preservation.skipped === null ? 'good' : 'warn'}
              />
            </div>
            {outcome.preservation.skipped === null ? null : (
              <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-widened">
                {outcome.preservation.skipped}
              </p>
            )}
          </Panel>

          <Panel
            title="Comparison"
            hint="Canonical XML per part, relationship-graph isomorphism, SHA-256 on media. Never raw ZIP bytes."
          >
            <div className="grid grid-cols-4 gap-4 p-4">
              <Stat label="compared as xml" value={outcome.counts.xml} />
              <Stat label="as relationships" value={outcome.counts.relationships} />
              <Stat label="as binary" value={outcome.counts.binary} />
              <Stat
                label="differences"
                value={outcome.differences.length}
                tone={outcome.differences.length === 0 ? 'good' : 'warn'}
              />
            </div>
            {outcome.differences.length === 0 ? null : (
              <Table head={['kind', 'part', 'detail']}>
                {outcome.differences.map((one, at) => (
                  <Row key={at}>
                    <Cell>
                      <Badge tone="warn">{one.kind}</Badge>
                    </Cell>
                    <Cell>
                      <Mono tone="dim">{one.part}</Mono>
                    </Cell>
                    <Cell>{one.detail}</Cell>
                  </Row>
                ))}
              </Table>
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

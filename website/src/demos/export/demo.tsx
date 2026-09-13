'use client';

import { useEffect, useMemo, useState } from 'react';

import { useDocument } from '@/deck/document';
import { DeckWorkerError, useDeckWorker } from '@/deck/worker/client';
import type { ExportDone } from '@/deck/worker/protocol';
import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState, Status, type Failure } from '@/design/state';
import { Cell, Row, Table } from '@/design/table';
import type { DemoProps } from '../registry';
import { stampLastModifiedBy } from './stamp';

const HOW = `import { openPackage, exportPackage, comparePackages } from '@pptx-studio/writer';

// Two stores over the same bytes: the one you edit, and the untouched twin
// every check compares against. Without the baseline, three rules, the media
// sweep and the preservation check all quietly skip themselves.
const { store, baseline, baselineBytes } = openPackage(bytes);
store.replacePart('/docProps/core.xml', edited);     // or nothing at all

const result = exportPackage({ store, baseline, baselineBytes });
result.rewritten;   // parts serialised afresh - empty on a no-op
result.streamed;    // parts copied across still compressed
result.bytes;       // only if the preservation check and the 29 rules passed`;

const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** What a streamed part is, by content type, for the table of what survived untouched. */
const KINDS: readonly { label: string; match: RegExp }[] = [
  { label: 'chart', match: /drawingml\.chart|chartex|chartstyle|chartcolorstyle/ },
  { label: 'SmartArt', match: /drawingml\.diagram/ },
  { label: 'OLE / embedded', match: /oleObject|vnd\.ms-excel|vnd\.ms-word|embeddedPackage|\.bin$/ },
  { label: 'macro', match: /vbaProject|vnd\.ms-office\.vba/ },
  { label: 'media', match: /^image\/|^video\/|^audio\/|\.emf$|\.wmf$/ },
  { label: 'font', match: /fontdata|x-fontdata/ },
  { label: 'notes and comments', match: /notesSlide|comments|commentAuthors/ },
];

interface Outcome extends ExportDone {
  /** The deck it was made from; an outcome for other bytes is stale. */
  readonly forBytes: Uint8Array;
  readonly url: string;
  readonly edit: 'none' | 'stamp';
}

export default function ExportDemo({ full }: DemoProps) {
  const { parsed, failure: openFailure, loading, bytes, name } = useDocument();
  const worker = useDeckWorker();
  const [made, setMade] = useState<Outcome | null>(null);
  const [failure, setFailure] = useState<Failure | null>(null);
  const [busy, setBusy] = useState<'none' | 'stamp' | null>(null);
  const outcome = made !== null && made.forBytes === bytes ? made : null;

  useEffect(
    () => () => {
      if (made !== null) URL.revokeObjectURL(made.url);
    },
    [made],
  );

  const run = (edit: 'none' | 'stamp') => {
    const current = worker.current;
    if (parsed === null || bytes === null || current === null) return;
    setBusy(edit);
    setFailure(null);
    const replaced = edit === 'stamp' ? stampLastModifiedBy(parsed.store, 'PPTX Studio') : null;
    current
      .export(bytes, replaced === null ? [] : [replaced])
      .then((done) => {
        setMade({
          ...done,
          forBytes: bytes,
          url: URL.createObjectURL(new Blob([done.bytes], { type: MIME })),
          edit,
        });
      })
      .catch((cause: unknown) => {
        setFailure(cause instanceof DeckWorkerError ? cause.failure : { message: String(cause) });
      })
      .finally(() => setBusy(null));
  };

  const streamedByKind = useMemo(() => {
    if (parsed === null || outcome === null) return [];
    const rewritten = new Set(outcome.rewritten);
    const counts = new Map<string, number>();
    for (const part of parsed.store.partNames) {
      if (rewritten.has(part)) continue;
      const type = `${parsed.store.contentTypeOf(part) ?? ''} ${part}`;
      const kind = KINDS.find((one) => one.match.test(type));
      if (kind !== undefined) counts.set(kind.label, (counts.get(kind.label) ?? 0) + 1);
    }
    return [...counts.entries()];
  }, [parsed, outcome]);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (openFailure !== null) return <ErrorState {...openFailure} />;
  if (parsed === null || bytes === null) return null;

  const canStamp = parsed.store.has('/docProps/core.xml');

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap items-center gap-2">
        <Button variant="primary" onClick={() => run('none')} disabled={busy !== null}>
          {busy === 'none' ? 'Exporting…' : 'Export unchanged'}
        </Button>
        <Button onClick={() => run('stamp')} disabled={busy !== null || !canStamp}>
          {busy === 'stamp' ? 'Exporting…' : 'Export with one edit'}
        </Button>
        <span className="text-[12px] text-fg-muted">
          {canStamp
            ? 'The edit sets cp:lastModifiedBy in /docProps/core.xml - the smallest edit a package can carry.'
            : 'This deck has no core properties part, so there is nothing small to edit.'}
        </span>
      </div>

      {failure === null ? null : (
        <ErrorState
          {...failure}
          advice={
            failure.advice ?? 'The firewall refused to hand over bytes; the message names the rule.'
          }
        />
      )}

      {outcome === null ? (
        <Panel title="What happens when you press it">
          <div className="space-y-2 p-4 text-[13px] leading-relaxed text-fg-muted">
            <p>
              Four things, in order, in a Worker: the package is opened twice - once to edit and
              once as the untouched baseline; the prepare hooks and the media sweep run; the archive
              is written; then the preservation check and the twenty-nine rules run over what was
              written. Bytes come back only if both pass.
            </p>
            <p>
              On a deck nothing has edited, <Mono>rewritten</Mono> is empty and every part is
              streamed out of the source archive still compressed. That is not an optimisation - it
              is why charts, SmartArt, animations and OLE objects survive a save from a tool that
              cannot draw them.
            </p>
          </div>
        </Panel>
      ) : (
        <>
          <Panel
            title={outcome.edit === 'none' ? 'Exported unchanged' : 'Exported with one edit'}
            aside={
              <div className="flex items-center gap-2">
                {outcome.comparison.ok ? (
                  <Badge tone="good">same deck</Badge>
                ) : (
                  <Badge tone="warn">{outcome.comparison.differences.length} difference(s)</Badge>
                )}
                <a
                  href={outcome.url}
                  download={name ?? 'deck.pptx'}
                  className="text-[13px] text-accent hover:underline"
                >
                  Download
                </a>
              </div>
            }
          >
            <StatRow>
              <Stat
                label="rewritten"
                value={outcome.rewritten.length}
                tone={outcome.rewritten.length === 0 ? 'good' : 'warn'}
              />
              <Stat label="streamed" value={outcome.streamed} tone="good" />
              <Stat label="parts same" value={outcome.comparison.counts.same} />
              <Stat label="bytes out" value={outcome.bytes.byteLength.toLocaleString()} />
              <Stat label="took" value={`${outcome.ms.toFixed(0)} ms`} />
            </StatRow>
            {outcome.rewritten.length === 0 ? null : (
              <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
                Serialised afresh:{' '}
                {outcome.rewritten.map((part) => (
                  <Mono key={part}>{part} </Mono>
                ))}
              </p>
            )}
            <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
              Open the download in PowerPoint. A repair prompt is a bug - please report it with the
              census.
            </p>
          </Panel>

          <div className="grid gap-4 lg:grid-cols-2">
            <Panel title="Preservation check" hint="The bytes, not the store's opinion of them.">
              <StatRow>
                <Stat label="entries checked" value={outcome.preservation.checked} />
                <Stat label="skipped as ours" value={outcome.preservation.rewritten} />
              </StatRow>
              <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
                {outcome.preservation.skipped === null
                  ? 'Every entry not rewritten this session is byte-identical to the archive it came from.'
                  : outcome.preservation.skipped}
              </p>
            </Panel>

            <Panel
              title="What survived untouched"
              hint="Streamed parts by kind - the content this tool cannot draw yet."
            >
              {streamedByKind.length === 0 ? (
                <p className="p-4 text-[13px] text-fg-muted">
                  Nothing of these kinds in this deck.
                </p>
              ) : (
                <Table head={['kind', 'parts streamed']} caption="Streamed parts by kind">
                  {streamedByKind.map(([kind, count]) => (
                    <Row key={kind}>
                      <Cell>{kind}</Cell>
                      <Cell className="tabular">{count}</Cell>
                    </Row>
                  ))}
                </Table>
              )}
            </Panel>
          </div>

          <Panel
            title="Round trip comparison"
            hint="Not a byte comparison: XML canonically, relationships as graphs with opaque ids, binaries by digest."
          >
            <StatRow>
              <Stat label="compared as xml" value={outcome.comparison.counts.xml} />
              <Stat label="as relationships" value={outcome.comparison.counts.relationships} />
              <Stat label="as binary" value={outcome.comparison.counts.binary} />
            </StatRow>
            {outcome.comparison.differences.length === 0 ? null : (
              <Table head={['kind', 'part', 'detail']} caption="Differences after the round trip">
                {outcome.comparison.differences.map((one, at) => (
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

          {outcome.report === null ? null : (
            <Panel
              title="The firewall, with a baseline"
              hint="V027 to V029 compare against the package as opened; they only run on an export."
            >
              <StatRow>
                <Stat label="ran" value={outcome.report.checked.length} />
                <Stat
                  label="findings"
                  value={outcome.report.findings.length}
                  tone={outcome.report.findings.length > 0 ? 'warn' : 'good'}
                />
                <Stat
                  label="blocking"
                  value={outcome.report.blocking}
                  tone={outcome.report.blocking > 0 ? 'bad' : 'good'}
                />
              </StatRow>
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
              Nothing here proves PowerPoint opens the file: no hosted runner has Office, so that is
              measured by hand. The media sweep collects only what an edit orphaned, and no edit
              here orphans anything.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

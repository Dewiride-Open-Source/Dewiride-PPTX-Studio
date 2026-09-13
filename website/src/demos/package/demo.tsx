'use client';

import { useMemo, useState } from 'react';

import { DEFAULT_ZIP_LIMITS, isRelationshipPartName, sha256Hex } from '@pptx-studio/opc';

import { useDocument } from '@/deck/document';
import { Badge } from '@/design/badge';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { TextInput } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState, Status } from '@/design/state';
import { Cell, Row, Table } from '@/design/table';
import type { DemoProps } from '../registry';

const HOW = `import { PartStore } from '@pptx-studio/opc';

const store = PartStore.open(bytes);          // hostile-input limits on by default

store.partNames;                              // every part in the package
store.contentTypeOf('/ppt/slides/slide1.xml');
store.info(name)?.size;                       // uncompressed bytes

// Relationships are scoped to one part's own .rels, never global.
const rels = store.relationships('/ppt/slides/slide1.xml');
rels.all;                                     // { id, type, target, targetMode }
rels.resolve(rels.all[0]);                    // the part it points at
store.danglingRelationships();                // edges that point at nothing`;

/** The tail of a relationship type URI, which is the part worth reading. */
function shortType(type: string): string {
  const cut = type.lastIndexOf('/');
  return cut < 0 ? type : type.slice(cut + 1);
}

function bytes(size: number): string {
  if (size < 1024) return `${String(size)} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} kB`;
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}

function hexPreview(data: Uint8Array): string {
  const head = data.subarray(0, 64);
  const lines: string[] = [];
  for (let at = 0; at < head.length; at += 16) {
    const chunk = head.subarray(at, at + 16);
    const hex = [...chunk].map((byte) => byte.toString(16).padStart(2, '0')).join(' ');
    const ascii = [...chunk]
      .map((byte) => (byte >= 0x20 && byte < 0x7f ? String.fromCharCode(byte) : '.'))
      .join('');
    lines.push(`${at.toString(16).padStart(4, '0')}  ${hex.padEnd(47)}  ${ascii}`);
  }
  return lines.join('\n');
}

const LIMITS: readonly { label: string; value: string }[] = [
  { label: 'archive', value: bytes(DEFAULT_ZIP_LIMITS.maxArchiveBytes) },
  { label: 'inflated, total', value: bytes(DEFAULT_ZIP_LIMITS.maxTotalInflatedBytes) },
  { label: 'inflated, one entry', value: bytes(DEFAULT_ZIP_LIMITS.maxEntryInflatedBytes) },
  { label: 'entries', value: String(DEFAULT_ZIP_LIMITS.maxEntries) },
  { label: 'compression ratio', value: `${String(DEFAULT_ZIP_LIMITS.maxCompressionRatio)}:1` },
];

export default function PackageDemo({ full }: DemoProps) {
  const { parsed, failure, loading } = useDocument();
  const [selected, setSelected] = useState<string | null>(null);
  const [filter, setFilter] = useState('');

  const parts = useMemo(() => {
    if (parsed === null) return [];
    return parsed.store.partNames
      .map((name) => ({
        name,
        contentType: parsed.store.contentTypeOf(name) ?? '(none)',
        origin: parsed.store.contentTypes.resolve(name)?.origin ?? null,
        size: parsed.store.info(name)?.size ?? 0,
        rels: isRelationshipPartName(name),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [parsed]);

  const dangling = useMemo(
    () => (parsed === null ? [] : parsed.store.danglingRelationships()),
    [parsed],
  );

  const detail = useMemo(() => {
    if (parsed === null || selected === null) return null;
    const contentType = parsed.store.contentTypeOf(selected) ?? '';
    const isXml = contentType.endsWith('+xml') || contentType.endsWith('/xml');
    const content = parsed.store.read(selected);
    return {
      relationships: isRelationshipPartName(selected)
        ? []
        : parsed.store.relationships(selected).all,
      digest: isXml ? null : sha256Hex(content),
      hex: hexPreview(content),
      text: isXml ? new TextDecoder().decode(content.subarray(0, 600)) : null,
    };
  }, [parsed, selected]);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (parsed === null) return null;

  const total = parts.reduce((sum, part) => sum + part.size, 0);
  const xml = parts.filter((part) => part.contentType.endsWith('xml')).length;
  const shown = parts.filter((part) => part.name.toLowerCase().includes(filter.toLowerCase()));

  return (
    <div className="flex flex-col gap-4">
      <Panel>
        <StatRow>
          <Stat label="parts" value={parts.length} />
          <Stat label="xml parts" value={xml} />
          <Stat label="rels parts" value={parts.filter((part) => part.rels).length} />
          <Stat label="inflated" value={bytes(total)} />
          <Stat
            label="dangling rels"
            value={dangling.length}
            tone={dangling.length === 0 ? 'good' : 'warn'}
          />
        </StatRow>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[1.4fr_1fr]">
        <Panel
          title="Parts"
          hint="Pick one to see what it points at."
          aside={
            <TextInput
              aria-label="Filter parts"
              placeholder="filter…"
              value={filter}
              onChange={(event) => setFilter(event.target.value)}
              mono
              className="w-40"
            />
          }
        >
          <div className="max-h-[28rem] overflow-y-auto">
            <Table head={['part', 'content type', 'size']} caption="Parts in the package">
              {shown.map((part) => (
                <Row
                  key={part.name}
                  selected={part.name === selected}
                  onClick={() => setSelected(part.name)}
                >
                  <Cell>
                    <Mono>{part.name}</Mono>
                  </Cell>
                  <Cell className="text-[11px] text-fg-muted">
                    {part.contentType.replace(
                      'application/vnd.openxmlformats-officedocument.',
                      '…',
                    )}
                    {part.origin === null ? null : (
                      <span className="ml-1.5 text-fg-faint">({part.origin})</span>
                    )}
                  </Cell>
                  <Cell className="tabular whitespace-nowrap">{bytes(part.size)}</Cell>
                </Row>
              ))}
            </Table>
          </div>
        </Panel>

        <div className="flex flex-col gap-4">
          <Panel
            title="Relationships"
            hint={
              selected === null
                ? 'Select a part.'
                : 'From that part’s own .rels - rIds are scoped to one file.'
            }
          >
            {selected === null || detail === null ? (
              <p className="p-4 text-[13px] text-fg-muted">
                A slide does not name its layout; only <Mono tone="dim">_rels/slideN.xml.rels</Mono>{' '}
                does. Pick a slide to see it.
              </p>
            ) : detail.relationships.length === 0 ? (
              <p className="p-4 text-[13px] text-fg-muted">This part declares no relationships.</p>
            ) : (
              <Table head={['rId', 'type', 'target']} caption={`Relationships of ${selected}`}>
                {detail.relationships.map((rel) => (
                  <Row key={rel.id}>
                    <Cell className="whitespace-nowrap">
                      <Mono>{rel.id}</Mono>
                    </Cell>
                    <Cell className="text-[11px]">
                      {shortType(rel.type)}{' '}
                      {rel.targetMode === 'External' ? <Badge tone="warn">external</Badge> : null}
                    </Cell>
                    <Cell>
                      <Mono tone="dim">{rel.target}</Mono>
                    </Cell>
                  </Row>
                ))}
              </Table>
            )}
          </Panel>

          {detail === null ? null : (
            <Panel
              title={
                detail.digest === null ? 'The first 600 bytes' : 'SHA-256, and the first 64 bytes'
              }
              hint={
                detail.digest === null
                  ? 'An XML part is compared canonically on a round trip.'
                  : 'Binary parts are compared by digest, never re-encoded.'
              }
            >
              <div className="p-4">
                {detail.digest === null ? null : <Mono>{detail.digest}</Mono>}
                <pre className="mt-2 overflow-x-auto rounded-control bg-sunken p-3 font-mono text-[11px] leading-relaxed text-fg-muted">
                  {detail.text ?? detail.hex}
                </pre>
              </div>
            </Panel>
          )}

          {full ? (
            <Panel
              title="Reader limits"
              hint="What PartStore.open refuses, so a hostile file cannot take the tab down."
            >
              <Table head={['limit', 'default']} caption="Default ZIP limits">
                {LIMITS.map((limit) => (
                  <Row key={limit.label}>
                    <Cell>{limit.label}</Cell>
                    <Cell className="tabular">{limit.value}</Cell>
                  </Row>
                ))}
              </Table>
            </Panel>
          ) : null}
        </div>
      </div>

      {dangling.length === 0 ? null : (
        <Panel title="Dangling relationships" hint="Edges whose target is not in the package.">
          <Table head={['source', 'rId', 'target']} caption="Dangling relationships">
            {dangling.map((edge) => (
              <Row key={`${edge.source}-${edge.relationship.id}`}>
                <Cell>
                  <Mono tone="dim">{edge.source}</Mono>
                </Cell>
                <Cell>
                  <Mono>{edge.relationship.id}</Mono>
                </Cell>
                <Cell>
                  <Mono>{edge.target}</Mono>
                </Cell>
              </Row>
            ))}
          </Table>
        </Panel>
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
              This page reads. Adding, replacing and removing parts is the writer&apos;s job, and
              ZIP64 archives are refused rather than read.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

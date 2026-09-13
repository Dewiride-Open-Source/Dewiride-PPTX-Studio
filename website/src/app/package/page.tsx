'use client';

import { useMemo, useState } from 'react';

import { isRelationshipPartName, sha256Hex } from '@pptx-studio/opc';

import { useDocument } from '@/deck/document';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `import { PartStore } from '@pptx-studio/opc';

const store = PartStore.open(bytes);

store.partNames;                    // every part in the package
store.contentTypeOf('/ppt/slides/slide1.xml');
store.info(name)?.size;             // uncompressed bytes

// Relationships are scoped to one part's own .rels, never global.
const rels = store.relationships('/ppt/slides/slide1.xml');
rels.all;                           // { id, type, target, targetMode }
rels.resolve(rels.all[0]);          // the part it points at`;

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

export default function PackagePage() {
  const { parsed, error, loading } = useDocument();
  const [selected, setSelected] = useState<string | null>(null);

  const parts = useMemo(() => {
    if (parsed === null) return [];
    return parsed.store.partNames
      .map((name) => ({
        name,
        contentType: parsed.store.contentTypeOf(name) ?? '(none)',
        size: parsed.store.info(name)?.size ?? 0,
        rels: isRelationshipPartName(name),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [parsed]);

  const relationships = useMemo(() => {
    if (parsed === null || selected === null) return [];
    if (isRelationshipPartName(selected)) return [];
    try {
      return parsed.store.relationships(selected).all;
    } catch {
      return [];
    }
  }, [parsed, selected]);

  const digest = useMemo(() => {
    if (parsed === null || selected === null) return null;
    const contentType = parsed.store.contentTypeOf(selected) ?? '';
    if (contentType.endsWith('+xml') || contentType.endsWith('/xml')) return null;
    try {
      return sha256Hex(parsed.store.read(selected));
    } catch {
      return null;
    }
  }, [parsed, selected]);

  if (loading) return <p className="text-sm text-ink-400">Reading the deck…</p>;
  if (error !== null) return <p className="text-sm text-handle">{error}</p>;
  if (parsed === null) return null;

  const total = parts.reduce((sum, part) => sum + part.size, 0);
  const xml = parts.filter((part) => part.contentType.endsWith('xml')).length;

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Package</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/opc</Mono> - the ZIP, the content types, and the relationship graph.
        </p>
      </header>

      <Panel>
        <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-4">
          <Stat label="parts" value={parts.length} />
          <Stat label="xml parts" value={xml} />
          <Stat label="rels parts" value={parts.filter((part) => part.rels).length} />
          <Stat label="inflated" value={bytes(total)} />
        </div>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[1.4fr_1fr]">
        <Panel title="Parts" hint="Click one to see what it points at.">
          <div className="max-h-[26rem] overflow-y-auto">
            <Table head={['part', 'content type', 'size']}>
              {parts.map((part) => (
                <Row
                  key={part.name}
                  selected={part.name === selected}
                  onClick={() => setSelected(part.name)}
                >
                  <Cell>
                    <Mono>{part.name}</Mono>
                  </Cell>
                  <Cell className="text-[11px] text-ink-400">
                    {part.contentType.replace(
                      'application/vnd.openxmlformats-officedocument.',
                      '…',
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
            {selected === null ? (
              <p className="p-4 text-[13px] text-ink-400">
                A slide does not name its layout; only <Mono tone="dim">_rels/slideN.xml.rels</Mono>{' '}
                does. Pick a slide to see it.
              </p>
            ) : relationships.length === 0 ? (
              <p className="p-4 text-[13px] text-ink-400">This part declares no relationships.</p>
            ) : (
              <Table head={['rId', 'type', 'target']}>
                {relationships.map((rel) => (
                  <Row key={rel.id}>
                    <Cell className="whitespace-nowrap">
                      <Mono>{rel.id}</Mono>
                    </Cell>
                    <Cell className="text-[11px]">
                      {shortType(rel.type)}
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

          {digest === null ? null : (
            <Panel title="SHA-256" hint="Binary parts are compared by digest, never re-encoded.">
              <p className="p-4">
                <Mono>{digest}</Mono>
              </p>
            </Panel>
          )}

          <Panel title="How">
            <div className="p-4">
              <Snippet code={HOW} />
            </div>
          </Panel>
        </div>
      </div>
    </div>
  );
}

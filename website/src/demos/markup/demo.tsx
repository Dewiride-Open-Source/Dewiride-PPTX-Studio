'use client';

import { useMemo, useState } from 'react';

import { isRelationshipPartName } from '@pptx-studio/opc';
import {
  XmlTokenizer,
  applyEdit,
  excerpt,
  firstDifference,
  parseXml,
  prefixMap,
  serializeXml,
  serializeXmlString,
  type XDocument,
  type XElement,
  type XmlEdit,
} from '@pptx-studio/xml';

import { useDocument } from '@/deck/document';
import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Select, TextInput } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState, Status } from '@/design/state';
import type { DemoProps } from '../registry';
import { TreeNode } from './tree';

const HOW = `import { parseXml, serializeXml, applyEdit } from '@pptx-studio/xml';

const doc = parseXml(store.read(part));
serializeXml(doc);                       // byte-identical: a clean node re-emits by slicing its source

const undo = applyEdit({ kind: 'setAttribute', element, qname: 'name', value: 'Renamed' });
serializeXml(doc);                       // the original bytes, except that one value
applyEdit(undo);                         // and back, exactly`;

interface PartCheck {
  readonly name: string;
  readonly inBytes: number;
  readonly identical: boolean;
  readonly tokens: number;
  readonly prefixes: readonly string[];
  readonly error: string | null;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let at = 0; at < a.byteLength; at++) if (a[at] !== b[at]) return false;
  return true;
}

interface Sandbox {
  readonly document: XDocument;
  readonly original: string;
  /** The inverse of the one edit applied, or null when the tree is as parsed. */
  readonly undo: XmlEdit | null;
  readonly version: number;
}

export default function MarkupDemo({ full }: DemoProps) {
  const { parsed, failure, loading } = useDocument();
  const [selected, setSelected] = useState<string | null>(null);
  const [element, setElement] = useState<XElement | null>(null);
  const [sandbox, setSandbox] = useState<Sandbox | null>(null);
  const [attribute, setAttribute] = useState('');
  const [value, setValue] = useState('');

  const checks = useMemo<readonly PartCheck[]>(() => {
    if (parsed === null) return [];
    return parsed.store.partNames
      .filter((name) => {
        if (isRelationshipPartName(name)) return true;
        const type = parsed.store.contentTypeOf(name) ?? '';
        return type.endsWith('+xml') || type.endsWith('/xml');
      })
      .sort((a, b) => a.localeCompare(b))
      .map((name) => {
        try {
          const original = parsed.store.read(name);
          const tree = parseXml(original);
          let tokens = 0;
          const tokenizer = new XmlTokenizer(new TextDecoder().decode(original));
          for (let t = tokenizer.next(); t !== undefined; t = tokenizer.next()) tokens += 1;
          return {
            name,
            inBytes: original.byteLength,
            identical: sameBytes(original, serializeXml(tree)),
            tokens,
            prefixes: [...prefixMap(tree).keys()].filter((one) => one !== ''),
            error: null,
          };
        } catch (cause) {
          return {
            name,
            inBytes: 0,
            identical: false,
            tokens: 0,
            prefixes: [],
            error: cause instanceof Error ? cause.message : String(cause),
          };
        }
      });
  }, [parsed]);

  const choose = (name: string) => {
    if (parsed === null) return;
    setSelected(name);
    setElement(null);
    setAttribute('');
    setValue('');
    const document = parseXml(parsed.store.read(name));
    setSandbox({ document, original: serializeXmlString(document), undo: null, version: 0 });
  };

  const pick = (found: XElement) => {
    setElement(found);
    const first = found.attributes[0];
    setAttribute(first?.qname ?? '');
    setValue(first?.value ?? '');
  };

  const edit = () => {
    if (sandbox === null || element === null || attribute === '') return;
    if (sandbox.undo !== null) applyEdit(sandbox.undo);
    const undo = applyEdit({ kind: 'setAttribute', element, qname: attribute, value });
    setSandbox({ ...sandbox, undo, version: sandbox.version + 1 });
  };

  const revert = () => {
    if (sandbox === null || sandbox.undo === null) return;
    applyEdit(sandbox.undo);
    setSandbox({ ...sandbox, undo: null, version: sandbox.version + 1 });
  };

  const diff = useMemo(() => {
    if (sandbox === null) return null;
    const now = serializeXmlString(sandbox.document);
    const at = firstDifference(sandbox.original, now);
    return {
      identical: at < 0,
      at,
      before: at < 0 ? '' : excerpt(sandbox.original, at, 40),
      after: at < 0 ? '' : excerpt(now, at, 40),
      bytes: new TextEncoder().encode(now).byteLength,
    };
  }, [sandbox]);

  if (loading) return <Status busy>Reading the deck…</Status>;
  if (failure !== null) return <ErrorState {...failure} />;
  if (parsed === null) return null;

  const green = checks.filter((one) => one.identical).length;
  const chosen = checks.find((one) => one.name === selected) ?? null;

  return (
    <div className="flex flex-col gap-4">
      <Panel
        title="Parse, then serialise"
        hint="Every XML part of this deck, read into a tree and written back out."
        aside={
          <Badge tone={green === checks.length ? 'good' : 'bad'}>
            {`${String(green)}/${String(checks.length)} byte-identical`}
          </Badge>
        }
      >
        <StatRow>
          <Stat label="xml parts" value={checks.length} />
          <Stat label="identical" value={green} tone={green === checks.length ? 'good' : 'bad'} />
          <Stat
            label="tokens"
            value={checks.reduce((sum, one) => sum + one.tokens, 0).toLocaleString()}
          />
          <Stat
            label="bytes in"
            value={checks.reduce((sum, one) => sum + one.inBytes, 0).toLocaleString()}
          />
        </StatRow>
        <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
          A clean node re-emits by slicing the buffer it was parsed from, so this is not a
          normalisation that happens to match - nothing was rewritten at all.
        </p>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[1fr_1.4fr]">
        <Panel title="Parts">
          <div className="max-h-[30rem] overflow-y-auto p-2">
            {checks.map((one) => (
              <button
                key={one.name}
                type="button"
                onClick={() => choose(one.name)}
                aria-pressed={one.name === selected}
                className={`flex w-full items-center justify-between gap-2 rounded-control px-2 py-1.5 text-left ${
                  one.name === selected ? 'bg-accent-soft' : 'hover:bg-sunken'
                }`}
              >
                <Mono>{one.name}</Mono>
                {one.error !== null ? (
                  <Badge tone="bad">error</Badge>
                ) : one.identical ? (
                  <Badge tone="good">=</Badge>
                ) : (
                  <Badge tone="bad">differs</Badge>
                )}
              </button>
            ))}
          </div>
        </Panel>

        <div className="flex flex-col gap-4">
          <Panel
            title={selected ?? 'Tree'}
            hint={
              chosen === null
                ? 'Select a part.'
                : `${String(chosen.tokens)} tokens · click an element to edit it`
            }
            aside={
              chosen === null || chosen.prefixes.length === 0 ? null : (
                <span className="font-mono text-[11px] text-fg-faint">
                  {chosen.prefixes.join(' ')}
                </span>
              )
            }
          >
            {sandbox === null ? (
              <p className="p-4 text-[13px] text-fg-muted">
                Pick a part to walk its tree. A node with no source span is marked{' '}
                <Badge tone="warn">dirty</Badge> - none are, until something edits one.
              </p>
            ) : (
              <div className="max-h-96 overflow-auto p-3" key={sandbox.version}>
                <TreeNode
                  element={sandbox.document.root}
                  depth={0}
                  selected={element}
                  onSelect={pick}
                />
              </div>
            )}
          </Panel>

          {sandbox === null || diff === null ? null : (
            <Panel
              title="Edit one attribute, and undo it"
              hint="An edit marks one node dirty; its inverse comes back from applyEdit."
              aside={
                <Badge tone={diff.identical ? 'good' : 'warn'}>
                  {diff.identical ? 'byte-identical' : `differs at byte ${String(diff.at)}`}
                </Badge>
              }
            >
              <div className="flex flex-col gap-3 p-4">
                {element === null ? (
                  <p className="text-[13px] text-fg-muted">Click an element in the tree.</p>
                ) : (
                  <div className="flex flex-wrap items-end gap-2">
                    <Mono>{element.qname}</Mono>
                    <Select
                      aria-label="Attribute"
                      value={attribute}
                      onChange={(event) => {
                        setAttribute(event.target.value);
                        setValue(
                          element.attributes.find((a) => a.qname === event.target.value)?.value ??
                            '',
                        );
                      }}
                      className="font-mono text-[12px]"
                    >
                      {element.attributes.length === 0 ? (
                        <option value="">no attributes</option>
                      ) : null}
                      {element.attributes.map((attr) => (
                        <option key={attr.qname} value={attr.qname}>
                          {attr.qname}
                        </option>
                      ))}
                    </Select>
                    <TextInput
                      aria-label="New value"
                      value={value}
                      onChange={(event) => setValue(event.target.value)}
                      mono
                      className="w-48"
                    />
                    <Button size="sm" variant="primary" onClick={edit} disabled={attribute === ''}>
                      Apply
                    </Button>
                    <Button size="sm" onClick={revert} disabled={sandbox.undo === null}>
                      Undo
                    </Button>
                  </div>
                )}
                {diff.identical ? null : (
                  <div className="grid gap-2 font-mono text-[11px] sm:grid-cols-2">
                    <div className="rounded-control bg-sunken p-2">
                      <div className="mb-1 text-fg-faint">was</div>
                      <div className="break-all text-fg-muted">{diff.before}</div>
                    </div>
                    <div className="rounded-control bg-sunken p-2">
                      <div className="mb-1 text-fg-faint">now</div>
                      <div className="break-all text-fg">{diff.after}</div>
                    </div>
                  </div>
                )}
                {sandbox.undo === null ? null : (
                  <Code
                    lang="json"
                    title="the inverse applyEdit handed back"
                    code={JSON.stringify(
                      {
                        ...sandbox.undo,
                        element: (sandbox.undo as { element?: XElement }).element?.qname,
                      },
                      null,
                      2,
                    )}
                  />
                )}
              </div>
            </Panel>
          )}
        </div>
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
              The sandbox edits a copy of the part and exports nothing; the writer demo is where an
              edit becomes a file. Prefixes are never rewritten, and only UTF-8 parts are read.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

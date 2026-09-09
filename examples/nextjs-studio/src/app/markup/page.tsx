'use client';

import { useMemo, useState } from 'react';

import { isRelationshipPartName } from '@pptx-studio/opc';
import {
  XmlTokenizer,
  childElements,
  parseXml,
  prefixMap,
  serializeXml,
  type XElement,
} from '@pptx-studio/xml';

import { useDocument } from '@/deck/document';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';

const HOW = `import { parseXml, serializeXml } from '@pptx-studio/xml';

const tree = parseXml(store.read(part));
const out = serializeXml(tree);

// Byte-identical, because a node nobody edited re-emits by slicing the
// original buffer rather than being written out again.
out.every((byte, at) => byte === original[at]);`;

interface PartCheck {
  readonly name: string;
  readonly inBytes: number;
  readonly outBytes: number;
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

/** One element as an indented line, children beneath it. */
function TreeNode({ element, depth }: { element: XElement; depth: number }) {
  const [open, setOpen] = useState(depth < 2);
  const kids = childElements(element);
  return (
    <div style={{ paddingLeft: depth === 0 ? 0 : 14 }}>
      <button
        type="button"
        onClick={() => setOpen((on) => !on)}
        className="flex w-full items-baseline gap-1.5 rounded px-1 text-left hover:bg-ink-800"
      >
        <span className="w-3 shrink-0 font-mono text-[10px] text-ink-500">
          {kids.length === 0 ? '' : open ? '-' : '+'}
        </span>
        <span className="font-mono text-[12px] text-chrome">{element.qname}</span>
        {element.attributes.slice(0, 3).map((attr) => (
          <span key={attr.qname} className="font-mono text-[11px] text-ink-500">
            {attr.qname}=<span className="text-ink-400">&quot;{attr.value}&quot;</span>
          </span>
        ))}
        {element.attributes.length > 3 ? (
          <span className="font-mono text-[11px] text-ink-600">
            +{element.attributes.length - 3}
          </span>
        ) : null}
        {element.dirty ? <Badge tone="warn">dirty</Badge> : null}
      </button>
      {open ? kids.map((kid, at) => <TreeNode key={at} element={kid} depth={depth + 1} />) : null}
    </div>
  );
}

export default function MarkupPage() {
  const { parsed, error, loading } = useDocument();
  const [selected, setSelected] = useState<string | null>(null);

  const xmlParts = useMemo(() => {
    if (parsed === null) return [];
    return parsed.store.partNames
      .filter((name) => {
        if (isRelationshipPartName(name)) return true;
        const type = parsed.store.contentTypeOf(name) ?? '';
        return type.endsWith('+xml') || type.endsWith('/xml');
      })
      .sort((a, b) => a.localeCompare(b));
  }, [parsed]);

  const checks = useMemo<readonly PartCheck[]>(() => {
    if (parsed === null) return [];
    return xmlParts.map((name) => {
      try {
        const original = parsed.store.read(name);
        const tree = parseXml(original);
        const out = serializeXml(tree);
        let tokens = 0;
        const tokenizer = new XmlTokenizer(new TextDecoder().decode(original));
        for (let t = tokenizer.next(); t !== undefined; t = tokenizer.next()) tokens += 1;
        return {
          name,
          inBytes: original.byteLength,
          outBytes: out.byteLength,
          identical: sameBytes(original, out),
          tokens,
          prefixes: [...prefixMap(tree).keys()].filter((one) => one !== ''),
          error: null,
        };
      } catch (cause) {
        return {
          name,
          inBytes: 0,
          outBytes: 0,
          identical: false,
          tokens: 0,
          prefixes: [],
          error: cause instanceof Error ? cause.message : String(cause),
        };
      }
    });
  }, [parsed, xmlParts]);

  const tree = useMemo(() => {
    if (parsed === null || selected === null) return null;
    try {
      return parseXml(parsed.store.read(selected)).root;
    } catch {
      return null;
    }
  }, [parsed, selected]);

  if (loading) return <p className="text-sm text-ink-400">Reading the deck…</p>;
  if (error !== null) return <p className="text-sm text-handle">{error}</p>;
  if (parsed === null) return null;

  const green = checks.filter((one) => one.identical).length;
  const chosen = checks.find((one) => one.name === selected) ?? null;

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Markup</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/xml</Mono> - the tree, and the byte-identity gate the whole
          architecture rests on.
        </p>
      </header>

      <Panel
        title="Parse, then serialise"
        hint="Every XML part of this deck, read into a tree and written back out."
        aside={
          green === checks.length ? (
            <Badge tone="good">{`${String(green)}/${String(checks.length)} byte-identical`}</Badge>
          ) : (
            <Badge tone="bad">{`${String(green)}/${String(checks.length)} byte-identical`}</Badge>
          )
        }
      >
        <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-4">
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
        </div>
        <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-ink-500">
          A clean node re-emits by slicing the buffer it was parsed from, so this is not a
          normalisation that happens to match - nothing was rewritten at all.
        </p>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-[1fr_1.3fr]">
        <Panel title="Parts">
          <div className="max-h-[28rem] overflow-y-auto p-2">
            {checks.map((one) => (
              <button
                key={one.name}
                type="button"
                onClick={() => setSelected(one.name)}
                className={`flex w-full items-center justify-between gap-2 rounded px-2 py-1.5 text-left ${
                  one.name === selected ? 'bg-chrome/10' : 'hover:bg-ink-800'
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
            hint={chosen === null ? 'Select a part.' : `${String(chosen.tokens)} tokens`}
            aside={
              chosen === null || chosen.prefixes.length === 0 ? null : (
                <span className="font-mono text-[11px] text-ink-500">
                  {chosen.prefixes.join(' ')}
                </span>
              )
            }
          >
            {tree === null ? (
              <p className="p-4 text-[13px] text-ink-400">
                Pick a part to walk its tree. A node with no source span is marked{' '}
                <Badge tone="warn">dirty</Badge> - none are, until something edits one.
              </p>
            ) : (
              <div className="max-h-96 overflow-auto p-3">
                <TreeNode element={tree} depth={0} />
              </div>
            )}
          </Panel>

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

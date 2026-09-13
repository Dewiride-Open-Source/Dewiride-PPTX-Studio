'use client';

import { useState } from 'react';

import { childElements, type XElement } from '@pptx-studio/xml';

import { Badge } from '@/design/badge';

interface TreeNodeProps {
  readonly element: XElement;
  readonly depth: number;
  readonly selected: XElement | null;
  readonly onSelect: (element: XElement) => void;
}

/** One element as a line, its children beneath it; the byte span says where it came from. */
export function TreeNode({ element, depth, selected, onSelect }: TreeNodeProps) {
  const [open, setOpen] = useState(depth < 2);
  const kids = childElements(element);
  const chosen = element === selected;
  return (
    <div style={{ paddingLeft: depth === 0 ? 0 : 14 }}>
      <div
        className={`flex items-baseline gap-1.5 rounded-control px-1 ${chosen ? 'bg-accent-soft' : 'hover:bg-sunken'}`}
      >
        <button
          type="button"
          onClick={() => setOpen((on) => !on)}
          aria-label={kids.length === 0 ? undefined : open ? 'Collapse' : 'Expand'}
          aria-expanded={kids.length === 0 ? undefined : open}
          className="w-3 shrink-0 font-mono text-[10px] text-fg-faint"
        >
          {kids.length === 0 ? '' : open ? '−' : '+'}
        </button>
        <button
          type="button"
          onClick={() => onSelect(element)}
          className="flex min-w-0 flex-1 flex-wrap items-baseline gap-1.5 text-left"
        >
          <span className="font-mono text-[12px] text-accent">{element.qname}</span>
          {element.attributes.slice(0, 3).map((attr) => (
            <span key={attr.qname} className="font-mono text-[11px] text-fg-faint">
              {attr.qname}=<span className="text-fg-muted">&quot;{attr.value}&quot;</span>
            </span>
          ))}
          {element.attributes.length > 3 ? (
            <span className="font-mono text-[11px] text-fg-faint">
              +{element.attributes.length - 3}
            </span>
          ) : null}
          <span className="ml-auto font-mono text-[10px] text-fg-faint">
            {element.start}–{element.end}
          </span>
          {element.dirty ? <Badge tone="warn">dirty</Badge> : null}
        </button>
      </div>
      {open
        ? kids.map((kid, at) => (
            <TreeNode
              key={at}
              element={kid}
              depth={depth + 1}
              selected={selected}
              onSelect={onSelect}
            />
          ))
        : null}
    </div>
  );
}

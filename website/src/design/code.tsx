'use client';

import { DynamicCodeBlock } from 'fumadocs-ui/components/dynamic-codeblock';
import type { ReactNode } from 'react';

/** A part URI, a qname, an XPath - the identifiers this domain is written in. */
export function Mono({
  children,
  tone = 'plain',
}: {
  children: ReactNode;
  tone?: 'plain' | 'dim';
}) {
  return (
    <code
      className={`font-mono text-[12px] break-all ${tone === 'dim' ? 'text-fg-muted' : 'text-fg'}`}
    >
      {children}
    </code>
  );
}

/** A key the visitor can press. */
export function Kbd({ children }: { children: ReactNode }) {
  return (
    <kbd className="inline-block rounded border border-line-strong bg-sunken px-1.5 font-mono text-[11px] leading-5 text-fg-muted">
      {children}
    </kbd>
  );
}

export type CodeLang = 'ts' | 'tsx' | 'sh' | 'xml' | 'json' | 'txt';

interface CodeProps {
  readonly code: string;
  readonly lang?: CodeLang;
  readonly title?: string;
}

/** How a caller would write what a demo just did: highlighted, with a copy button. */
export function Code({ code, lang = 'ts', title }: CodeProps) {
  return (
    <div className="text-[12.5px] [&_figure]:my-0">
      <DynamicCodeBlock lang={lang} code={code} codeblock={{ title }} />
    </div>
  );
}

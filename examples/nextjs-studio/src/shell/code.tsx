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
      className={`font-mono text-[12px] ${tone === 'dim' ? 'text-ink-400' : 'text-ink-200'} break-all`}
    >
      {children}
    </code>
  );
}

/** How a caller would write what this tool just did. */
export function Snippet({ code }: { code: string }) {
  return (
    <pre className="overflow-x-auto rounded-md border border-ink-700 bg-ink-900 p-3 font-mono text-[12px] leading-relaxed text-ink-200">
      {code}
    </pre>
  );
}

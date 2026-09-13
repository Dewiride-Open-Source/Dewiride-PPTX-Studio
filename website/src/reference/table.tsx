import { Badge, type Tone } from '@/design/badge';
import type { PackageName } from '@/site/packages';
import { readReference, type SymbolKind } from './read';

const KIND_TONE: Record<SymbolKind, Tone> = {
  function: 'accent',
  class: 'info',
  interface: 'plain',
  type: 'plain',
  enum: 'plain',
  const: 'site',
  're-export': 'warn',
};

/** A JSDoc sentence: backticks become code, and bold markers go. */
function Doc({ text }: { text: string }) {
  const parts = text.replaceAll('**', '').split(/(`[^`]*`)/);
  return (
    <span className="min-w-0 flex-1 text-fg-muted">
      {parts.map((part, at) =>
        part.startsWith('`') ? (
          <code key={at} className="font-mono text-[12px] text-fg">
            {part.slice(1, -1)}
          </code>
        ) : (
          <span key={at}>{part}</span>
        ),
      )}
    </span>
  );
}

/** Every export of one package, grouped by the source file it was declared in, from the installed d.ts. */
export function ApiTable({ pkg }: { pkg: PackageName }) {
  const reference = readReference(pkg);
  return (
    <div className="not-prose space-y-6">
      <p className="text-[13px] text-fg-muted">
        {reference.exported} exports of <code className="font-mono">@pptx-studio/{pkg}</code>{' '}
        {reference.version}, read from the package this site installed. Click a name for its
        declaration.
      </p>
      {reference.groups.map((group) => (
        <section key={group.file} id={`file-${group.file.replace(/[^a-z0-9]+/gi, '-')}`}>
          <h3 className="mb-2 font-mono text-[12px] font-medium text-fg-muted">{group.file}</h3>
          <ul className="divide-y divide-line rounded-panel border border-line">
            {group.symbols.map((symbol) => (
              <li key={symbol.name} id={`${pkg}-${symbol.name}`}>
                <details className="group">
                  <summary className="flex cursor-pointer flex-wrap items-baseline gap-x-3 gap-y-1 px-3 py-2 text-[13px] hover:bg-sunken">
                    <code className="font-mono font-medium text-fg">{symbol.name}</code>
                    <Badge tone={KIND_TONE[symbol.kind]}>
                      {symbol.isType ? 'type' : symbol.kind}
                    </Badge>
                    {symbol.doc === '' ? null : <Doc text={symbol.doc} />}
                  </summary>
                  <pre className="overflow-x-auto border-t border-line bg-sunken px-3 py-2 font-mono text-[12px] leading-relaxed text-fg">
                    {symbol.signature}
                  </pre>
                </details>
              </li>
            ))}
          </ul>
        </section>
      ))}
    </div>
  );
}

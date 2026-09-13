import type { ReactNode } from 'react';

type CalloutKind = 'note' | 'warn' | 'honest';

const RAIL: Record<CalloutKind, string> = {
  note: 'border-l-accent bg-accent/5',
  warn: 'border-l-warn bg-warn/5',
  honest: 'border-l-warn bg-warn/5',
};

interface CalloutProps {
  readonly kind?: CalloutKind;
  readonly title?: string;
  readonly children: ReactNode;
}

/** A short aside. The `honest` kind always says what a page does not do. */
export function Callout({ kind = 'note', title, children }: CalloutProps) {
  const heading = kind === 'honest' ? 'What this page does not do' : title;
  return (
    <aside
      className={`rounded-panel border border-line border-l-4 px-4 py-3 text-[13px] leading-relaxed text-fg-muted ${RAIL[kind]}`}
    >
      {heading === undefined ? null : (
        <p className="mb-1 text-[12px] font-semibold tracking-wide text-fg uppercase">{heading}</p>
      )}
      <div className="space-y-2">{children}</div>
    </aside>
  );
}

import type { ReactNode } from 'react';

import { TONE_TEXT, type Tone } from './badge';

interface PanelProps {
  readonly title?: string;
  readonly hint?: ReactNode;
  readonly aside?: ReactNode;
  readonly children: ReactNode;
  readonly className?: string;
  readonly id?: string;
}

/** A bordered section with an optional header row: the unit every demo is laid out in. */
export function Panel({ title, hint, aside, children, className = '', id }: PanelProps) {
  return (
    <section
      id={id}
      className={`rounded-panel border border-line bg-surface shadow-panel ${className}`}
    >
      {title === undefined && aside === undefined ? null : (
        <header className="flex items-baseline justify-between gap-4 border-b border-line px-4 py-3">
          <div className="min-w-0">
            {title === undefined ? null : (
              <h2 className="text-sm font-semibold text-fg">{title}</h2>
            )}
            {hint === undefined ? null : <p className="mt-0.5 text-xs text-fg-muted">{hint}</p>}
          </div>
          {aside}
        </header>
      )}
      {children}
    </section>
  );
}

/** A labelled figure, for the counts every demo reports. */
export function Stat({
  label,
  value,
  tone = 'plain',
}: {
  label: string;
  value: string | number;
  tone?: Tone;
}) {
  return (
    <div className="min-w-0">
      <div
        className={`tabular text-xl font-semibold ${tone === 'plain' ? 'text-fg' : TONE_TEXT[tone]}`}
      >
        {value}
      </div>
      <div className="mt-0.5 text-[11px] tracking-wide text-fg-muted uppercase">{label}</div>
    </div>
  );
}

/** A row of stats with the padding a panel expects. */
export function StatRow({ children }: { children: ReactNode }) {
  return <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-3 lg:grid-cols-5">{children}</div>;
}

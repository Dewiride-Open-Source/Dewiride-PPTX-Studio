import type { ReactNode } from 'react';

export function Panel({
  title,
  hint,
  aside,
  children,
  className = '',
}: {
  title?: string;
  hint?: string;
  aside?: ReactNode;
  children: ReactNode;
  className?: string;
}) {
  return (
    <section className={`rounded-lg border border-ink-700 bg-ink-850 ${className}`}>
      {title === undefined && aside === undefined ? null : (
        <header className="flex items-baseline justify-between gap-4 border-b border-ink-700 px-4 py-3">
          <div className="min-w-0">
            {title === undefined ? null : (
              <h2 className="text-sm font-semibold text-ink-100">{title}</h2>
            )}
            {hint === undefined ? null : <p className="mt-0.5 text-xs text-ink-400">{hint}</p>}
          </div>
          {aside}
        </header>
      )}
      {children}
    </section>
  );
}

/** A labelled figure, for the counts every tool reports. */
export function Stat({
  label,
  value,
  tone = 'plain',
}: {
  label: string;
  value: string | number;
  tone?: 'plain' | 'good' | 'warn' | 'bad';
}) {
  const color =
    tone === 'good'
      ? 'text-good'
      : tone === 'warn'
        ? 'text-widened'
        : tone === 'bad'
          ? 'text-handle'
          : 'text-ink-100';
  return (
    <div className="min-w-0">
      <div className={`tabular text-xl font-semibold ${color}`}>{value}</div>
      <div className="mt-0.5 text-[11px] tracking-wide text-ink-400 uppercase">{label}</div>
    </div>
  );
}

import type { ReactNode } from 'react';

export type Tone = 'plain' | 'accent' | 'info' | 'site' | 'good' | 'warn' | 'bad';

export const TONE_TEXT: Record<Tone, string> = {
  plain: 'text-fg-muted',
  accent: 'text-accent',
  info: 'text-info',
  site: 'text-site',
  good: 'text-good',
  warn: 'text-warn',
  bad: 'text-bad',
};

const TONE_PILL: Record<Tone, string> = {
  plain: 'border-line-strong text-fg-muted',
  accent: 'border-accent/40 bg-accent/10 text-accent',
  info: 'border-info/40 bg-info/10 text-info',
  site: 'border-site/40 bg-site/10 text-site',
  good: 'border-good/40 bg-good/10 text-good',
  warn: 'border-warn/40 bg-warn/10 text-warn',
  bad: 'border-bad/40 bg-bad/10 text-bad',
};

/** A word in a pill: a status, a kind, an origin. Never colour alone. */
export function Badge({ tone = 'plain', children }: { tone?: Tone; children: ReactNode }) {
  return (
    <span
      className={`inline-flex items-center rounded-full border px-2 py-0.5 font-mono text-[11px] leading-4 whitespace-nowrap ${TONE_PILL[tone]}`}
    >
      {children}
    </span>
  );
}

interface ChipProps {
  readonly tone?: Tone;
  readonly selected?: boolean;
  readonly onClick?: () => void;
  readonly onRemove?: () => void;
  readonly children: ReactNode;
}

/** A badge that can be pressed or taken away: a filter, an applied transform. */
export function Chip({ tone = 'plain', selected = false, onClick, onRemove, children }: ChipProps) {
  const look = selected ? 'border-accent bg-accent-soft text-fg' : TONE_PILL[tone];
  const body = (
    <span className="inline-flex items-center gap-1.5">
      {children}
      {onRemove === undefined ? null : (
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            onRemove();
          }}
          aria-label="Remove"
          className="-mr-1 rounded-full px-1 leading-none hover:bg-fg/10"
        >
          ×
        </button>
      )}
    </span>
  );
  const className = `inline-flex items-center rounded-full border px-2.5 py-1 text-[12px] leading-4 whitespace-nowrap transition-colors ${look}`;
  return onClick === undefined ? (
    <span className={className}>{body}</span>
  ) : (
    <button type="button" onClick={onClick} aria-pressed={selected} className={className}>
      {body}
    </button>
  );
}

import type { ReactNode } from 'react';

export type Tone = 'plain' | 'good' | 'warn' | 'bad' | 'info';

const TONES: Record<Tone, string> = {
  plain: 'border-ink-600 bg-ink-800 text-ink-300',
  good: 'border-good/40 bg-good/10 text-good',
  warn: 'border-widened/40 bg-widened/10 text-widened',
  bad: 'border-handle/40 bg-handle/10 text-handle',
  info: 'border-chrome/40 bg-chrome/10 text-chrome',
};

export function Badge({ tone = 'plain', children }: { tone?: Tone; children: ReactNode }) {
  return (
    <span
      className={`inline-flex shrink-0 items-center rounded border px-1.5 py-0.5 font-mono text-[11px] ${TONES[tone]}`}
    >
      {children}
    </span>
  );
}

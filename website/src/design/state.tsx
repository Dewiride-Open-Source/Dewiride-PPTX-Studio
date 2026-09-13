import type { ReactNode } from 'react';

interface EmptyStateProps {
  readonly title: string;
  readonly children?: ReactNode;
  readonly action?: ReactNode;
}

/** Nothing to show yet, and what would change that. */
export function EmptyState({ title, children, action }: EmptyStateProps) {
  return (
    <div className="flex flex-col items-center gap-3 px-6 py-12 text-center">
      <p className="text-sm font-medium text-fg">{title}</p>
      {children === undefined ? null : (
        <div className="max-w-md text-[13px] leading-relaxed text-fg-muted">{children}</div>
      )}
      {action}
    </div>
  );
}

export interface Failure {
  /** A machine-readable code when the cause carried one, e.g. `ERR_NOT_A_ZIP`. */
  readonly code?: string;
  readonly message: string;
  /** What the visitor can do about it, in plain words. */
  readonly advice?: string;
}

/** Something failed: the code, the message, and what to do. */
export function ErrorState({ code, message, advice }: Failure) {
  return (
    <div role="alert" className="rounded-panel border border-bad/40 bg-bad/5 px-4 py-3 text-[13px]">
      <p className="text-fg">
        {code === undefined ? null : (
          <code className="mr-2 font-mono text-[12px] text-bad">{code}</code>
        )}
        {message}
      </p>
      {advice === undefined ? null : <p className="mt-1 text-fg-muted">{advice}</p>}
    </div>
  );
}

/** A grey slide-shaped box while a slide is on its way. */
export function Skeleton({
  aspect = 16 / 9,
  className = '',
}: {
  aspect?: number;
  className?: string;
}) {
  return (
    <div
      aria-hidden="true"
      className={`animate-pulse rounded-[2px] bg-sunken ${className}`}
      style={{ aspectRatio: String(aspect) }}
    />
  );
}

/** A one-line status the screen reader also hears. */
export function Status({ children, busy = false }: { children: ReactNode; busy?: boolean }) {
  return (
    <p role="status" aria-live="polite" aria-busy={busy} className="text-[12px] text-fg-muted">
      {busy ? (
        <span
          aria-hidden="true"
          className="mr-2 inline-block h-3 w-3 animate-spin rounded-full border-2 border-line-strong border-t-accent align-[-2px]"
        />
      ) : null}
      {children}
    </p>
  );
}

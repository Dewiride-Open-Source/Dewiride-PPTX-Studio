'use client';

import Link from 'next/link';
import { Component, type ErrorInfo, type ReactNode } from 'react';

import { DeckPicker } from '@/deck/picker';
import { describeFailure } from '@/deck/failures';
import { Badge } from '@/design/badge';
import { ErrorState } from '@/design/state';
import { SiteError } from '@/site/errors';
import { LOADED } from './loaded';
import { demoNamed, type DemoEntry } from './registry';

interface BoundaryState {
  readonly failed: unknown;
}

/** A demo that throws stays a broken panel, not a broken page. */
class DemoBoundary extends Component<{ children: ReactNode; name: string }, BoundaryState> {
  override state: BoundaryState = { failed: undefined };

  static getDerivedStateFromError(failed: unknown): BoundaryState {
    return { failed };
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error(`demo ${this.props.name} failed`, error, info.componentStack);
  }

  override render(): ReactNode {
    if (this.state.failed !== undefined)
      return <ErrorState {...describeFailure(this.state.failed)} />;
    return this.props.children;
  }
}

interface DemoProps {
  readonly name: string;
  /** True on the demo's own page; false inside a docs page. */
  readonly full?: boolean;
}

/** One implementation, two places: the demo's own page and its package's docs page. */
export function Demo({ name, full = false }: DemoProps) {
  const entry = demoNamed(name);
  if (entry === undefined) throw new SiteError('SITE_UNKNOWN_DEMO', `no demo called ${name}`);
  return <LoadedDemo entry={entry} full={full} />;
}

function LoadedDemo({ entry, full }: { entry: DemoEntry; full: boolean }) {
  const Loaded = LOADED[entry.name];
  return (
    <div className="flex flex-col gap-4">
      {full ? null : (
        <div className="flex flex-wrap items-center justify-between gap-2">
          <DeckPicker compact />
          <Link href={`/demos/${entry.name}`} className="text-[12px] text-accent hover:underline">
            Open the full demo →
          </Link>
        </div>
      )}
      <DemoBoundary name={entry.name}>
        <Loaded full={full} />
      </DemoBoundary>
    </div>
  );
}

/** The header every demo page and docs embed opens with. */
export function DemoHeader({ entry }: { entry: DemoEntry }) {
  return (
    <header className="flex flex-wrap items-end justify-between gap-3">
      <div>
        <h1 className="text-xl font-semibold tracking-tight">{entry.title}</h1>
        <p className="mt-1 text-[13px] text-fg-muted">
          <code className="font-mono text-fg">@pptx-studio/{entry.name}</code> · {entry.blurb}
        </p>
      </div>
      <div className="flex items-center gap-2">
        {entry.prerendered === true ? (
          <Badge tone="warn">rendered at build time</Badge>
        ) : (
          <Badge tone="good">runs in your tab</Badge>
        )}
        <Link
          href={`/docs/packages/${entry.name}`}
          className="text-[12px] text-accent hover:underline"
        >
          Docs →
        </Link>
      </div>
    </header>
  );
}

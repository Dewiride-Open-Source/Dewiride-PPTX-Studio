'use client';

import Link from 'next/link';
import { Component, useEffect, useRef, useState, type ErrorInfo, type ReactNode } from 'react';

import { DeckPicker } from '@/deck/picker';
import { describeFailure } from '@/deck/failures';
import { Badge } from '@/design/badge';
import { ErrorState, Skeleton } from '@/design/state';
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

/** True once the element is within a screen of the viewport, and from then on. */
function useNear(ref: { readonly current: HTMLElement | null }, always: boolean): boolean {
  const [near, setNear] = useState(always);
  useEffect(() => {
    const element = ref.current;
    if (always || element === null) return;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((entry) => entry.isIntersecting)) {
          setNear(true);
          observer.disconnect();
        }
      },
      { rootMargin: '100% 0px' },
    );
    observer.observe(element);
    return () => observer.disconnect();
  }, [ref, always]);
  return near;
}

function LoadedDemo({ entry, full }: { entry: DemoEntry; full: boolean }) {
  const Loaded = LOADED[entry.name];
  const root = useRef<HTMLDivElement>(null);
  // A docs page reads first; the demo's chunk and the deck parse wait until the reader is near it.
  const near = useNear(root, full);
  return (
    <div ref={root} className="flex flex-col gap-4">
      {full ? null : (
        <div className="flex flex-wrap items-center justify-between gap-2">
          <DeckPicker compact />
          <Link href={`/demos/${entry.name}`} className="text-[12px] text-accent hover:underline">
            Open the full demo →
          </Link>
        </div>
      )}
      <DemoBoundary name={entry.name}>
        {near ? <Loaded full={full} /> : <Skeleton className="w-full" />}
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
        <p className="mt-1 max-w-2xl text-[13px] text-fg-muted">{entry.blurb}</p>
        <p className="mt-1 text-[13px] text-fg">
          <span className="font-semibold">Try:</span> {entry.tryThis}
          <code className="ml-2 font-mono text-[12px] text-fg-faint">
            @pptx-studio/{entry.name}
          </code>
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

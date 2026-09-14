import type { Metadata } from 'next';
import Link from 'next/link';

import { Badge } from '@/design/badge';
import { DEMOS } from '@/demos/registry';
import { packageFacts, RUNTIME_LABELS } from '@/site/packages';

export const metadata: Metadata = {
  title: 'Demos',
  description: 'One live demo per package, on the sample decks or on a .pptx you drop on the page.',
};

export default function DemosIndex() {
  return (
    <div className="flex flex-col gap-6">
      <header>
        <h1 className="text-xl font-semibold tracking-tight">Twelve demos, one per package</h1>
        <p className="mt-1 max-w-2xl text-[13px] text-fg-muted">
          Every one runs on the deck chosen above - a shipped sample or one of your own. Nothing is
          uploaded; the CLI page shows what Node rendered when the site was built.
        </p>
      </header>
      <ul className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">
        {DEMOS.map((demo) => {
          const facts = packageFacts(demo.name);
          return (
            <li key={demo.name}>
              <Link
                href={`/demos/${demo.name}`}
                className="flex h-full flex-col rounded-panel border border-line bg-surface p-4 shadow-panel transition-colors hover:border-accent"
              >
                <div className="flex items-baseline justify-between gap-2">
                  <span className="text-sm font-semibold text-fg">{demo.title}</span>
                  <code className="font-mono text-[11px] text-fg-faint">{demo.name}</code>
                </div>
                <p className="mt-2 flex-1 text-[13px] leading-relaxed text-fg-muted">
                  {demo.blurb}
                </p>
                <p className="mt-2 text-[12px] leading-relaxed text-fg">
                  <span className="font-semibold">Try:</span> {demo.tryThis}
                </p>
                <div className="mt-3 flex flex-wrap gap-1.5">
                  {facts.runsIn.map((runtime) => (
                    <Badge key={runtime}>{RUNTIME_LABELS[runtime]}</Badge>
                  ))}
                  {demo.prerendered === true ? <Badge tone="warn">build time</Badge> : null}
                </div>
              </Link>
            </li>
          );
        })}
      </ul>
    </div>
  );
}

import Link from 'next/link';

import { Badge } from '@/design/badge';
import { GROUP_TITLES, PACKAGES, RUNTIME_LABELS, type PackageGroup } from '@/site/packages';

const ORDER: readonly PackageGroup[] = ['read', 'draw', 'write', 'command-line'];

export function PackageGrid() {
  return (
    <section aria-labelledby="twelve-packages" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <h2 id="twelve-packages" className="text-2xl font-semibold tracking-tight">
        Twelve packages
      </h2>
      <p className="mt-2 max-w-2xl text-[14px] text-fg-muted">
        Install only what you import from; the rest arrive as dependencies. Every one is ESM with
        its own types.
      </p>
      {ORDER.map((group) => (
        <div key={group} className="mt-8">
          <h3 className="text-[12px] font-semibold tracking-wider text-fg-muted uppercase">
            {GROUP_TITLES[group]}
          </h3>
          <ul className="mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {PACKAGES.filter((one) => one.group === group).map((one) => (
              <li
                key={one.name}
                className="flex flex-col rounded-panel border border-line bg-surface p-4 shadow-panel"
              >
                <div className="flex items-baseline justify-between gap-2">
                  <code className="font-mono text-[13px] font-medium text-fg">
                    @pptx-studio/{one.name}
                  </code>
                  <span className="font-mono text-[11px] text-fg-faint">{one.version}</span>
                </div>
                <p className="mt-2 flex-1 text-[13px] leading-relaxed text-fg-muted">{one.blurb}</p>
                <div className="mt-3 flex flex-wrap items-center gap-1.5">
                  {one.runsIn.map((runtime) => (
                    <Badge key={runtime}>{RUNTIME_LABELS[runtime]}</Badge>
                  ))}
                  <span className="ml-auto flex gap-3 text-[12px]">
                    <Link
                      href={`/docs/packages/${one.name}`}
                      className="text-accent hover:underline"
                    >
                      Docs
                    </Link>
                    <Link href={`/demos/${one.name}`} className="text-accent hover:underline">
                      Demo
                    </Link>
                  </span>
                </div>
              </li>
            ))}
          </ul>
        </div>
      ))}
    </section>
  );
}

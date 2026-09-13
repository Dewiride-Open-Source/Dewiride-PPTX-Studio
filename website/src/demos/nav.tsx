'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';

import { DEMOS } from './registry';

/** The twelve demos, one per package, with the current one marked. */
export function DemosNav() {
  const here = usePathname().replace(/\/$/, '');
  return (
    <nav aria-label="Demos" className="flex flex-col gap-0.5">
      {DEMOS.map((demo) => {
        const href = `/demos/${demo.name}`;
        const on = here === href;
        return (
          <Link
            key={demo.name}
            href={href}
            aria-current={on ? 'page' : undefined}
            className={`rounded-control px-3 py-1.5 transition-colors ${
              on ? 'bg-accent-soft text-fg' : 'text-fg-muted hover:bg-sunken hover:text-fg'
            }`}
          >
            <span className="block text-[13px] font-medium">{demo.title}</span>
            <span className="block font-mono text-[11px] text-fg-faint">{demo.name}</span>
          </Link>
        );
      })}
    </nav>
  );
}

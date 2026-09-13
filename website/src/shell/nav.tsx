'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';

import { TOOLS } from './tools';

export function Nav() {
  const here = usePathname();
  return (
    <nav className="flex flex-col gap-0.5 p-3">
      <Link
        href="/"
        className={`rounded px-3 py-2 text-sm ${
          here === '/' ? 'bg-ink-700 text-ink-100' : 'text-ink-300 hover:bg-ink-800'
        }`}
      >
        Overview
      </Link>
      <p className="mt-4 mb-1 px-3 text-[11px] tracking-wider text-ink-500 uppercase">
        One tool per package
      </p>
      {TOOLS.map((tool) => {
        const on = here === tool.href;
        return (
          <Link
            key={tool.href}
            href={tool.href}
            className={`group rounded px-3 py-2 ${on ? 'bg-ink-700' : 'hover:bg-ink-800'}`}
          >
            <span className={`block text-sm ${on ? 'text-ink-100' : 'text-ink-200'}`}>
              {tool.name}
              {tool.server === true ? (
                <span className="ml-1.5 align-middle text-[10px] text-widened">server</span>
              ) : null}
            </span>
            <span className="block font-mono text-[11px] text-ink-500">{tool.package}</span>
          </Link>
        );
      })}
    </nav>
  );
}

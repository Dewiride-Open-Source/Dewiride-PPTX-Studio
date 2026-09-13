import Link from 'next/link';

import { Mark } from './brand';
import { VersionBadge } from './version-badge';
import { NPM_ORG, REPOSITORY } from '@/site/navigation';

const COLUMNS: readonly { title: string; links: readonly { text: string; href: string }[] }[] = [
  {
    title: 'Site',
    links: [
      { text: 'Docs', href: '/docs' },
      { text: 'Demos', href: '/demos' },
      { text: 'Playground', href: '/playground' },
      { text: 'Status and roadmap', href: '/docs/status' },
    ],
  },
  {
    title: 'Project',
    links: [
      { text: 'GitHub', href: REPOSITORY },
      { text: 'npm', href: NPM_ORG },
      { text: 'Licence: Apache-2.0', href: `${REPOSITORY}/blob/main/LICENSE` },
      { text: 'Security policy', href: `${REPOSITORY}/blob/main/SECURITY.md` },
    ],
  },
];

export function Footer() {
  return (
    <footer className="border-t border-line bg-surface">
      <div className="mx-auto grid max-w-6xl gap-8 px-4 py-10 sm:px-6 md:grid-cols-[1.4fr_1fr_1fr]">
        <div className="space-y-3 text-[13px] text-fg-muted">
          <p className="inline-flex items-center gap-2 font-semibold text-fg">
            <Mark /> PPTX Studio
          </p>
          <p>
            No cookies, no analytics, no third-party requests. A deck you open here never leaves
            your browser.
          </p>
          <p>
            Pre-alpha. Built against <VersionBadge /> and eleven siblings; the site is rebuilt after
            each release and can sit one release behind npm.
          </p>
          <p className="text-fg-faint">
            Not affiliated with or endorsed by Microsoft. PowerPoint is a trademark of Microsoft
            Corporation.
          </p>
        </div>
        {COLUMNS.map((column) => (
          <nav key={column.title} aria-label={column.title} className="text-[13px]">
            <p className="mb-2 font-semibold text-fg">{column.title}</p>
            <ul className="space-y-1.5">
              {column.links.map((link) => (
                <li key={link.href}>
                  <Link
                    href={link.href}
                    className="text-fg-muted hover:text-accent"
                    {...(/^https?:/.test(link.href) ? { target: '_blank', rel: 'noreferrer' } : {})}
                  >
                    {link.text}
                  </Link>
                </li>
              ))}
            </ul>
          </nav>
        ))}
      </div>
    </footer>
  );
}

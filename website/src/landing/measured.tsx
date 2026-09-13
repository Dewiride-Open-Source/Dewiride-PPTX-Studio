import Link from 'next/link';

import { plan } from '@/status/plan';

const TILES: readonly { value: string; label: string; href: string }[] = [
  {
    value: '187',
    label: 'preset shapes transcoded and cross-checked',
    href: '/docs/packages/geometry',
  },
  {
    value: '454/455',
    label: 'colour swatches predicted exactly, read back from PowerPoint two ways',
    href: '/docs/packages/paint',
  },
  { value: '2212/2212', label: 'line-break probes fitted', href: '/docs/packages/text' },
  {
    value: '29',
    label: 'rules, each one a package PowerPoint refused',
    href: '/docs/packages/validate',
  },
  {
    value: `${String(plan.roundTrip.decks)}/${String(plan.roundTrip.total)}`,
    label: 'corpus decks round-trip in CI',
    href: '/docs/status',
  },
];

export function Measured() {
  return (
    <section aria-labelledby="measured" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <h2 id="measured" className="text-2xl font-semibold tracking-tight">
        Measured, not assumed
      </h2>
      <p className="mt-2 max-w-3xl text-[14px] text-fg-muted">
        Every rule in these packages is either cited to ECMA-376 by clause or measured against
        PowerPoint 16.0.20326 and committed as a fixture. The reasoning lives in one architecture
        decision record per sub-phase.
      </p>
      <ul className="mt-6 grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
        {TILES.map((tile) => (
          <li key={tile.label}>
            <Link
              href={tile.href}
              className="block h-full rounded-panel border border-line bg-surface p-4 shadow-panel transition-colors hover:border-accent"
            >
              <div className="tabular font-mono text-2xl font-semibold text-fg">{tile.value}</div>
              <div className="mt-1 text-[12px] leading-snug text-fg-muted">{tile.label}</div>
            </Link>
          </li>
        ))}
      </ul>
    </section>
  );
}

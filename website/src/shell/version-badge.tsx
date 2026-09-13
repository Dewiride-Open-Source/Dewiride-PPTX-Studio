import { PACKAGES, packageFacts } from '@/site/packages';

/** The versions this site was built against, so a visitor can tell how far behind npm it is. */
export function VersionBadge() {
  const lead = packageFacts('render-svg');
  const all = PACKAGES.map((one) => `@pptx-studio/${one.name} ${one.version}`).join('\n');
  return (
    <span
      title={all}
      className="inline-flex items-center rounded-full border border-line-strong px-2 py-0.5 font-mono text-[11px] text-fg-muted"
    >
      render-svg {lead.version}
    </span>
  );
}

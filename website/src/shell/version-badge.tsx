import { PACKAGES } from '@/site/packages';
import { installedVersion } from '@/site/versions';

/** The versions this site was built against, so a visitor can tell how far behind npm it is. */
export function VersionBadge() {
  const all = PACKAGES.map((one) => `@pptx-studio/${one.name} ${installedVersion(one.name)}`).join(
    '\n',
  );
  return (
    <span
      title={all}
      className="inline-flex items-center rounded-full border border-line-strong px-2 py-0.5 font-mono text-[11px] text-fg-muted"
    >
      render-svg {installedVersion('render-svg')}
    </span>
  );
}

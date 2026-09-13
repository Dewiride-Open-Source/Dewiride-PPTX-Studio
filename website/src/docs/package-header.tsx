import Link from 'next/link';

import { Badge } from '@/design/badge';
import { RUNTIME_LABELS, packageFacts, type PackageName } from '@/site/packages';
import { installedVersion } from '@/site/versions';

/** The facts every package page opens with: version, where it runs, the install line, the demo. */
export function PackageHeader({ pkg }: { pkg: PackageName }) {
  const facts = packageFacts(pkg);
  const install = ['opc', 'xml', 'geometry', 'paint', 'text', 'cli'].includes(pkg)
    ? `npm install @pptx-studio/${pkg}`
    : `npm install @pptx-studio/${pkg} ${facts.dependsOn.map((one) => `@pptx-studio/${one}`).join(' ')}`.trim();
  return (
    <div className="not-prose mb-6 flex flex-col gap-3 rounded-panel border border-line bg-surface p-4 shadow-panel">
      <div className="flex flex-wrap items-center gap-2">
        <code className="font-mono text-[13px] font-medium text-fg">@pptx-studio/{pkg}</code>
        <Badge>{installedVersion(pkg)}</Badge>
        <span className="text-[12px] text-fg-muted">runs in</span>
        {facts.runsIn.map((runtime) => (
          <Badge key={runtime} tone="good">
            {RUNTIME_LABELS[runtime]}
          </Badge>
        ))}
        <span className="ml-auto flex gap-3 text-[12px]">
          <Link href={`/demos/${pkg}`} className="text-accent hover:underline">
            Live demo →
          </Link>
          <a
            href={`https://www.npmjs.com/package/@pptx-studio/${pkg}`}
            target="_blank"
            rel="noreferrer"
            className="text-accent hover:underline"
          >
            npm →
          </a>
        </span>
      </div>
      <p className="text-[13px] text-fg-muted">{facts.blurb}</p>
      <pre className="overflow-x-auto rounded-control bg-sunken px-3 py-2 font-mono text-[12px] text-fg">
        {install}
      </pre>
    </div>
  );
}

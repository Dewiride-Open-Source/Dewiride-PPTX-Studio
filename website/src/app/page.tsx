import Link from 'next/link';

import { Panel } from '@/shell/panel';
import { Snippet } from '@/shell/code';
import { TOOLS } from '@/shell/tools';

export default function Overview() {
  return (
    <div className="mx-auto flex max-w-5xl flex-col gap-6">
      <header className="pt-2">
        <h1 className="text-2xl font-semibold text-ink-100">
          Eleven packages, eleven tools, nothing linked
        </h1>
        <p className="mt-2 max-w-2xl text-sm leading-relaxed text-ink-300">
          Every <code className="font-mono text-[12px] text-ink-200">@pptx-studio/*</code> import on
          this site was installed from the public npm registry. This directory sits outside the
          project&apos;s pnpm workspace on purpose, so nothing here resolves through a workspace
          link. If a published tarball were broken, this app would not build.
        </p>
      </header>

      <Panel
        title="Install"
        hint="Ten of the eleven run in the tab; the CLI is the one Node package."
      >
        <div className="p-4">
          <Snippet
            code={
              'npm install @pptx-studio/opc @pptx-studio/xml @pptx-studio/model \\n  @pptx-studio/render-svg @pptx-studio/census @pptx-studio/validate \\n  @pptx-studio/writer @pptx-studio/geometry @pptx-studio/paint \\n  @pptx-studio/text @pptx-studio/cli'
            }
          />
        </div>
      </Panel>

      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
        {TOOLS.map((tool) => (
          <Link
            key={tool.href}
            href={tool.href}
            className="group rounded-lg border border-ink-700 bg-ink-850 p-4 transition-colors hover:border-chrome"
          >
            <div className="flex items-baseline justify-between gap-2">
              <h2 className="text-sm font-semibold text-ink-100">{tool.name}</h2>
              {tool.server === true ? (
                <span className="text-[10px] text-widened">server</span>
              ) : null}
            </div>
            <p className="mt-1 font-mono text-[11px] text-chrome">@pptx-studio/{tool.package}</p>
            <p className="mt-2 text-[13px] leading-relaxed text-ink-300">{tool.blurb}</p>
          </Link>
        ))}
      </div>

      <Panel title="What this is not">
        <div className="space-y-2 p-4 text-[13px] leading-relaxed text-ink-300">
          <p>
            Not a PowerPoint clone. The editing on the Slides page is three XML edits with their own
            inverses, not the command bus that Phase 5 of the plan builds - there is no snapping, no
            resize solver and no multi-select yet, and this app does not pretend otherwise.
          </p>
          <p>
            Charts, SmartArt, tables and OLE objects are carried through an export byte-for-byte but
            are not drawn yet; they are later phases. What is drawn is geometry, fills, gradients,
            patterns, strokes, effects, pictures, nested rotated groups and text.
          </p>
        </div>
      </Panel>
    </div>
  );
}

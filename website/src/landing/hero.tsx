import { LinkButton } from '@/design/button';
import { VersionBadge } from '@/shell/version-badge';

export function Hero() {
  return (
    <section className="mx-auto max-w-6xl px-4 pt-16 pb-10 sm:px-6 sm:pt-24">
      <p className="mb-4 flex flex-wrap items-center gap-2 text-[12px] text-fg-muted">
        <span className="rounded-full border border-warn/40 bg-warn/10 px-2 py-0.5 font-mono text-[11px] text-warn">
          pre-alpha
        </span>
        <span>Apache-2.0</span>
        <span aria-hidden="true">·</span>
        <span>
          <code className="font-mono">@pptx-studio/*</code> on npm
        </span>
        <span aria-hidden="true">·</span>
        <span className="inline-flex items-center gap-1">
          built against <VersionBadge />
        </span>
      </p>
      <h1 className="max-w-4xl text-4xl leading-[1.1] font-semibold tracking-[-0.02em] text-balance sm:text-5xl">
        Open a <code className="font-mono font-medium text-accent">.pptx</code> in the browser, draw
        it the way PowerPoint does, and write it back without losing anything you did not touch.
      </h1>
      <p className="mt-6 max-w-3xl text-lg leading-relaxed text-fg-muted">
        PPTX Studio is twelve small TypeScript packages for reading, rendering and writing
        PowerPoint files. They treat the file itself as the document: slides keep inheriting from
        their layouts and masters instead of being flattened, every rule for colour, shapes and text
        was measured against real PowerPoint, and on export the parts you did not edit are copied
        out byte for byte. It all runs in a browser tab or a Web Worker with no server, and nothing
        you open is uploaded anywhere. One package, the CLI, runs the same renderer in Node.
      </p>
      <div className="mt-8 flex flex-wrap gap-3">
        <LinkButton href="/playground" variant="primary">
          Try it with a deck
        </LinkButton>
        <LinkButton href="/docs" variant="secondary">
          Read the docs
        </LinkButton>
      </div>
    </section>
  );
}

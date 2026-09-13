import Link from 'next/link';

import { DEFAULT_SAMPLE, sampleTitle } from '@/deck/samples';
import { publicUrl } from '@/site/base-path';

const DECK_ID = DEFAULT_SAMPLE.file.replace(/\.ppt[xm]$/, '');

/** The first slide of the default deck, as `@pptx-studio/cli` drew it when this site was built. */
export function Showcase() {
  return (
    <section aria-labelledby="showcase" className="mx-auto max-w-6xl px-4 pb-6 sm:px-6">
      <h2 id="showcase" className="sr-only">
        A rendered slide
      </h2>
      <figure className="rounded-frame border border-line bg-desk p-3 shadow-panel sm:p-6">
        <div className="stage-surface mx-auto aspect-video w-full max-w-4xl overflow-hidden rounded-[2px]">
          <img
            src={publicUrl(`/rendered/${DECK_ID}/slide-001.svg`)}
            alt={`The first slide of ${sampleTitle(DEFAULT_SAMPLE)}, rendered to SVG`}
            width={1280}
            height={720}
            className="h-full w-full"
          />
        </div>
        <figcaption className="mt-3 flex flex-wrap items-center justify-between gap-2 text-[12px] text-fg-muted">
          <span>
            <span className="text-fg">{sampleTitle(DEFAULT_SAMPLE)}</span> · {DEFAULT_SAMPLE.slides}{' '}
            slides · {(DEFAULT_SAMPLE.bytes / 1024).toFixed(0)} kB · drawn by{' '}
            <code className="font-mono">@pptx-studio/cli</code> when this site was built
          </span>
          <Link href="/demos/render-dom" className="text-accent hover:underline">
            Open it live, or drop your own →
          </Link>
        </figcaption>
      </figure>
    </section>
  );
}

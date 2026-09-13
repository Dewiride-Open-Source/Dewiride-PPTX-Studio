import { HomeLayout } from 'fumadocs-ui/layouts/home';
import Link from 'next/link';

import { baseOptions } from '@/site/navigation';

export default function Landing() {
  return (
    <HomeLayout {...baseOptions()}>
      <main className="mx-auto flex max-w-3xl flex-col gap-6 px-4 py-16">
        <h1 className="text-3xl font-semibold">
          Open a .pptx in the browser, draw it the way PowerPoint does, and write it back without
          losing anything you did not touch.
        </h1>
        <p className="text-fd-muted-foreground text-lg">
          PPTX Studio is twelve TypeScript packages for reading, rendering and writing PowerPoint
          files, entirely in a browser tab or a Web Worker.
        </p>
        <div className="flex gap-3">
          <Link
            href="/docs"
            className="rounded-md bg-fd-primary px-4 py-2 text-fd-primary-foreground"
          >
            Read the docs
          </Link>
          <Link href="/demos" className="rounded-md border border-fd-border px-4 py-2">
            Try the demos
          </Link>
        </div>
      </main>
    </HomeLayout>
  );
}

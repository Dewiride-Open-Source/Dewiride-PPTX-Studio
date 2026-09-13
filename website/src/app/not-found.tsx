import { HomeLayout } from 'fumadocs-ui/layouts/home';
import Link from 'next/link';

import { baseOptions } from '@/site/navigation';

export default function NotFound() {
  return (
    <HomeLayout {...baseOptions()}>
      <main className="mx-auto flex max-w-3xl flex-col gap-4 px-4 py-16">
        <h1 className="text-2xl font-semibold">There is no page here</h1>
        <p className="text-fd-muted-foreground">
          The address may have changed. Everything on this site is reachable from these four places.
        </p>
        <ul className="flex flex-wrap gap-3">
          <li>
            <Link href="/" className="underline">
              Home
            </Link>
          </li>
          <li>
            <Link href="/docs" className="underline">
              Docs
            </Link>
          </li>
          <li>
            <Link href="/demos" className="underline">
              Demos
            </Link>
          </li>
        </ul>
      </main>
    </HomeLayout>
  );
}

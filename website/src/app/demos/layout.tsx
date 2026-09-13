import { HomeLayout } from 'fumadocs-ui/layouts/home';
import type { ReactNode } from 'react';

import { DeckPicker } from '@/deck/picker';
import { DemosNav } from '@/demos/nav';
import { Footer } from '@/shell/footer';
import { baseOptions } from '@/site/navigation';

export default function DemosLayout({ children }: { children: ReactNode }) {
  return (
    <HomeLayout {...baseOptions()}>
      <div className="mx-auto w-full max-w-[1440px] px-4 py-4 sm:px-6">
        <DeckPicker />
        <div className="mt-4 flex gap-6">
          <aside className="sticky top-20 hidden w-48 shrink-0 self-start lg:block">
            <DemosNav />
          </aside>
          <main id="main" className="min-w-0 flex-1">
            {children}
          </main>
        </div>
      </div>
      <Footer />
    </HomeLayout>
  );
}

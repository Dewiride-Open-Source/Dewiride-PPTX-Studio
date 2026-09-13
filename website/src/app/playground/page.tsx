import { HomeLayout } from 'fumadocs-ui/layouts/home';
import type { Metadata } from 'next';

import { DeckPicker } from '@/deck/picker';
import { Bench } from '@/playground/bench';
import { Footer } from '@/shell/footer';
import { baseOptions } from '@/site/navigation';

export const metadata: Metadata = {
  title: 'Playground',
  description: 'Open your own .pptx and exercise every package on it, entirely in the browser.',
};

export default function PlaygroundPage() {
  return (
    <HomeLayout {...baseOptions()}>
      <div className="mx-auto w-full max-w-[1440px] px-4 py-4 sm:px-6">
        <header className="mb-4">
          <h1 className="text-xl font-semibold tracking-tight">Playground</h1>
          <p className="mt-1 text-[13px] text-fg-muted">
            One deck, every package. Drop your own - it never leaves your browser.
          </p>
        </header>
        <DeckPicker />
        <main id="main" className="mt-4">
          <Bench />
        </main>
      </div>
      <Footer />
    </HomeLayout>
  );
}

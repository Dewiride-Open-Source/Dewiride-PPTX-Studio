import type { ReactNode } from 'react';

import { DeckPicker } from '@/deck/picker';
import { Nav } from '@/shell/nav';

export default function DemosLayout({ children }: { children: ReactNode }) {
  return (
    <div className="flex min-h-screen bg-ink-900 text-ink-100 [color-scheme:dark]">
      <aside className="sticky top-0 hidden h-screen w-56 shrink-0 overflow-y-auto border-r border-ink-700 bg-ink-850 lg:block">
        <div className="border-b border-ink-700 px-4 py-4">
          <p className="text-sm font-semibold text-ink-100">PPTX Studio</p>
          <p className="font-mono text-[11px] text-ink-500">demos</p>
        </div>
        <Nav />
      </aside>

      <div className="flex min-w-0 flex-1 flex-col">
        <header className="sticky top-0 z-10 border-b border-ink-700 bg-ink-900/90 px-5 py-3 backdrop-blur">
          <DeckPicker />
        </header>
        <main className="min-w-0 flex-1 p-5">{children}</main>
      </div>
    </div>
  );
}

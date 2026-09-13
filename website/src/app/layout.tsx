import type { Metadata } from 'next';
import { IBM_Plex_Mono, IBM_Plex_Sans } from 'next/font/google';

import { DeckProvider } from '@/deck/provider';
import { DeckPicker } from '@/deck/picker';
import { Nav } from '@/shell/nav';

import './globals.css';

const sans = IBM_Plex_Sans({
  subsets: ['latin'],
  weight: ['400', '500', '600'],
  variable: '--font-plex-sans',
});

const mono = IBM_Plex_Mono({
  subsets: ['latin'],
  weight: ['400', '500'],
  variable: '--font-plex-mono',
});

export const metadata: Metadata = {
  title: 'PPTX Studio - package tour',
  description: 'A Next.js app built on the published @pptx-studio packages, one tool per package.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className={`${sans.variable} ${mono.variable}`}>
      <body className="font-sans antialiased">
        <DeckProvider>
          <div className="flex min-h-screen">
            <aside className="sticky top-0 hidden h-screen w-56 shrink-0 overflow-y-auto border-r border-ink-700 bg-ink-850 lg:block">
              <div className="border-b border-ink-700 px-4 py-4">
                <p className="text-sm font-semibold text-ink-100">PPTX Studio</p>
                <p className="font-mono text-[11px] text-ink-500">@pptx-studio/* 0.1.0</p>
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
        </DeckProvider>
      </body>
    </html>
  );
}

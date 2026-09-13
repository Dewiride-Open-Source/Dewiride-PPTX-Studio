import { RootProvider } from 'fumadocs-ui/provider/next';
import type { Metadata } from 'next';
import type { ReactNode } from 'react';

import { DeckProvider } from '@/deck/provider';
import { StaticSearchDialog } from '@/docs/search';

import './globals.css';

export const metadata: Metadata = {
  title: { default: 'PPTX Studio', template: '%s · PPTX Studio' },
  description:
    'Twelve TypeScript packages that open, draw and write back .pptx files in the browser, keeping every part you did not touch.',
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body className="flex min-h-screen flex-col font-sans antialiased">
        <RootProvider search={{ SearchDialog: StaticSearchDialog }}>
          <DeckProvider>{children}</DeckProvider>
        </RootProvider>
      </body>
    </html>
  );
}

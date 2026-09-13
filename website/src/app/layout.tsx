import { RootProvider } from 'fumadocs-ui/provider/next';
import type { Metadata } from 'next';
import type { ReactNode } from 'react';

import { DeckProvider } from '@/deck/provider';
import { LazySearchDialog } from '@/docs/search-lazy';
import { SITE_URL } from '@/site/navigation';

import './globals.css';

export const metadata: Metadata = {
  metadataBase: new URL(SITE_URL),
  title: {
    default: 'PPTX Studio - PowerPoint files in the browser, byte for byte',
    template: '%s · PPTX Studio',
  },
  openGraph: {
    type: 'website',
    siteName: 'PPTX Studio',
    images: [
      {
        url: '/og.png',
        width: 1200,
        height: 630,
        alt: 'PPTX Studio - PowerPoint files in the browser, byte for byte',
      },
    ],
  },
  twitter: { card: 'summary_large_image', images: ['/og.png'] },
  description:
    'Twelve TypeScript packages that open, draw and write back .pptx files in the browser, keeping every part you did not touch.',
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body className="flex min-h-screen flex-col font-sans antialiased">
        <RootProvider search={{ SearchDialog: LazySearchDialog }}>
          <DeckProvider>{children}</DeckProvider>
        </RootProvider>
      </body>
    </html>
  );
}

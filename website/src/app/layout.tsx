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

const ANALYTICS = 'https://analytics.dewiride.com';
const ANALYTICS_SITE = '01a0eb67-0b50-77a3-b07c-1acb8c3256f6';
/** Page views are counted on the published site only; `next dev` is never counted. ADR 0062. */
const COUNTED = process.env.NODE_ENV === 'production';

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      {COUNTED && (
        <head>
          <script defer src={`${ANALYTICS}/dw.js`} data-site={ANALYTICS_SITE} />
        </head>
      )}
      <body className="flex min-h-screen flex-col font-sans antialiased">
        {COUNTED && (
          <noscript>
            <img
              src={`${ANALYTICS}/collect/pixel.gif?site=${ANALYTICS_SITE}`}
              referrerPolicy="no-referrer-when-downgrade"
              alt=""
              width={1}
              height={1}
              style={{ position: 'absolute' }}
            />
          </noscript>
        )}
        <RootProvider search={{ SearchDialog: LazySearchDialog }}>
          <DeckProvider>{children}</DeckProvider>
        </RootProvider>
      </body>
    </html>
  );
}

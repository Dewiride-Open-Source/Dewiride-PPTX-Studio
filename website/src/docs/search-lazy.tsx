'use client';

import dynamic from 'next/dynamic';

/** The search dialog and its client, loaded on first open rather than with every page. */
export const LazySearchDialog = dynamic(
  () => import('./search').then((module) => ({ default: module.StaticSearchDialog })),
  { ssr: false },
);

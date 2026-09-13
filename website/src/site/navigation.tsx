import type { BaseLayoutProps } from 'fumadocs-ui/layouts/shared';

import { Wordmark } from '@/shell/brand';

export const REPOSITORY = 'https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio';
export const NPM_ORG = 'https://www.npmjs.com/org/pptx-studio';
export const SITE_URL = 'https://dewiride-open-source.github.io/Dewiride-PPTX-Studio';

/** The chrome every layout shares: the mark, the top links, the repository. */
export function baseOptions(): BaseLayoutProps {
  return {
    nav: { title: <Wordmark />, url: '/' },
    githubUrl: REPOSITORY,
    links: [
      { text: 'Docs', url: '/docs', active: 'nested-url' },
      { text: 'Demos', url: '/demos', active: 'nested-url' },
      { text: 'Playground', url: '/playground', active: 'nested-url' },
    ],
  };
}

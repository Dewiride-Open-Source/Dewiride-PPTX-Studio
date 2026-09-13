import type { BaseLayoutProps } from 'fumadocs-ui/layouts/shared';

export const REPOSITORY = 'https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio';
export const NPM_ORG = 'https://www.npmjs.com/org/pptx-studio';

/** The chrome every Fumadocs layout shares: the mark, the top links, the repository. */
export function baseOptions(): BaseLayoutProps {
  return {
    nav: { title: 'PPTX Studio' },
    githubUrl: REPOSITORY,
    links: [
      { text: 'Docs', url: '/docs', active: 'nested-url' },
      { text: 'Demos', url: '/demos', active: 'nested-url' },
    ],
  };
}

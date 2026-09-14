import { defineConfig } from 'fumadocs-mdx/config';

import { CODE_THEMES } from './src/site/code-themes';

export default defineConfig({
  mdxOptions: { rehypeCodeOptions: { themes: CODE_THEMES, defaultColor: false } },
});

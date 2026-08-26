import type { ESLint } from 'eslint';
import noTopLevelAwait from './no-top-level-await.ts';

/**
 * Local ESLint plugin for architecture rules that have no upstream equivalent.
 * Registered as `pptx-studio/*` in eslint.config.mjs.
 */
const plugin: ESLint.Plugin = {
  meta: { name: 'eslint-plugin-pptx-studio', version: '0.0.0' },
  rules: {
    'no-top-level-await': noTopLevelAwait,
  },
};

export default plugin;

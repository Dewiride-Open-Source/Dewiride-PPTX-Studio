import { builtinModules } from 'node:module';
import js from '@eslint/js';
import globals from 'globals';
import prettier from 'eslint-config-prettier/flat';
import tseslint from 'typescript-eslint';
import pptxStudio from './tools/eslint-rules/index.ts';

/**
 * There is exactly one ESLint config in this repo, and it lives here.
 *
 * ESLint 10 resolves `eslint.config.*` starting from each linted file's own
 * directory rather than from the cwd, so a stray config dropped into a package
 * would silently take that package over. Don't add one.
 *
 * The three bans below are the reason this file exists at all. Everything else
 * is hygiene.
 */

/** Every Node builtin, bare and `node:`-prefixed. */
const NODE_BUILTINS = [...new Set(builtinModules.flatMap((m) => [m, `node:${m}`]))].sort();

const NODE_BAN_MESSAGE =
  'Core packages run in a browser tab and in a Web Worker. There is no Node here. ' +
  'If you need a Node-only capability, it belongs in @pptx-studio/cli or in tools/, ' +
  'behind an interface the core defines.';

const REACT_BAN_MESSAGE =
  'Only @pptx-studio/react and apps/ may reference React. The core is framework-free by ' +
  'contract, and that contract is what makes a Vue or Svelte binding possible later.';

const DEEP_IMPORT_MESSAGE =
  'Import the package root (@pptx-studio/<name>). Deep imports bypass the exports map, so ' +
  'they break the moment the internal file layout changes and they hide layering violations ' +
  'from tools/layering.';

/** Directories under packages/ that ship to the browser. Mirrors tools/layering/layers.ts. */
const BROWSER_PACKAGES = [
  'opc',
  'xml',
  'fonts-metric-compat',
  'geometry',
  'paint',
  'text',
  'fonts',
  'model',
  'render-svg',
  'render-dom',
  'validate',
  'writer',
  'editor-core',
  'text-editor',
  'hard-content',
  'react',
];

const browserSources = BROWSER_PACKAGES.map((p) => `packages/${p}/src/**/*.ts`);
const nonReactSources = BROWSER_PACKAGES.filter((p) => p !== 'react').map(
  (p) => `packages/${p}/src/**/*.ts`,
);

export default tseslint.config(
  {
    ignores: [
      '**/dist/**',
      '**/node_modules/**',
      '**/.turbo/**',
      '**/coverage/**',
      '**/*.gen.ts',
      'corpus/**',
    ],
  },

  js.configs.recommended,
  tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        projectService: {
          // Build config files sit outside the include globs of every
          // auto-discovered tsconfig, so the project service cannot type them.
          // They are typechecked separately by tsconfig.node.json - which is
          // deliberately not named tsconfig.json, so the service does not find
          // it and then report the same file as belonging to two projects.
          allowDefaultProject: ['vitest.config.ts', 'packages/*/tsdown.config.ts'],
        },
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },

  // ---------------------------------------------------------------------------
  // Repo-wide baseline.
  // ---------------------------------------------------------------------------
  {
    files: ['**/*.ts'],
    plugins: { 'pptx-studio': pptxStudio },
    rules: {
      'pptx-studio/no-top-level-await': 'error',

      // `import type` / `export type` everywhere, so the emitted JS has no
      // phantom imports and `verbatimModuleSyntax` stays honest.
      '@typescript-eslint/consistent-type-imports': [
        'error',
        { prefer: 'type-imports', fixStyle: 'inline-type-imports' },
      ],
      '@typescript-eslint/no-import-type-side-effects': 'error',

      // Unused code is a smell in a codebase whose whole thesis is "never emit
      // markup you did not read". Leading underscore is the escape hatch.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          destructuredArrayIgnorePattern: '^_',
        },
      ],

      'no-console': ['error', { allow: ['warn', 'error'] }],
      eqeqeq: ['error', 'always', { null: 'ignore' }],
      'no-var': 'error',
      'prefer-const': 'error',
    },
  },

  // ---------------------------------------------------------------------------
  // Ban 1 - no Node in the browser packages.
  // ---------------------------------------------------------------------------
  {
    files: browserSources,
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: NODE_BUILTINS.map((name) => ({ name, message: NODE_BAN_MESSAGE })),
          patterns: [
            { group: ['node:*'], message: NODE_BAN_MESSAGE },
            {
              group: ['@pptx-studio/*/src/*', '@pptx-studio/*/dist/*'],
              message: DEEP_IMPORT_MESSAGE,
            },
          ],
        },
      ],
      // The globals that survive bundling even without an import.
      'no-restricted-globals': [
        'error',
        { name: 'process', message: NODE_BAN_MESSAGE },
        { name: 'Buffer', message: NODE_BAN_MESSAGE + ' Use Uint8Array.' },
        { name: '__dirname', message: NODE_BAN_MESSAGE },
        { name: '__filename', message: NODE_BAN_MESSAGE },
        { name: 'require', message: 'This repo is ESM-only.' },
      ],
    },
  },

  // ---------------------------------------------------------------------------
  // Ban 2 - React lives in exactly one package.
  // ---------------------------------------------------------------------------
  {
    files: nonReactSources,
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: [
            ...NODE_BUILTINS.map((name) => ({ name, message: NODE_BAN_MESSAGE })),
            { name: 'react', message: REACT_BAN_MESSAGE },
            { name: 'react-dom', message: REACT_BAN_MESSAGE },
          ],
          patterns: [
            { group: ['node:*'], message: NODE_BAN_MESSAGE },
            { group: ['react/*', 'react-dom/*'], message: REACT_BAN_MESSAGE },
            { group: ['@pptx-studio/react', '@pptx-studio/react/*'], message: REACT_BAN_MESSAGE },
            {
              group: ['@pptx-studio/*/src/*', '@pptx-studio/*/dist/*'],
              message: DEEP_IMPORT_MESSAGE,
            },
          ],
        },
      ],
    },
  },

  // ---------------------------------------------------------------------------
  // Node-side code: tools/, cli/, and the root config files.
  // ---------------------------------------------------------------------------
  {
    files: ['tools/**/*.ts', 'packages/cli/src/**/*.ts', '*.ts', '*.mjs'],
    languageOptions: { globals: globals.nodeBuiltin },
    rules: {
      // tools/ is executed directly by Node as ESM. Top-level await is fine
      // there; nothing publishes it and nothing `require()`s it.
      'pptx-studio/no-top-level-await': 'off',
      'no-console': 'off',
    },
  },

  // Config files are plain JS: no type information, so type-aware rules must be
  // switched off or every one of them errors on an untyped file.
  {
    files: ['**/*.mjs', '**/*.js'],
    extends: [tseslint.configs.disableTypeChecked],
  },

  // Test files may reach for `any` while pinning down parser edge cases.
  {
    files: ['**/*.test.ts', '**/*.spec.ts'],
    rules: {
      '@typescript-eslint/no-non-null-assertion': 'off',
      '@typescript-eslint/no-unsafe-assignment': 'off',
      '@typescript-eslint/no-unsafe-member-access': 'off',
    },
  },

  // Must stay last: turns off every rule Prettier owns.
  prettier,
);

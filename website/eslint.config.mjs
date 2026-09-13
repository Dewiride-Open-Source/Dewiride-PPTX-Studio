import js from '@eslint/js';
import next from '@next/eslint-plugin-next';
import prettier from 'eslint-config-prettier/flat';
import reactHooks from 'eslint-plugin-react-hooks';
import globals from 'globals';
import tseslint from 'typescript-eslint';

// The repository's config ignores this directory: its packages come from the
// registry, which the root project service cannot resolve. This one can.
export default tseslint.config(
  {
    ignores: [
      '**/node_modules/**',
      '.next/**',
      'out/**',
      'next-env.d.ts',
      'public/rendered/**',
      'src/reference/generated/**',
    ],
  },

  js.configs.recommended,
  tseslint.configs.recommendedTypeChecked,
  next.configs['core-web-vitals'],
  reactHooks.configs.flat.recommended,
  {
    languageOptions: {
      parserOptions: {
        projectService: { allowDefaultProject: ['*.mjs', 'deploy/*.mjs', 'prerender/*.mjs'] },
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },

  {
    files: ['**/*.ts', '**/*.tsx'],
    rules: {
      '@typescript-eslint/consistent-type-imports': [
        'error',
        { prefer: 'type-imports', fixStyle: 'inline-type-imports' },
      ],
      '@typescript-eslint/no-import-type-side-effects': 'error',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          destructuredArrayIgnorePattern: '^_',
        },
      ],
      '@typescript-eslint/restrict-template-expressions': ['error', { allowNumber: false }],
      'no-console': ['error', { allow: ['warn', 'error'] }],
      // A static export has no image optimiser, and every image here is an SVG.
      '@next/next/no-img-element': 'off',
      eqeqeq: ['error', 'always', { null: 'ignore' }],
      'no-var': 'error',
      'prefer-const': 'error',
    },
  },

  // The docs' code: a line that names a value is the point, and printing it is too.
  {
    files: ['snippets/**/*.ts'],
    rules: { '@typescript-eslint/no-unused-expressions': 'off', 'no-console': 'off' },
  },

  // Build-time scripts: Node, plain JavaScript, and they print.
  {
    files: ['*.mjs', 'deploy/*.mjs', 'prerender/*.mjs'],
    languageOptions: { globals: globals.node },
    rules: {
      'no-console': 'off',
      ...Object.fromEntries(
        Object.keys(tseslint.configs.recommendedTypeChecked.at(-1)?.rules ?? {}).map((rule) => [
          rule,
          'off',
        ]),
      ),
    },
  },

  prettier,
);

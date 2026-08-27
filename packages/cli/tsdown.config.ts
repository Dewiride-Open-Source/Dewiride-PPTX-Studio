import { defineConfig } from 'tsdown';

export default defineConfig({
  // Two entries: the library surface, and the bin. `cli.ts` carries the
  // shebang and calls straight into `main`, so `npx @pptx-studio/cli` and
  // `import { main } from '@pptx-studio/cli'` run identical code.
  entry: ['src/index.ts', 'src/cli.ts'],
  format: 'esm',
  platform: 'node',
  target: 'node24',
  // `.js`, not the `.mjs` tsdown picks for a Node target. The package is
  // `"type": "module"`, so `.js` already means ESM here, and matching the other
  // packages means one spelling of a dist path across the repository - `bin`
  // and `exports` entries that disagree with what the build emitted are what
  // publint exists to catch.
  outExtensions: () => ({ js: '.js', dts: '.d.ts' }),
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

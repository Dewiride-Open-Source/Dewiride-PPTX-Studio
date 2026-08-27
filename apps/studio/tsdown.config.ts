import { defineConfig } from 'tsdown';

export default defineConfig({
  // Two entries, because a module worker is a second entry point by definition:
  // `new Worker(new URL('./worker.js', import.meta.url))` in `main.js` resolves
  // against the emitted file's own URL, so `worker.js` has to sit beside it.
  entry: ['src/main.ts', 'src/worker.ts'],
  format: 'esm',
  platform: 'browser',
  target: 'es2023',
  // An app is not a library: nothing imports it, so everything it uses has to
  // be in the bundle. Without this the emitted files would `import` bare
  // specifiers that a browser cannot resolve, and a Worker cannot be given an
  // import map to fix it with - import maps are document-scoped.
  deps: { alwaysBundle: [/^@pptx-studio\//, 'fflate'] },
  dts: false,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts'],
  format: 'esm',
  platform: 'browser',
  target: 'es2023',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

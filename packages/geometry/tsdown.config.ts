import { defineConfig } from 'tsdown';

export default defineConfig({
  // Six extra entries so a consumer that knows which shapes it needs can import
  // one bucket and leave the other 158 presets out of its bundle. The shared
  // decoder is split into a chunk both entries reference.
  entry: ['src/index.ts', 'src/presets/*.gen.ts'],
  format: 'esm',
  platform: 'browser',
  target: 'es2023',
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
});

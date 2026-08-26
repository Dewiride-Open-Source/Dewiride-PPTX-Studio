import { readdirSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { playwright } from '@vitest/browser-playwright';
import { defineConfig } from 'vitest/config';

const root = fileURLToPath(new URL('.', import.meta.url));

/**
 * Cross-package resolution reads source, not dist, so `pnpm test` never depends
 * on a build. Kept in step with the `paths` block in tsconfig.json - if these
 * two disagree, tests and typecheck are looking at different code.
 */
function workspaceAliases(): Record<string, string> {
  const packagesDir = join(root, 'packages');
  const alias: Record<string, string> = {};
  for (const name of readdirSync(packagesDir)) {
    alias[`@pptx-studio/${name}`] = join(packagesDir, name, 'src/index.ts');
  }
  return alias;
}

export default defineConfig({
  resolve: { alias: workspaceAliases() },
  // Scan every package source up front. Left to itself Vite discovers a
  // runtime dependency on first import inside a browser test, re-optimises and
  // reloads mid-run, which Vitest reports as a source of flaky and duplicated
  // tests. `entries` rather than `include`, because these dependencies resolve
  // from each package's own node_modules and not from the workspace root.
  optimizeDeps: { entries: ['packages/*/src/**/*.ts'] },
  test: {
    projects: [
      {
        // The core packages ship to a browser. Under jsdom, `node:fs` resolves
        // and a Node API leak stays invisible until a user opens a tab; under
        // real Chromium it throws. That is the whole reason for browser mode.
        extends: true,
        test: {
          name: 'core',
          include: ['packages/*/src/**/*.test.ts'],
          browser: {
            provider: playwright(),
            enabled: true,
            headless: true,
            instances: [{ browser: 'chromium' }],
          },
        },
      },
      {
        // tools/ is Node-side by definition: it reads the filesystem and the
        // workspace manifests.
        extends: true,
        test: {
          name: 'tools',
          environment: 'node',
          include: ['tools/**/*.test.ts'],
        },
      },
    ],
  },
});

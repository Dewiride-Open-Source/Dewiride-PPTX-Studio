/** The workspace root, found by walking up to `pnpm-workspace.yaml`. */

import { existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

function findRoot(from: string): string {
  let at = from;
  for (;;) {
    if (existsSync(join(at, 'pnpm-workspace.yaml'))) return at;
    const up = dirname(at);
    if (up === at) {
      throw new Error(`no pnpm-workspace.yaml at or above ${from}; this is not the workspace`);
    }
    at = up;
  }
}

export const REPO_ROOT = findRoot(fileURLToPath(new URL('.', import.meta.url)));

/** An absolute path to something in the repository, from repo-relative parts. */
export function repoPath(...parts: readonly string[]): string {
  return resolve(REPO_ROOT, ...parts);
}

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import type { PackageName } from './packages';

/** The version of a package as installed when this site was built; read on the server, never bundled. */
export function installedVersion(name: PackageName): string {
  const manifest = JSON.parse(
    readFileSync(join(process.cwd(), 'node_modules', '@pptx-studio', name, 'package.json'), 'utf8'),
  ) as { version: string };
  return manifest.version;
}

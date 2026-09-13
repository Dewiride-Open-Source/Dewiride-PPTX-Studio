import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { SiteError } from '@/site/errors';
import type { PackageName } from '@/site/packages';

export type SymbolKind =
  'function' | 'class' | 'interface' | 'type' | 'enum' | 'const' | 're-export';

export interface ReferenceSymbol {
  readonly name: string;
  readonly kind: SymbolKind;
  readonly isType: boolean;
  readonly signature: string;
  readonly doc: string;
}

export interface ReferenceGroup {
  readonly file: string;
  readonly symbols: readonly ReferenceSymbol[];
}

export interface PackageReference {
  readonly name: PackageName;
  readonly version: string;
  readonly description: string;
  readonly exported: number;
  readonly groups: readonly ReferenceGroup[];
  readonly errors: { readonly className: string; readonly codes: readonly string[] } | null;
}

const GENERATED = join(process.cwd(), 'src', 'reference', 'generated');

/** The reference `prerender/reference.mjs` wrote for one package, read while the site is built. */
export function readReference(name: PackageName): PackageReference {
  let text: string;
  try {
    text = readFileSync(join(GENERATED, `${name}.json`), 'utf8');
  } catch {
    throw new SiteError(
      'SITE_REFERENCE_MISSING',
      `no reference for @pptx-studio/${name}; run: node prerender/reference.mjs`,
    );
  }
  const parsed = JSON.parse(text) as PackageReference;
  if (parsed.name !== name || !Array.isArray(parsed.groups)) {
    throw new SiteError(
      'SITE_REFERENCE_MISSING',
      `the reference for ${name} is not what prerender/reference.mjs writes`,
    );
  }
  return parsed;
}

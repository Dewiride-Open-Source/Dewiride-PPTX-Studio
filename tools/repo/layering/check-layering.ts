#!/usr/bin/env node
/**
 * Enforce the one-way dependency graph at the *manifest* level.
 *
 * Import-level bans live in eslint.config.mjs and are the second net. This is
 * the first one, because manifests are what actually ship: a package.json that
 * declares `react` has already broken the framework-free contract whether or
 * not any file imports it yet.
 *
 * Run: `pnpm layering`  (Node >= 22.18 strips the types; no build step.)
 */
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

import { REPO_ROOT as ROOT } from '../root.ts';
import { parse as parseYaml } from 'yaml';
import { checkLayering, type Manifest, type Violation } from './check.ts';

function readJson(path: string): unknown {
  return JSON.parse(readFileSync(path, 'utf8'));
}

/** Directories directly under `parent` that contain a package.json. */
function collectManifests(parent: string): Manifest[] {
  const abs = join(ROOT, parent);
  let entries: string[];
  try {
    entries = readdirSync(abs);
  } catch {
    return []; // workspace root declares apps/ before apps/ exists
  }
  const out: Manifest[] = [];
  for (const entry of entries.sort()) {
    const manifestPath = join(abs, entry, 'package.json');
    try {
      if (!statSync(manifestPath).isFile()) continue;
    } catch {
      continue;
    }
    out.push({
      dir: relative(ROOT, join(abs, entry)).split('\\').join('/'),
      json: readJson(manifestPath) as Manifest['json'],
    });
  }
  return out;
}

function readCatalog(): Record<string, string> {
  const raw = readFileSync(join(ROOT, 'pnpm-workspace.yaml'), 'utf8');
  const parsed = parseYaml(raw) as { catalog?: Record<string, string> } | null;
  return parsed?.catalog ?? {};
}

function format(violations: readonly Violation[]): string {
  const byRule = new Map<string, Violation[]>();
  for (const v of violations) {
    const bucket = byRule.get(v.rule);
    if (bucket) bucket.push(v);
    else byRule.set(v.rule, [v]);
  }
  const lines: string[] = [];
  for (const [rule, items] of [...byRule].sort(([a], [b]) => a.localeCompare(b))) {
    lines.push('');
    lines.push('  ' + rule);
    for (const item of items) {
      lines.push('    ' + item.where + '/package.json');
      lines.push('      ' + item.message);
    }
  }
  return lines.join('\n');
}

const manifests = [...collectManifests('packages'), ...collectManifests('apps')];
const violations = checkLayering({ manifests, catalog: readCatalog() });

const checked = manifests.filter((m) => m.dir.startsWith('packages/')).length;

if (violations.length === 0) {
  console.log('layering: ' + checked + ' package(s) checked, no violations');
  process.exit(0);
}

console.error('layering: ' + violations.length + ' violation(s) across ' + checked + ' package(s)');
console.error(format(violations));
console.error('');
console.error(
  '  The layer table is tools/repo/layering/layers.ts. If a dependency genuinely belongs',
);
console.error('  where it is, move the package in the table and say why in the commit message.');
process.exit(1);

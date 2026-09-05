#!/usr/bin/env node
/**
 * Enforce the corpus's provenance rules.
 *
 * ```
 * pnpm corpus
 * ```
 *
 * `LEGAL.md` commits this project to a licence field on every fixture and to a
 * closed set of sources they may come from. This is the gate that makes that
 * true rather than aspirational, and sub-phase 1.1's verification is precisely
 * "CI fails on any entry lacking a license".
 *
 * Two design notes, both load-bearing:
 *
 * - **It walks the filesystem rather than asking git.** `readdirSync` sees
 *   untracked files, so a local `pnpm check` catches an unlicensed deck dropped
 *   into `corpus/` *before* it is committed. `git ls-files` would not, and it
 *   would also stop working in a source tarball with no `.git`. That is
 *   `C008-orphan`, the only rule that fires when somebody adds a file without
 *   touching a manifest - which is the failure this whole check exists for.
 * - **It runs before the build and needs nothing from it.** No workspace
 *   package is imported; the census's key list is committed as
 *   `census-keys.gen.ts`. So this is the second gate in `pnpm check`, after
 *   layering, and a contributor learns about a licence problem in a second
 *   rather than four minutes into a Chromium install.
 *
 * Exit status: 0 clean, 1 violations, 2 the check could not run. The third is
 * separate on purpose - a malformed manifest is a different incident from a
 * manifest that says something wrong.
 */

import { createHash } from 'node:crypto';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { CENSUS_FEATURE_KEYS } from './census-keys.gen.ts';
import { checkCorpus, humanBytes, type CorpusFile, type ManifestFile } from './check.ts';
import type { Violation } from './schema.ts';

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const CORPUS = join(ROOT, 'corpus');

/** Workspace-relative, forward slashes, so a message is the same on every OS. */
function rel(absolute: string): string {
  return relative(ROOT, absolute).split('\\').join('/');
}

interface Walked {
  readonly manifests: string[];
  readonly files: string[];
}

function walk(dir: string, out: Walked = { manifests: [], files: [] }): Walked {
  let entries: string[];
  try {
    entries = readdirSync(dir);
  } catch {
    return out;
  }
  for (const name of entries.sort()) {
    const absolute = join(dir, name);
    let stats;
    try {
      stats = statSync(absolute);
    } catch {
      continue;
    }
    if (stats.isDirectory()) {
      walk(absolute, out);
      continue;
    }
    if (!stats.isFile()) continue;
    if (name === 'manifest.json') out.manifests.push(absolute);
    else out.files.push(absolute);
  }
  return out;
}

/** Extensions `.gitattributes` marks `binary`, for `C016`. */
function readBinaryExtensions(): Set<string> {
  const out = new Set<string>();
  let text: string;
  try {
    text = readFileSync(join(ROOT, '.gitattributes'), 'utf8');
  } catch {
    return out;
  }
  for (const line of text.split(/\r?\n/)) {
    const match = /^\s*\*\.([A-Za-z0-9]+)\s+.*\bbinary\b/.exec(line);
    if (match?.[1] !== undefined) out.add(match[1].toLowerCase());
  }
  return out;
}

/** Every file under `tools/`, workspace-relative, for `C014`. */
function readToolPaths(): Set<string> {
  const out = new Set<string>();
  const recurse = (dir: string): void => {
    let entries: string[];
    try {
      entries = readdirSync(dir);
    } catch {
      return;
    }
    for (const name of entries) {
      const absolute = join(dir, name);
      try {
        if (statSync(absolute).isDirectory()) recurse(absolute);
        else out.add(rel(absolute));
      } catch {
        /* a file that vanished between readdir and stat is not our problem */
      }
    }
  };
  recurse(join(ROOT, 'tools'));
  return out;
}

function format(violations: readonly Violation[]): string {
  const byRule = new Map<string, Violation[]>();
  for (const violation of violations) {
    const bucket = byRule.get(violation.rule);
    if (bucket) bucket.push(violation);
    else byRule.set(violation.rule, [violation]);
  }
  const lines: string[] = [];
  for (const [rule, items] of [...byRule].sort(([a], [b]) => a.localeCompare(b))) {
    lines.push('');
    lines.push('  ' + rule);
    for (const item of items) {
      lines.push('    ' + item.where);
      for (const line of item.message.split('\n')) lines.push('      ' + line.trimEnd());
    }
  }
  return lines.join('\n');
}

// --------------------------------------------------------------------- run it

let walked: Walked;
try {
  statSync(CORPUS);
  walked = walk(CORPUS);
} catch {
  console.error('corpus: no corpus/ directory at ' + rel(CORPUS));
  process.exit(2);
}

const manifests: ManifestFile[] = [];
for (const absolute of walked.manifests) {
  let json: unknown;
  try {
    json = JSON.parse(readFileSync(absolute, 'utf8'));
  } catch (error) {
    console.error('corpus: ' + rel(absolute) + ' is not valid JSON');
    console.error('  ' + (error instanceof Error ? error.message : String(error)));
    process.exit(2);
  }
  manifests.push({
    path: rel(absolute),
    dir: rel(join(absolute, '..')),
    json,
  });
}

const files: CorpusFile[] = walked.files.map((absolute) => {
  const bytes = readFileSync(absolute);
  return {
    path: rel(absolute),
    bytes: bytes.byteLength,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  };
});

const violations = checkCorpus({
  manifests,
  files,
  binaryExtensions: readBinaryExtensions(),
  toolPaths: readToolPaths(),
  censusKeys: CENSUS_FEATURE_KEYS,
});

const entryCount = manifests.reduce((total, manifest) => {
  const entries = (manifest.json as { entries?: unknown[] } | null)?.entries;
  return total + (Array.isArray(entries) ? entries.length : 0);
}, 0);
const totalBytes = files.reduce((total, file) => total + file.bytes, 0);

if (violations.length === 0) {
  console.log(
    'corpus: ' +
      String(manifests.length) +
      ' manifest(s), ' +
      String(entryCount) +
      ' entrie(s), ' +
      String(files.length) +
      ' file(s), ' +
      humanBytes(totalBytes) +
      ', no violations',
  );
  process.exit(0);
}

const offending = new Set(violations.map((violation) => violation.where));
console.error(
  'corpus: ' +
    String(violations.length) +
    ' violation(s) across ' +
    String(offending.size) +
    ' place(s) in ' +
    String(manifests.length) +
    ' manifest(s)',
);
console.error(format(violations));
console.error('');
console.error('  The enums are tools/corpus/schema.ts. The policy they encode is LEGAL.md,');
console.error('  "Corpus licensing". An entry with no licence is not a to-do: establish the');
console.error('  provenance or remove the file. Do not guess a licence to make this pass.');
process.exit(1);

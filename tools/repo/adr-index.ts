/**
 * `docs/adr/README.md` - the ordered log across the per-phase directories.
 *
 * ```
 * node tools/repo/adr-index.ts            # write it
 * node tools/repo/adr-index.ts --check    # fail if it is out of date
 * ```
 */

import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { format, resolveConfig } from 'prettier';

import { repoPath } from './root.ts';

const ADR = repoPath('docs/adr');
const INDEX = join(ADR, 'README.md');

interface Record {
  readonly phase: string;
  readonly file: string;
  readonly number: string;
  readonly title: string;
}

/** `# 0031 — Bullets, fields and script runs` -> `Bullets, fields and script runs`. */
function titleOf(text: string, fallback: string): string {
  const heading = /^#\s+(.+)$/m.exec(text)?.[1];
  if (heading === undefined) return fallback;
  return heading.replace(/^(?:ADR\s+)?\d{4}\s*[-—]\s*/, '').trim();
}

function read(): Record[] {
  const records: Record[] = [];
  const phases = readdirSync(ADR, { withFileTypes: true })
    .filter((e) => e.isDirectory())
    .map((e) => e.name)
    .sort((a, b) => a.localeCompare(b));
  for (const phase of phases) {
    const files = readdirSync(join(ADR, phase))
      .filter((f) => /^\d{4}-.*\.md$/.test(f))
      .sort((a, b) => a.localeCompare(b));
    for (const file of files) {
      records.push({
        phase,
        file,
        number: file.slice(0, 4),
        title: titleOf(readFileSync(join(ADR, phase, file), 'utf8'), file),
      });
    }
  }
  return records;
}

/** `phase-2-geometry-and-paint` -> `Phase 2 — geometry and paint`. */
function phaseTitle(slug: string): string {
  const match = /^phase-(\d+)-(.+)$/.exec(slug);
  if (match === null) return slug;
  return `Phase ${match[1] ?? ''} — ${(match[2] ?? '').replace(/-/g, ' ')}`;
}

function render(records: readonly Record[]): string {
  const lines: string[] = [
    '# Architecture decision records',
    '',
    'One record per sub-phase: what was decided, what was measured to decide it, which alternatives',
    'the measurement ruled out, and what is still open. Records are grouped by phase so the directory',
    'stays navigable at ninety-six of them, and this index is the ordered log that grouping would',
    'otherwise cost.',
    '',
    'Numbers are permanent. A record is never renumbered and never edited to say something it did not',
    'say — a decision that turns out to be wrong gets a new record that supersedes it, and the',
    'measurement that refuted the old one goes in the new one.',
    '',
    '**This file is generated** by `tools/repo/adr-index.ts`; `pnpm references` fails if it is out of',
    'date.',
  ];
  let current = '';
  for (const record of records) {
    if (record.phase !== current) {
      current = record.phase;
      lines.push('', `## ${phaseTitle(record.phase)}`, '');
      lines.push('| # | decision |', '| --- | --- |');
    }
    lines.push(`| [${record.number}](${record.phase}/${record.file}) | ${record.title} |`);
  }
  lines.push('');
  return lines.join('\n');
}

/** Formatted through Prettier so that `pnpm format:check` and this agree. */
const config = await resolveConfig(INDEX);
const wanted = await format(render(read()), { ...config, filepath: INDEX });

if (process.argv.includes('--check')) {
  let found: string;
  try {
    found = readFileSync(INDEX, 'utf8');
  } catch {
    found = '';
  }
  if (found !== wanted) {
    console.error('docs/adr/README.md is out of date - run: node tools/repo/adr-index.ts');
    process.exitCode = 1;
  }
} else {
  writeFileSync(INDEX, wanted);
  console.log(`wrote ${INDEX}`);
}

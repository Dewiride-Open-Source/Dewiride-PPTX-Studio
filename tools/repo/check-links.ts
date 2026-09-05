/**
 * Every reference to a file, in prose or in a comment, resolves.
 *
 * ```
 * node tools/repo/check-links.ts
 * ```
 */

import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { join, posix } from 'node:path';

import { REPO_ROOT as ROOT } from './root.ts';

/** `[text](target)`, minus URLs, anchors and mailto. */
const LINK = /\[[^\]]*\]\((?!https?:|#|mailto:|data:)([^)\s]+?)(?:#[^)]*)?\)/g;

/** A repo-relative path in prose: one of the five top-level directories, and a real extension. */
const PROSE_PATH =
  /(?<![\w./-])((?:packages|tools|apps|corpus|docs)\/[A-Za-z0-9_./-]*\.[A-Za-z0-9]{2,5})(?![\w-])/g;

/** Where a reference may name something generated rather than committed. */
const GENERATED = [/\/dist\//, /\/node_modules\//, /__screenshots__/, /\.turbo\//];

/**
 * Files whose paths are not descriptions of the repository as it stands: an ADR
 * is a historical record, and a test names synthetic paths on purpose.
 */
const NOT_DESCRIPTIVE = [/^docs\/adr\//, /\.test\.tsx?$/];

const SCANNED = /\.(md|ts|tsx|ps1)$/;

interface Problem {
  readonly file: string;
  readonly line: number;
  readonly target: string;
  readonly why: string;
}

function trackedFiles(): string[] {
  return execFileSync('git', ['ls-files'], { cwd: ROOT, encoding: 'utf8' })
    .split('\n')
    .filter((line) => line.length > 0);
}

function lineOf(text: string, index: number): number {
  let line = 1;
  for (let i = 0; i < index; i++) if (text[i] === '\n') line++;
  return line;
}

function check(file: string, text: string): Problem[] {
  const problems: Problem[] = [];
  const here = posix.dirname(file);

  if (file.endsWith('.md')) {
    for (const match of text.matchAll(LINK)) {
      const target = match[1];
      if (target === undefined) continue;
      const repoRel = posix.normalize(posix.join(here, decodeURIComponent(target)));
      if (repoRel.startsWith('..')) {
        problems.push({
          file,
          line: lineOf(text, match.index),
          target,
          why: 'link escapes the repository',
        });
        continue;
      }
      if (GENERATED.some((re) => re.test(repoRel))) continue;
      if (!existsSync(join(ROOT, repoRel))) {
        problems.push({ file, line: lineOf(text, match.index), target, why: 'no such file' });
      }
    }
  }

  if (!NOT_DESCRIPTIVE.some((re) => re.test(file))) {
    for (const match of text.matchAll(PROSE_PATH)) {
      const target = match[1];
      if (target === undefined) continue;
      if (GENERATED.some((re) => re.test(target))) continue;
      if (!existsSync(join(ROOT, target))) {
        problems.push({
          file,
          line: lineOf(text, match.index),
          target,
          why: 'path named in prose does not exist',
        });
      }
    }
  }

  return problems;
}

const files = trackedFiles().filter((f) => SCANNED.test(f));
const problems: Problem[] = [];
for (const file of files) {
  problems.push(...check(file, await readFile(join(ROOT, file), 'utf8')));
}

if (problems.length > 0) {
  for (const p of problems) console.error(`${p.file}:${String(p.line)}  ${p.target} - ${p.why}`);
  console.error(
    `\n${String(problems.length)} broken reference(s) across ${String(files.length)} file(s)`,
  );
  process.exitCode = 1;
} else {
  console.log(`docs: every reference in ${String(files.length)} file(s) resolves`);
}

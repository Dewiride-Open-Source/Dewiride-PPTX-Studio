/**
 * Prove the tarballs this commit would publish, before publishing them.
 *
 * ```
 * pnpm candidate <work-dir>                 # pack, then consume
 * pnpm candidate <work-dir> --pack-only     # pack into <work-dir>/tarballs, with packed.json
 * pnpm candidate <work-dir> --packed        # consume what --pack-only recorded
 * ```
 *
 * Packs every publishable package, installs the tarballs into a scratch project
 * outside the workspace, refuses any resolution that did not come from one of
 * them, and runs the consumer smoke test against the result. ADR 0046, ADR 0059.
 */

import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';

import { repoPath } from '../../repo/root.ts';
import {
  packedFrom,
  provenanceFailures,
  readPackedRecord,
  scratchManifest,
  staleRanges,
  tamperedTarballs,
  type Lockfile,
  type Packed,
  type PackedRecord,
} from './consumer.ts';
import { toolCommand, type Tool } from './spawn.ts';

const SCOPE = '@pptx-studio/';

/** Where a registry lookup for our own scope has to fail loudly rather than resolve. */
const NOWHERE = 'http://127.0.0.1:1/';

/** What a local install or build leaves in `website/`; the scratch copy makes its own. */
const GENERATED =
  /node_modules|package-lock\.json|[\\/](?:\.next|out|public[\\/]rendered|src[\\/]reference[\\/]generated)(?=[\\/]|$)/;

const [work, mode] = process.argv.slice(2);
if (work === undefined || (mode !== undefined && mode !== '--pack-only' && mode !== '--packed')) {
  throw new Error('usage: candidate.ts <work-dir> [--pack-only | --packed]');
}
const workDir = resolve(work);
const tarballs = join(workDir, 'tarballs');
const RECORD = join(tarballs, 'packed.json');

function run(name: Tool, args: readonly string[], cwd: string): string {
  const host = { platform: process.platform, execPath: process.execPath, env: process.env };
  const command = toolCommand(name, args, { ...host, exists: existsSync });
  const result = spawnSync(command.file, command.args, {
    cwd,
    env: command.env,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'inherit'],
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.error !== undefined) throw result.error;
  if (result.status !== 0) {
    throw new Error(
      `${name} ${args.join(' ')} exited with ${String(result.status ?? result.signal)}`,
    );
  }
  return result.stdout;
}

/** `pnpm pack` rewrites `workspace:^` and `catalog:` exactly as publish does. */
function pack(into: string): readonly Packed[] {
  const raw = run(
    'pnpm',
    ['-r', '--filter', './packages/*', 'pack', '--pack-destination', into, '--json'],
    repoPath('.'),
  );
  const packed = packedFrom(raw);
  if (packed.length === 0) throw new Error(`pnpm pack reported no tarballs:\n${raw}`);
  return packed.map((entry) => ({ ...entry, filename: resolve(entry.filename) }));
}

function integrityOf(filename: string): string {
  return `sha512-${createHash('sha512').update(readFileSync(filename)).digest('base64')}`;
}

/** The tarballs `--pack-only` recorded, held to their digests before anything installs them. */
function recorded(): readonly Packed[] {
  const record = readPackedRecord(readFileSync(RECORD, 'utf8'));
  const tampered = tamperedTarballs(record, (file) => integrityOf(join(tarballs, file)));
  if (tampered.length > 0) {
    throw new Error(`tarballs changed since they were packed:\n  ${tampered.join('\n  ')}`);
  }
  return record.map((entry) => ({ ...entry, filename: join(tarballs, entry.file) }));
}

function packAndRecord(): readonly Packed[] {
  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(tarballs, { recursive: true });
  const packed = pack(tarballs);
  const record: PackedRecord[] = packed.map((entry) => ({
    name: entry.name,
    version: entry.version,
    file: basename(entry.filename),
    integrity: integrityOf(entry.filename),
  }));
  writeFileSync(RECORD, `${JSON.stringify(record, null, 2)}\n`);
  return packed;
}

const packed = mode === '--packed' ? recorded() : packAndRecord();
console.log(
  `${mode === '--packed' ? 'recorded' : 'packed'} ${String(packed.length)} candidate(s):`,
);
for (const entry of packed) console.log(`  ${entry.name}@${entry.version}`);

/** Install the candidates into the website and run everything it declares. */
function consume(packed: readonly Packed[]): void {
  // The website is the consumer: one app, one smoke test, and the only thing that
  // changes between this gate, the canary and the deploy is where the packages
  // came from. Copied out so the working tree keeps its registry ranges.
  const consumer = join(workDir, 'consumer');
  rmSync(consumer, { recursive: true, force: true });
  cpSync(repoPath('website'), consumer, {
    recursive: true,
    filter: (from) => !GENERATED.test(from),
  });

  const website = JSON.parse(readFileSync(repoPath('website/package.json'), 'utf8')) as {
    scripts: Record<string, string>;
    dependencies: Record<string, string>;
    devDependencies: Record<string, string>;
    overrides?: Record<string, unknown>;
  };
  // `pnpm version-packages` writes these ranges; one it did not write is a
  // manifest that no longer says what it was proved against.
  const stale = staleRanges(website.dependencies, packed, SCOPE);
  if (stale.length > 0) {
    throw new Error(
      'website/package.json does not name the versions this commit would publish:\n' +
        `  ${stale.join('\n  ')}\n` +
        'node tools/release/candidate/website.ts writes them',
    );
  }

  const offScope = Object.fromEntries(
    Object.entries(website.dependencies).filter(([name]) => !name.startsWith(SCOPE)),
  );

  const manifest = scratchManifest(
    packed,
    {
      dependencies: { ...offScope, ...website.devDependencies },
      scripts: website.scripts,
      overrides: website.overrides ?? {},
    },
    SCOPE,
  );
  writeFileSync(join(consumer, 'package.json'), `${JSON.stringify(manifest, null, 2)}\n`);

  // A lookup for our own scope now has nowhere to go, so a fallback to the
  // registry is an error rather than a pass against the previous release.
  writeFileSync(join(consumer, '.npmrc'), `${SCOPE.slice(0, -1)}:registry=${NOWHERE}\n`);

  console.log('installing the candidates into a scratch consumer...');
  run('npm', ['install', '--no-audit', '--no-fund'], consumer);

  const lock = JSON.parse(readFileSync(join(consumer, 'package-lock.json'), 'utf8')) as Lockfile;
  const failures = provenanceFailures(lock, packed, SCOPE, integrityOf);
  if (failures.length > 0) {
    throw new Error(
      `the scratch install did not come from this commit's tarballs:\n  ${failures.join('\n  ')}`,
    );
  }
  console.log(
    `provenance: ${String(packed.length)} package(s), every one from its candidate tarball`,
  );

  run('npm', ['run', 'typecheck'], consumer);
  console.log('typecheck: clean');
  run('npm', ['run', 'lint'], consumer);
  console.log('lint: clean');
  run('npm', ['test'], consumer);
  run('npm', ['run', 'smoke'], consumer);
  run('npm', ['run', 'build'], consumer);
  console.log('the packages this commit would publish work for a consumer');
}

if (mode !== '--pack-only') consume(packed);

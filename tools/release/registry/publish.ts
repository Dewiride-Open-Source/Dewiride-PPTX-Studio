/**
 * Publish the tarballs the release proved, and nothing else.
 *
 * ```
 * node tools/release/registry/publish.ts <tarball-dir>
 * ```
 *
 * Runs in the release's `publish` job, which holds `id-token: write` and installs nothing. The
 * directory's `packed.json` has already been held to the digest `prove` recorded before any
 * consumer code ran; each tarball is held to `packed.json` here. ADR 0059.
 */

import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { readPackedRecord, tamperedTarballs } from '../candidate/consumer.ts';
import { toolCommand, type Tool } from '../candidate/spawn.ts';
import { escapedName } from './oidc.ts';

const REGISTRY = 'https://registry.npmjs.org';

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: publish.ts <tarball-dir>');
const tarballs = resolve(dir);

const packed = readPackedRecord(readFileSync(join(tarballs, 'packed.json'), 'utf8'));
const integrityOf = (file: string): string =>
  `sha512-${createHash('sha512')
    .update(readFileSync(join(tarballs, file)))
    .digest('base64')}`;

const tampered = tamperedTarballs(packed, integrityOf);
if (tampered.length > 0) {
  throw new Error(
    `refusing to publish tarballs that are not the ones packed:\n  ${tampered.join('\n  ')}`,
  );
}

function run(tool: Tool | 'git', args: readonly string[]): void {
  const command =
    tool === 'git'
      ? { file: 'git', args, env: process.env }
      : toolCommand(tool, args, {
          platform: process.platform,
          execPath: process.execPath,
          env: process.env,
          exists: existsSync,
        });
  const result = spawnSync(command.file, command.args, { env: command.env, stdio: 'inherit' });
  if (result.error !== undefined) throw result.error;
  if (result.status !== 0) {
    throw new Error(
      `${tool} ${args.join(' ')} exited with ${String(result.status ?? result.signal)}`,
    );
  }
}

let published = 0;
for (const entry of packed) {
  const at = `${REGISTRY}/${escapedName(entry.name)}/${entry.version}`;
  const response = await fetch(at, { cache: 'no-store' });
  if (response.ok) {
    console.log(`on npm      ${entry.name}@${entry.version}`);
    continue;
  }
  if (response.status !== 404) {
    throw new Error(
      `${entry.name}@${entry.version}: the registry answered ${String(response.status)}`,
    );
  }
  run('npm', ['publish', join(tarballs, entry.file), '--access', 'public', '--provenance']);
  const tag = `${entry.name}@${entry.version}`;
  run('git', ['tag', '-a', tag, '-m', tag]);
  console.log(`published   ${tag}`);
  published += 1;
}
console.log(
  `\n${String(published)} published, ${String(packed.length - published)} already on npm`,
);

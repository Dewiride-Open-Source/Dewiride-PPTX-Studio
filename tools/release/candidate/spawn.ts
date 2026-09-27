/**
 * How the candidate gate starts npm and pnpm, with no shell on any platform.
 *
 * On Windows both are `.cmd` shims, which Node refuses to spawn without a shell
 * (CVE-2024-27980), so this runs `node.exe` on the script the shim would run. ADR 0058.
 */

import { basename, dirname, join } from 'node:path';

export type Tool = 'npm' | 'pnpm';

/** What the resolver reads from the running process, injected so every branch is testable. */
export interface Host {
  readonly platform: NodeJS.Platform;
  readonly execPath: string;
  readonly env: Readonly<Record<string, string | undefined>>;
  readonly exists: (path: string) => boolean;
}

/** A command and its arguments, each argument passed as one element. */
export interface Command {
  readonly file: string;
  readonly args: readonly string[];
  readonly env: Readonly<Record<string, string | undefined>>;
}

const PNPM_ENTRY = /^pnpm(?:\.[cm]?js|\.exe)?$/i;
const SCRIPT = /\.[cm]?js$/i;

/** Variables pnpm sets for a script it runs, which npm would read as its own configuration. */
const PNPM_LIFECYCLE =
  /^(?:npm_config_|npm_lifecycle_|npm_package_)|^npm_(?:execpath|node_execpath)$/;

function withoutPnpmLifecycle(
  env: Readonly<Record<string, string | undefined>>,
): Record<string, string | undefined> {
  const out: Record<string, string | undefined> = {};
  for (const [name, value] of Object.entries(env)) {
    if (!PNPM_LIFECYCLE.test(name)) out[name] = value;
  }
  return out;
}

export function toolCommand(tool: Tool, args: readonly string[], host: Host): Command {
  if (host.platform !== 'win32') return { file: tool, args, env: host.env };
  if (tool === 'npm') {
    const cli = join(dirname(host.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js');
    if (!host.exists(cli)) throw new Error(`no npm beside ${host.execPath}: expected ${cli}`);
    return { file: host.execPath, args: [cli, ...args], env: withoutPnpmLifecycle(host.env) };
  }
  const entry = host.env['npm_execpath'];
  if (entry === undefined || !PNPM_ENTRY.test(basename(entry))) {
    throw new Error(
      'Node will not start the pnpm shim without a shell on Windows; run this as ' +
        '`pnpm candidate <work-dir>` so npm_execpath names pnpm itself',
    );
  }
  return SCRIPT.test(entry)
    ? { file: host.execPath, args: [entry, ...args], env: host.env }
    : { file: entry, args, env: host.env };
}

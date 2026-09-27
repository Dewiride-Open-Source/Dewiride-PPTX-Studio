/**
 * The candidate gate's commands, shown starting npm and pnpm without a shell on every platform.
 *
 * The Windows branch never runs in CI, whose candidate job is Linux, so it is proved here
 * against an injected host. ADR 0058.
 */

import { win32 } from 'node:path';

import { describe, expect, it } from 'vitest';

import { toolCommand, type Host } from './spawn.ts';

const NODE = 'C:\\Program Files\\nodejs\\node.exe';
const NPM_CLI = win32.join('C:\\Program Files\\nodejs', 'node_modules', 'npm', 'bin', 'npm-cli.js');

function windows(env: Record<string, string | undefined> = {}, present = [NPM_CLI]): Host {
  return { platform: 'win32', execPath: NODE, env, exists: (path) => present.includes(path) };
}

const linux: Host = {
  platform: 'linux',
  execPath: '/usr/bin/node',
  env: { PATH: '/usr/bin', npm_config_user_agent: 'pnpm/11' },
  exists: () => false,
};

describe('off Windows', () => {
  it('starts the tool by name, with its arguments as they were given', () => {
    const command = toolCommand('pnpm', ['pack', '--json'], linux);
    expect(command.file).toBe('pnpm');
    expect(command.args).toEqual(['pack', '--json']);
    expect(command.env).toBe(linux.env);
  });
});

describe('npm on Windows', () => {
  it('runs the npm bundled beside node.exe, which is what npm.cmd runs', () => {
    const command = toolCommand('npm', ['install', '--no-audit'], windows());
    expect(command.file).toBe(NODE);
    expect(command.args).toEqual([NPM_CLI, 'install', '--no-audit']);
  });

  it('refuses when there is no npm beside node.exe, rather than looking for a shell', () => {
    expect(() => toolCommand('npm', ['install'], windows({}, []))).toThrow(/no npm beside/);
  });

  it("drops pnpm's lifecycle variables, which npm would read as its own configuration", () => {
    const env = {
      PATH: 'C:\\bin',
      NPM_CONFIG_USERCONFIG: 'C:\\npmrc',
      npm_config_user_agent: 'pnpm/11.28.0',
      npm_config_node_gyp: 'C:\\gyp',
      npm_execpath: 'C:\\pnpm\\bin\\pnpm.mjs',
      npm_node_execpath: NODE,
      npm_lifecycle_event: 'candidate',
      npm_package_json: 'C:\\repo\\package.json',
    };
    expect(toolCommand('npm', ['install'], windows(env)).env).toEqual({
      PATH: 'C:\\bin',
      NPM_CONFIG_USERCONFIG: 'C:\\npmrc',
    });
  });
});

describe('pnpm on Windows', () => {
  it('runs the script pnpm itself was started from', () => {
    const entry = 'C:\\Users\\x\\AppData\\Local\\pnpm\\bin\\pnpm.mjs';
    const command = toolCommand('pnpm', ['-r', 'pack'], windows({ npm_execpath: entry }));
    expect(command.file).toBe(NODE);
    expect(command.args).toEqual([entry, '-r', 'pack']);
  });

  it('starts a native pnpm executable directly', () => {
    const entry = 'C:\\tools\\pnpm.exe';
    const command = toolCommand('pnpm', ['pack'], windows({ npm_execpath: entry }));
    expect(command.file).toBe(entry);
    expect(command.args).toEqual(['pack']);
  });

  it('refuses when it was not started through pnpm', () => {
    expect(() => toolCommand('pnpm', ['pack'], windows())).toThrow(/pnpm candidate/);
  });

  it('refuses an npm_execpath that names npm rather than pnpm', () => {
    const npmEntry = windows({ npm_execpath: NPM_CLI });
    expect(() => toolCommand('pnpm', ['pack'], npmEntry)).toThrow(/pnpm candidate/);
  });
});

describe('arguments', () => {
  it('passes a path with spaces, quotes and shell metacharacters as one element', () => {
    const awkward = 'C:\\work dir\\a "quoted" & $name %PATH%\\tarballs';
    for (const host of [linux, windows()]) {
      const command = toolCommand('npm', ['install', awkward], host);
      expect(command.args.at(-1), host.platform).toBe(awkward);
    }
  });
});

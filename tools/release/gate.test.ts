/**
 * The release gate names jobs, so the names have to be the ones CI runs.
 *
 * The gate refuses a name it cannot find rather than passing vacuously, so a
 * rename costs a failed dispatch instead of an unguarded publish. This turns
 * that into a failed test, which is cheaper to read. ADR 0046.
 */

import { readFileSync } from 'node:fs';

import { parse } from 'yaml';
import { describe, expect, it } from 'vitest';

import { repoPath } from '../repo/root.ts';

interface Step {
  readonly name?: string;
  readonly run?: string;
  readonly uses?: string;
  readonly env?: Readonly<Record<string, string>>;
}

interface Workflow {
  readonly on?: Readonly<Record<string, unknown>>;
  readonly jobs: Readonly<
    Record<
      string,
      {
        readonly name?: string;
        readonly permissions?: Readonly<Record<string, string>>;
        readonly steps?: readonly Step[];
      }
    >
  >;
}

function workflow(file: string): Workflow {
  return parse(readFileSync(repoPath(`.github/workflows/${file}`), 'utf8')) as Workflow;
}

const ci = workflow('ci.yml');
const release = workflow('release.yml');
const canary = workflow('canary.yml');

const jobNames = (w: Workflow): string[] =>
  Object.entries(w.jobs).map(([id, job]) => job.name ?? id);

const gateStep = release.jobs['release']?.steps?.find(
  (step) => step.env?.['REQUIRED'] !== undefined,
);
const required = (gateStep?.env?.['REQUIRED'] ?? '')
  .split('\n')
  .map((line) => line.trim())
  .filter((line) => line !== '');

describe('the release gate', () => {
  it('names jobs that CI actually declares', () => {
    expect(required.length).toBeGreaterThan(0);
    for (const name of required) expect(jobNames(ci), name).toContain(name);
  });

  it('requires the gate that consumes the candidate tarballs', () => {
    expect(required).toContain('the packages this commit would publish');
  });

  it('does not require anything that reaches the registry', () => {
    // The deadlock: a monitor of what is published cannot gate the release that
    // would fix it. `canary.yml` exists so this list cannot name one.
    for (const name of jobNames(canary)) expect(required).not.toContain(name);
  });
});

describe('the canary', () => {
  it('is a monitor, so it runs on a schedule and never on a push', () => {
    expect(Object.keys(canary.on ?? {}).sort()).toEqual(['schedule', 'workflow_dispatch']);
  });

  it('is not a job in CI, where the release gate would read it', () => {
    expect(jobNames(ci)).not.toContain('the example, installed from npm');
  });
});

const releaseSteps = release.jobs['release']?.steps ?? [];
const ranBy = (steps: readonly Step[], pattern: RegExp): number =>
  steps.findIndex((step) => pattern.test(step.run ?? ''));

describe('publishing over OIDC', () => {
  it('is proved for this release before anything is published', () => {
    const asked = ranBy(releaseSteps, /publishers[.]ts --require-pending/);
    const published = ranBy(releaseSteps, /changeset publish/);
    expect(asked, 'the release never asks npm whether it may publish').toBeGreaterThan(-1);
    expect(published).toBeGreaterThan(asked);
  });

  it('never lets a token onto the publish path, which would skip provenance', () => {
    // With NODE_AUTH_TOKEN set the CLI takes the legacy route and attests
    // nothing. Ten packages shipped unverifiable that way. ADR 0047.
    expect(JSON.stringify(release)).not.toContain('NODE_AUTH_TOKEN');
  });

  it('is asked of every package by the canary, not just the ones releasing', () => {
    const steps = canary.jobs['publishers']?.steps ?? [];
    expect(ranBy(steps, /publishers[.]ts --require-all/)).toBeGreaterThan(-1);
    expect(canary.jobs['publishers']?.permissions?.['id-token']).toBe('write');
  });

  it('mints those tokens in a job that installs nothing', () => {
    const job = canary.jobs['publishers'];
    for (const step of job?.steps ?? []) {
      expect(step.run ?? '', step.name).not.toMatch(/(npm|pnpm|npx) (install|add|ci)/);
    }
    const uses = (job?.steps ?? []).map((step) => step.uses).filter((u) => u !== undefined);
    const pinned = (u: string): boolean => /@[0-9a-f]{40}$/.test(u);
    const allowed = ['actions/checkout@', 'actions/setup-node@'];
    expect(uses.every((u) => allowed.some((a) => u.startsWith(a)) && pinned(u))).toBe(true);
  });

  it('gives the job that installs from the registry no way to mint them', () => {
    expect(canary.jobs['published']?.permissions?.['id-token']).toBeUndefined();
  });
});

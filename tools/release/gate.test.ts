/**
 * The release gate names jobs, so the names have to be the ones CI runs.
 *
 * The gate refuses a name it cannot find rather than passing vacuously, so a
 * rename costs a failed dispatch instead of an unguarded publish. This turns
 * that into a failed test, which is cheaper to read. ADR 0046.
 */

import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { delimiter, join } from 'node:path';

import { parse } from 'yaml';
import { describe, expect, it } from 'vitest';

import { repoPath } from '../repo/root.ts';

interface Step {
  readonly name?: string;
  readonly if?: string;
  readonly run?: string;
  readonly uses?: string;
  readonly env?: Readonly<Record<string, string>>;
}

interface Job {
  readonly name?: string;
  readonly needs?: string | readonly string[];
  readonly if?: string;
  readonly 'runs-on'?: string;
  readonly permissions?: Readonly<Record<string, string>>;
  readonly env?: Readonly<Record<string, string>>;
  readonly steps?: readonly Step[];
}

interface Workflow {
  readonly on?: Readonly<Record<string, unknown>>;
  readonly permissions?: Readonly<Record<string, string>>;
  readonly jobs: Readonly<Record<string, Job>>;
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

  it('requires the fixture every reader test re-derives to hold on a Windows runner', () => {
    const job = ci.jobs['probe-fonts'];
    expect(required).toContain(job?.name);
    expect(job?.['runs-on']).toMatch(/^windows/);
    const runs = (job?.steps ?? []).map((step) => step.run ?? '');
    expect(
      runs.some((run) => run.includes('--fixture corpus/ground-truth/font-metrics.json')),
    ).toBe(true);
    expect(
      runs.some((run) =>
        run.includes('git diff --exit-code -- corpus/ground-truth/font-metrics.json'),
      ),
    ).toBe(true);
  });

  it('does not require the macOS run, which measures an assumption on dispatch only', () => {
    const job = ci.jobs['real-faces-macos'];
    expect(job?.['runs-on']).toMatch(/^macos/);
    expect(job?.if).toBe("github.event_name == 'workflow_dispatch'");
    expect(required).not.toContain(job?.name);
  });

  it('builds the reader it scores from this commit, never from a cache', () => {
    // A replayed `dist` scores the previous reader, and every one of these jobs
    // scores `packages/cli/dist`. ADR 0044, closed in ADR 0052.
    for (const id of ['real-faces', 'real-faces-macos', 'probe-fonts']) {
      const build = ci.jobs[id]?.steps?.find((step) => /turbo run build/.test(step.run ?? ''));
      expect(build?.run, id).toContain('--force');
    }
  });
});

describe('the canary', () => {
  it('is a monitor, so it runs on a schedule and never on a push', () => {
    expect(Object.keys(canary.on ?? {}).sort()).toEqual(['schedule', 'workflow_dispatch']);
  });

  it('is not a job in CI, where the release gate would read it', () => {
    for (const name of jobNames(canary)) expect(jobNames(ci)).not.toContain(name);
  });
});

/** A line that runs the tool, as opposed to prose that names it. */
const INSTALLS = /^\s*(?:npm|pnpm|npx)\s/m;
const runsATool = (job: Job | undefined): boolean =>
  (job?.steps ?? []).some((step) => INSTALLS.test(step.run ?? ''));
const usesAnAction = (job: Job | undefined): boolean =>
  (job?.steps ?? []).some((step) => step.uses !== undefined);

describe('the canary has an owner', () => {
  const installing = canary.jobs['published'];
  const reporting = canary.jobs['report'];

  it('reports from a job of its own, after the install has finished', () => {
    expect(reporting?.needs).toBe('published');
    expect(reporting?.if, 'a red install must not skip the report').toContain('always()');
  });

  it('opens on red and closes on green, not the other way round', () => {
    const opens = reporting?.steps?.find((step) => /gh issue create/.test(step.run ?? ''));
    const closes = reporting?.steps?.find((step) => /gh issue close/.test(step.run ?? ''));
    expect(opens?.if).toContain("needs.published.result == 'failure'");
    expect(closes?.if).toContain("needs.published.result == 'success'");
  });

  it('holds the right to write issues where nothing from the registry runs', () => {
    // 0048's isolation for a different permission: third-party code installed
    // beside `issues: write` could file, edit and close issues as this repository.
    expect(reporting?.permissions).toEqual({ issues: 'write' });
    expect(runsATool(reporting)).toBe(false);
    expect(usesAnAction(reporting)).toBe(false);
  });

  it('gives the job that installs from the registry nothing to write with', () => {
    expect(runsATool(installing), 'the install detector no longer sees the install').toBe(true);
    expect(installing?.permissions?.['issues']).toBeUndefined();
    expect(JSON.stringify(installing)).not.toContain('GITHUB_TOKEN');
    expect(canary.permissions).toEqual({ contents: 'read' });
  });

  it('asks the registry what a consumer can make of every published version', () => {
    expect(JSON.stringify(installing)).toContain('registry/attestations.ts');
    expect(JSON.stringify(installing)).not.toContain('--require-all');
  });
});

const website = workflow('website.yml');

describe('the website', () => {
  it('is never required by the release and never a job in CI', () => {
    for (const name of jobNames(website)) {
      expect(required).not.toContain(name);
      expect(jobNames(ci)).not.toContain(name);
    }
  });

  it('follows a release rather than being run by one', () => {
    // A push made with GITHUB_TOKEN starts no workflow, so `push` never sees
    // the version commit; `workflow_run` does and widens nothing. ADR 0055.
    expect(website.on?.['workflow_run']).toMatchObject({
      workflows: ['Release'],
      types: ['completed'],
    });
    expect(JSON.stringify(release)).not.toContain('website.yml');
    expect(release.jobs['release']?.permissions?.['actions']).toBe('read');
  });

  it('holds the right to deploy where nothing from the registry runs', () => {
    const building = website.jobs['build'];
    const deploying = website.jobs['deploy'];
    expect(runsATool(building), 'the install detector no longer sees the install').toBe(true);
    expect(building?.permissions).toBeUndefined();
    expect(deploying?.permissions).toEqual({ pages: 'write', 'id-token': 'write' });
    expect(runsATool(deploying)).toBe(false);
    expect((deploying?.steps ?? []).some((s) => /actions\/checkout/.test(s.uses ?? ''))).toBe(
      false,
    );
    expect(website.permissions).toEqual({ contents: 'read' });
  });

  it('cannot publish, and never asks whether it may', () => {
    expect(JSON.stringify(website)).not.toContain('publishers.ts');
    expect(JSON.stringify(website)).not.toContain('NODE_AUTH_TOKEN');
  });

  // The retry loop must read npm's exit code, not that of whatever it is
  // piped through; a fake npm says which one the step is reading.
  describe('the install step', () => {
    const script = website.jobs['build']?.steps?.find((step) =>
      /^npm install|until npm install/m.test(step.run ?? ''),
    )?.run;

    function runWith(npm: string): { status: number | null; output: string } {
      const dir = mkdtempSync(join(tmpdir(), 'website-install-'));
      writeFileSync(join(dir, 'npm'), `#!/usr/bin/env bash\n${npm}\n`, { mode: 0o755 });
      mkdirSync(join(dir, 'temp'));
      const result = spawnSync('bash', ['-e', '-c', script ?? 'exit 99'], {
        cwd: dir,
        encoding: 'utf8',
        env: {
          ...process.env,
          PATH: `${dir}${delimiter}${process.env['PATH'] ?? ''}`,
          RUNNER_TEMP: join(dir, 'temp'),
        },
      });
      rmSync(dir, { recursive: true, force: true });
      return { status: result.status, output: `${result.stdout}${result.stderr}` };
    }

    it('exists', () => {
      expect(script).toBeDefined();
    });

    it('passes when npm does', () => {
      expect(runWith('echo "added 300 packages"; exit 0').status).toBe(0);
    });

    it('fails at once on an error that is not the registry lagging', () => {
      const { status, output } = runWith('echo "npm error code ERESOLVE"; exit 1');
      expect(status).not.toBe(0);
      expect(output).toContain('ERESOLVE');
      expect(output).not.toContain('asking again');
    });
  });
});

const releaseSteps = release.jobs['release']?.steps ?? [];
const ranBy = (steps: readonly Step[], pattern: RegExp): number =>
  steps.findIndex((step) => pattern.test(step.run ?? ''));

describe('publishing over OIDC', () => {
  it('is proved for this release before anything is published', () => {
    const asked = ranBy(releaseSteps, /publishers[.]ts/);
    const published = ranBy(releaseSteps, /changeset publish/);
    expect(asked, 'the release never asks npm whether it may publish').toBeGreaterThan(-1);
    expect(published).toBeGreaterThan(asked);
  });

  it('never lets a token onto the publish path, which would skip provenance', () => {
    // With NODE_AUTH_TOKEN set the CLI takes the legacy route and attests
    // nothing. Ten packages shipped unverifiable that way. ADR 0047.
    expect(JSON.stringify(release)).not.toContain('NODE_AUTH_TOKEN');
  });

  it('is never asked by a workflow no publisher names', () => {
    // npm answers 404 to any workflow it cannot match to a publisher on the
    // package, so a monitor could only ask by holding publish rights of its
    // own - which is a worse thing to have than the answer is worth. ADR 0049.
    expect(JSON.stringify(canary)).not.toContain('publishers.ts');
    for (const [id, job] of Object.entries(canary.jobs)) {
      expect(job.permissions?.['id-token'], id).toBeUndefined();
    }
  });
});

import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { roundTripPackage } from '@pptx-studio/writer';
import { describe, expect, it } from 'vitest';
import { parse } from 'yaml';
import {
  badgeFor,
  badgeText,
  committedDecks,
  roundTripCorpus,
  summaryFor,
  COLLECTIONS,
} from './roundtrip-report.ts';

/**
 * Sub-phase 1.6: the badge says what the corpus does, and keeps saying it.
 *
 * A badge is a claim made to people who will not run the tests, so the ways it
 * can go wrong are all quiet ones: the number goes stale, the gate that checks
 * it stops being run, or the workflow that was supposed to run it gets
 * reorganised around it. None of those fail anything on their own. So each gets
 * an assertion here.
 *
 * This file reaches the round trip through Vitest's alias to **source**, while
 * `pnpm roundtrip` reaches it through `dist`. That is not duplication for its
 * own sake - it means a build that does not match the sources it came from
 * shows up as two different numbers rather than as one confident wrong one.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const BADGE = join(ROOT, '.github', 'badges', 'roundtrip.json');
const WORKFLOW = join(ROOT, '.github', 'workflows', 'ci.yml');

const report = roundTripCorpus(ROOT, roundTripPackage);

describe('the number on the README', () => {
  it('is the number the corpus actually produces', () => {
    // The whole point of committing the badge rather than publishing it from a
    // run: the answer is decided by the contents of this repository, so it can
    // be checked here, with no token, no bot and no write permission anywhere.
    expect(readFileSync(BADGE, 'utf8')).toBe(badgeText(report));
  });

  it('is a Shields endpoint document, in the shape Shields will read', () => {
    const badge = JSON.parse(readFileSync(BADGE, 'utf8')) as Record<string, unknown>;
    expect(badge['schemaVersion']).toBe(1);
    expect(badge['label']).toBe('round trip');
    expect(typeof badge['message']).toBe('string');
    expect(badge['message']).not.toBe('');
    // Shields caches, and a stale green badge over a red main is the failure
    // worth avoiding.
    expect(badge['cacheSeconds']).toBeLessThanOrEqual(300);
  });

  it('counts every committed deck and no other file', () => {
    // Not `readdirSync`. A manifest entry is one `pnpm corpus` has checked a
    // licence on, and counting a file no manifest claims would put an
    // unlicensed deck into the number on the front page.
    expect(report.total).toBe(committedDecks(ROOT).length);
    expect(report.total).toBe(52);
    expect(badgeFor(report).message).toBe('52/52 decks');
    expect(badgeFor(report).color).toBe('brightgreen');
  });

  it('agrees with what 1.4 pinned, part for part', () => {
    // The same 1474 the round-trip suite asserts, reached by a different route.
    expect(report.parts).toBe(1474);
    expect(report.xml).toBe(875);
    expect(report.relationships).toBe(553);
    expect(report.binary).toBe(46);
    expect(report.passed).toBe(report.total);
  });

  it('goes red rather than amber when a deck fails', () => {
    // Fifty-one of fifty-two is not 98% working - it means a package does not
    // survive being read and written - and an amber badge invites reading it
    // as a score.
    const failed = { ...report, ok: false, passed: report.total - 1 };
    expect(badgeFor(failed).color).toBe('red');
    expect(badgeFor(failed).message).toBe('51/52 decks');
  });
});

describe('the summary a pull request shows', () => {
  const summary = summaryFor(report);

  it('breaks the total down by collection, so a missing one is visible', () => {
    for (const collection of COLLECTIONS) {
      expect(summary, collection).toContain('| `' + collection + '` |');
    }
    expect(summary).toContain('| **total** | **52** | **1474** |');
  });

  it('says what it does not measure', () => {
    // The one way this badge could mislead: it is about reading and writing,
    // not about PowerPoint. No hosted runner has Office on it, so the claim has
    // to be stated narrowly wherever it is made.
    expect(summary).toContain('does not say the decks open in PowerPoint');
  });

  it('names the decks that moved when any did', () => {
    const failed = {
      ...report,
      ok: false,
      passed: report.total - 1,
      decks: report.decks.map((deck, index) =>
        index === 0 ? { ...deck, ok: false, differences: ['xml /ppt/slides/slide1.xml'] } : deck,
      ),
    };
    const text = summaryFor(failed);
    expect(text).toContain('did not survive the round trip');
    expect(text).toContain('### What moved');
    expect(text).toContain('xml /ppt/slides/slide1.xml');
  });
});

describe('the gate that keeps it honest', () => {
  it('runs on every pull request, in a job of its own', () => {
    const workflow = parse(readFileSync(WORKFLOW, 'utf8')) as {
      on: Record<string, unknown>;
      jobs: Record<string, { steps: { run?: string }[] }>;
    };
    expect(Object.keys(workflow.on)).toContain('pull_request');

    const job = workflow.jobs['roundtrip'];
    expect(job, 'ci.yml has no `roundtrip` job').toBeDefined();
    expect(job!.steps.some((step) => step.run === 'pnpm roundtrip')).toBe(true);
  });

  it('runs locally too, inside `pnpm check`', () => {
    // A gate that only runs in CI is one a contributor meets after pushing.
    const manifest = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as {
      scripts: Record<string, string>;
    };
    expect(manifest.scripts['check']).toContain('pnpm roundtrip');
    // ...and after the build, because the script loads the writer from `dist`.
    const check = manifest.scripts['check']!;
    expect(check.indexOf('pnpm build')).toBeLessThan(check.indexOf('pnpm roundtrip'));
  });

  it('never leaves the workflow able to push', () => {
    // `checkout` puts the token in `.git/config` unless told not to, where
    // every later step can read it. Nothing in this workflow pushes.
    const workflow = parse(readFileSync(WORKFLOW, 'utf8')) as {
      permissions: Record<string, string>;
      jobs: Record<string, { steps: { uses?: string; with?: Record<string, unknown> }[] }>;
    };
    expect(workflow.permissions).toEqual({ contents: 'read' });

    for (const [name, job] of Object.entries(workflow.jobs)) {
      for (const step of job.steps) {
        if (step.uses?.startsWith('actions/checkout@') !== true) continue;
        expect(step.with?.['persist-credentials'], name).toBe(false);
      }
    }
  });

  it('pins every action to a commit, because a tag can move', () => {
    const workflow = parse(readFileSync(WORKFLOW, 'utf8')) as {
      jobs: Record<string, { steps: { uses?: string }[] }>;
    };
    for (const [name, job] of Object.entries(workflow.jobs)) {
      for (const step of job.steps) {
        if (step.uses === undefined) continue;
        expect(step.uses, name + ': ' + step.uses).toMatch(/@[0-9a-f]{40}$/);
      }
    }
  });
});

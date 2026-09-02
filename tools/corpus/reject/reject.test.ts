import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { formatReport, RULE_IDS, validatePackage } from '@pptx-studio/validate';
import { describe, expect, it } from 'vitest';
import { packageBytes } from './deck.ts';
import { outputName, REJECT_FIXTURES } from './fixtures.ts';

/**
 * `C-REJECT`: every fixture in `corpus/reject/` is refused, by the rule that
 * claims it.
 *
 * The fifth corpus rule, and the only one that runs the other direction.
 * `C-CENSUS`, `C-REGEN`, `C-LEX` and `C-COV` are all about files that are
 * right. This one is about files that are wrong, and it exists because the
 * other direction cannot be checked with real decks at all: every deck in
 * `corpus/decks`, `corpus/authored` and `corpus/written` opens in PowerPoint,
 * so none of them contains a `p:control`, a `c:strLit` in a series title or a
 * placeholder typed `hdr`. Four of the validator's twenty-nine rules have
 * nothing in the good corpus to be quiet about, and without these fixtures
 * they would be enforced entirely on trust.
 *
 * Three assertions, and the third is the one that stops the collection rotting:
 *
 * 1. the fixture is refused - `report.ok` is false;
 * 2. by the rule its manifest entry names, fatally;
 * 3. and the manifest on disk still matches what the recipe builds, which is
 *    `C-REGEN` for this collection.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../../..');
const DIR = join(ROOT, 'corpus', 'reject');

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

interface Envelope {
  readonly collection: string;
  readonly entries: readonly {
    readonly id: string;
    readonly kind: string;
    readonly storage: string;
    readonly path: string;
    readonly sha256: string;
    readonly rule: string;
    readonly refusal: string;
    readonly tags: readonly string[];
  }[];
}

const MANIFEST = JSON.parse(readFileSync(join(DIR, 'manifest.json'), 'utf8')) as Envelope;

describe('C-REJECT', () => {
  it('has a fixture for every refusal the roster records as one we can check', () => {
    expect(REJECT_FIXTURES).toHaveLength(17);
    expect(MANIFEST.entries).toHaveLength(17);
    expect(MANIFEST.collection).toBe('reject');
  });

  it('refuses every one of them', () => {
    const accepted = REJECT_FIXTURES.filter((fixture) => {
      const bytes = new Uint8Array(readFileSync(join(DIR, outputName(fixture))));
      return validatePackage({ bytes }).ok;
    }).map((fixture) => fixture.id);

    // A fixture that validates clean is either a rule that has quietly stopped
    // working or a fixture that no longer contains what it claims to. Either
    // way the collection is lying, and that is worse than not having it.
    expect(accepted).toEqual([]);
  });

  it('refuses each one for the reason its manifest gives', () => {
    const wrong: string[] = [];
    for (const fixture of REJECT_FIXTURES) {
      const bytes = new Uint8Array(readFileSync(join(DIR, outputName(fixture))));
      const report = validatePackage({ bytes });
      const fatal = report.findings.filter((finding) => finding.severity === 'fatal');
      if (!fatal.some((finding) => finding.rule === fixture.rule)) {
        wrong.push(fixture.id + ' should trip ' + fixture.rule + '\n' + formatReport(report));
      }
    }

    // Named rather than counted. "Something was refused" is not the claim -
    // a fixture caught by the wrong rule means we would give a user the wrong
    // reason, which for a file PowerPoint refuses with no diagnostic at all is
    // most of the value gone.
    expect(wrong).toEqual([]);
  });

  it('names a rule that exists, in every entry', () => {
    for (const entry of MANIFEST.entries) {
      expect(RULE_IDS, entry.id).toContain(entry.rule);
      expect(entry.tags, entry.id).toEqual(['reject', entry.rule]);
      // Each entry carries what PowerPoint said. It is the only diagnostic
      // there is, so a fixture without one is a claim nobody can re-check.
      expect(entry.refusal.length, entry.id).toBeGreaterThan(80);
    }
  });

  it('covers the four rules no good deck can exercise', () => {
    // The gap `tools/corpus/validate.test.ts` writes down: `V022`, `V024`,
    // `V025` and `V026` have no instance anywhere in the fifty-two decks,
    // because every one of them is a file PowerPoint refuses.
    const covered = new Set(REJECT_FIXTURES.map((fixture) => fixture.rule));
    for (const id of ['V022', 'V024', 'V025', 'V026']) expect(covered).toContain(id);
  });

  it('rebuilds byte-for-byte from its recipe', () => {
    const wrong: string[] = [];
    for (const fixture of REJECT_FIXTURES) {
      const rebuilt = packageBytes(fixture.parts);
      const onDisk = new Uint8Array(readFileSync(join(DIR, outputName(fixture))));
      if (sha256(rebuilt) !== sha256(onDisk)) wrong.push(fixture.id);
    }
    expect(wrong).toEqual([]);
  });

  it('agrees with its own manifest about every hash', () => {
    for (const entry of MANIFEST.entries) {
      const onDisk = new Uint8Array(readFileSync(join(DIR, entry.path)));
      expect(sha256(onDisk), entry.id).toBe(entry.sha256);
      expect(entry.kind, entry.id).toBe('fixture');
    }
  });
});

describe('what the reject collection does not claim', () => {
  it('is not a census claim, so it cannot inflate C-COV', () => {
    // `C-COV` counts a census feature as covered when some entry of
    // `kind: "deck"` reports a non-zero count for it. These are packages
    // PowerPoint refuses; that a feature appears in one is not evidence the
    // feature works, and typing them as decks would have quietly widened the
    // coverage number by seventeen files nothing can open.
    for (const entry of MANIFEST.entries) {
      expect(entry.kind).toBe('fixture');
      expect(entry).not.toHaveProperty('features');
    }
  });
});

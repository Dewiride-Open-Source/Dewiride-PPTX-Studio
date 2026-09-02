import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { formatReport, RULES, validatePackage, type Report } from '@pptx-studio/validate';
import { describe, expect, it } from 'vitest';

/**
 * The other half of sub-phase 1.2's verification, and the harder half to fake.
 *
 * `packages/validate/src/validate.test.ts` breaks a deck twenty-nine ways and
 * checks that each rule fires. That proves the rules can detect. It cannot
 * prove the thing a validator actually has to be: **quiet on good files.**
 *
 * This does, and it does it against the one collection where the claim is
 * externally checkable. Every deck in `corpus/` opens in PowerPoint - that is
 * what the corpus *is*, established one bisection at a time across sub-phases
 * 0.7 and 1.1, with `C-REGEN` holding the bytes in place. So a rule that fires
 * on one of them is not a discovery, it is a bug, and there is no third
 * possibility to argue about.
 *
 * That asymmetry is worth stating plainly, because it is why a validator is a
 * risky thing to build. A false negative costs one missed defect. A false
 * positive blocks an export the user wanted, and a tool that blocks exports the
 * user wanted gets switched off - after which it catches nothing at all.
 * Fifty-one files that must produce silence is the cheapest insurance available
 * against writing the rules too tightly.
 *
 * ## Why this lives in `tools/` and not beside the package
 *
 * The core packages test in real Chromium, deliberately: under jsdom a `node:fs`
 * import resolves and a Node API leak stays invisible until somebody opens a
 * tab. A browser test therefore has no filesystem, and the corpus is on disk.
 * `tools/` runs under Node with the same Vitest alias table, so it can import
 * `@pptx-studio/validate` and read the decks. The split is the same one
 * `census-drift.test.ts` and `lexical.test.ts` already live on.
 */

const ROOT = resolve(fileURLToPath(import.meta.url), '../../..');
const COLLECTIONS = ['decks', 'authored', 'written'];

interface Envelope {
  readonly entries: readonly {
    readonly id: string;
    readonly kind: string;
    readonly storage: string;
    readonly path?: string;
  }[];
}

interface Deck {
  readonly id: string;
  readonly collection: string;
  readonly bytes: Uint8Array;
}

function committedDecks(): Deck[] {
  const decks: Deck[] = [];
  for (const collection of COLLECTIONS) {
    const dir = join(ROOT, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as Envelope;
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      decks.push({
        id: entry.id,
        collection,
        bytes: new Uint8Array(readFileSync(join(dir, entry.path ?? ''))),
      });
    }
  }
  return decks;
}

const DECKS = committedDecks();
const REPORTS = new Map<string, Report>(
  DECKS.map((deck) => [deck.id, validatePackage({ bytes: deck.bytes })]),
);

describe('the validator against the corpus', () => {
  it('reads all fifty-one committed decks', () => {
    expect(DECKS).toHaveLength(51);
  });

  it('finds nothing fatal in a single one of them', () => {
    // Reported as `id: <the whole report>` rather than as a count, because when
    // this fails the only useful information is *which rule* fired on *which
    // part of which deck* - and the answer decides whether the rule is wrong or
    // the deck is.
    const failures = DECKS.filter((deck) =>
      REPORTS.get(deck.id)!.findings.some((finding) => finding.severity === 'fatal'),
    ).map((deck) => deck.id + '\n' + formatReport(REPORTS.get(deck.id)!));

    expect(failures).toEqual([]);
  });

  it('warns exactly twice, and both are placeholders a deck meant to leave orphaned', () => {
    // Pinned rather than driven to zero, because both are correct - and the
    // first is the nicest result in this file. `a02-placeholders` exists to
    // probe the five-tier matcher, and its fifth shape is commented in the
    // generator as "Tier 5: orphan. Nothing to inherit". The rule found it and
    // described it in the deck's own words, without being told it was there.
    //
    // The second is the same thing incidentally: `a19-decorative` needed a
    // placeholder to hang `@descr` and `@title` on, gave it explicit geometry,
    // and never added a counterpart to the layout. It renders correctly because
    // it carries its own transform, which is exactly the case where a warning
    // is right and a refusal would not be.
    //
    // Zero fatal findings above and two accounted-for warnings here is the
    // whole claim of this file. Any third warning is a matcher bug or a new
    // deck, and either way somebody has to look.
    const warnings = DECKS.flatMap((deck) =>
      REPORTS.get(deck.id)!.findings.map(
        (finding) => deck.id + ' ' + finding.rule + ' ' + finding.where.xpath,
      ),
    );

    expect(warnings).toEqual([
      'a02-placeholders V021 /p:sld/p:cSld/p:spTree/p:sp[5]/p:nvSpPr/p:nvPr/p:ph',
      'a19-decorative V021 /p:sld/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:nvPr/p:ph',
    ]);
    for (const deck of DECKS) {
      for (const finding of REPORTS.get(deck.id)!.findings) {
        expect(finding.severity, deck.id).toBe('warning');
      }
    }
  });

  it('could read every part of every deck', () => {
    // A rule that never ran because a part would not parse is a rule that
    // passed for the wrong reason, and `Report.problems` is where that is
    // admitted. Zero across the corpus means the silence above is real silence.
    const unreadable = DECKS.flatMap((deck) =>
      REPORTS.get(deck.id)!.problems.map(
        (problem) => deck.id + ' ' + problem.part + ': ' + problem.message,
      ),
    );

    expect(unreadable).toEqual([]);
  });

  it('ran twenty-six rules on each; the other three want a baseline', () => {
    for (const deck of DECKS) {
      const report = REPORTS.get(deck.id)!;
      expect(report.checked, deck.id).toHaveLength(26);
      expect(
        report.skipped.map((entry) => entry.rule),
        deck.id,
      ).toEqual(['V027', 'V028', 'V029']);
    }
  });
});

describe('what the corpus exercises', () => {
  /**
   * Which rules have anything to look at.
   *
   * Not a coverage metric and not a target. It is here for the same reason
   * `coverage.test.ts` records that fourteen census keys rest on a single deck
   * each: "the corpus is silent" reads like a strong result, and part of it is
   * only silence about markup that is not there. A rule with no instance in the
   * corpus has been checked for false positives against nothing.
   */
  it('records which rules have no instance to be quiet about', () => {
    const noInstance = [
      // Nothing in the corpus is an ActiveX control, a chart style with a
      // missing entry, a `c:strLit` in a series title, or a `hdr`/`sldImg`
      // placeholder outside the notes family - because every one of those is a
      // file PowerPoint refuses, and the corpus is decks that open.
      'V022',
      'V024',
      'V025',
      'V026',
    ];

    // The list above is precisely `corpus/reject/`'s job in sub-phase 1.2's
    // second half: fixtures that must *fail*, for rules real decks cannot
    // exercise. Stated here so the gap is written down rather than implied.
    for (const id of noInstance) {
      expect(RULES.some((rule) => rule.id === id && rule.category === 'refused')).toBe(true);
    }
  });
});

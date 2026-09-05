/**
 * The number the badge carries, and the arithmetic behind it.
 *
 * ```
 * pnpm roundtrip          # check the committed badge against the corpus
 * pnpm roundtrip --write  # regenerate it
 * ```
 *
 * Sub-phase 1.6 is one line of plan - *round-trip job on every PR; badge* -
 * and the interesting part is what makes a badge worth believing.
 *
 * ## Why the count is committed rather than published from a run
 *
 * The usual way to get a number onto a README is a workflow that writes to a
 * gist or pushes a commit, which needs a token, a bot identity and write
 * permission on a workflow that otherwise needs none. All of that to publish a
 * fact that is **already determined by the contents of the repository**: the
 * corpus is committed, the writer is committed, and the round trip is
 * deterministic. Nothing about a CI run decides the answer.
 *
 * So the answer is committed too, as a Shields endpoint document, and CI
 * recomputes it and fails if the two disagree. The badge cannot claim 52/52
 * unless the last run to touch it measured 52/52, and no workflow needs write
 * access to anything.
 *
 * ## What it does not measure
 *
 * That the decks open in PowerPoint. Nothing on a Linux runner can ask, and
 * GitHub's hosted Windows images do not ship Office either - what they carry is
 * Visual Studio's Office *development* workload, which is not the same thing and
 * cannot open a file. `bisect --oracle powerpoint` is the answer to that
 * question and it runs on a machine with PowerPoint on it, by hand. The badge
 * says "round-trips", and that is the whole of what it says.
 *
 * ## No workspace import here, on purpose
 *
 * Nothing links `@pptx-studio/*` into `tools/`, so under plain `node` the
 * specifier does not resolve - the same constraint `write-census-keys.ts`
 * documents. Rather than reach into `dist` from the middle of the arithmetic,
 * the round trip arrives as a parameter: the CLI hands in the built one, the
 * test hands in the one Vitest aliases to source. Both paths run the same
 * counting, which is the part that could be wrong.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/** The shape of `roundTripPackage`'s result, named structurally. */
export interface RoundTripLike {
  readonly ok: boolean;
  readonly comparison: {
    readonly parts: readonly unknown[];
    readonly counts: {
      readonly xml: number;
      readonly binary: number;
      readonly relationships: number;
      readonly same: number;
    };
    readonly differences: readonly { readonly kind: string; readonly part: string }[];
  };
}

export type RoundTrip = (bytes: Uint8Array) => RoundTripLike;

/** The collections whose decks the badge counts. `reject` is not one of them. */
export const COLLECTIONS = ['decks', 'authored', 'written'] as const;

export interface DeckResult {
  readonly id: string;
  readonly collection: string;
  readonly bytes: number;
  readonly parts: number;
  readonly xml: number;
  readonly relationships: number;
  readonly binary: number;
  readonly ok: boolean;
  /** Empty when it round-tripped. Names what moved when it did not. */
  readonly differences: readonly string[];
}

export interface CorpusRoundTrip {
  readonly ok: boolean;
  readonly passed: number;
  readonly total: number;
  readonly parts: number;
  readonly xml: number;
  readonly relationships: number;
  readonly binary: number;
  readonly decks: readonly DeckResult[];
}

interface ManifestEntry {
  readonly id: string;
  readonly kind: string;
  readonly storage: string;
  readonly path?: string;
}

/**
 * The committed decks, from the manifests rather than from the directory.
 *
 * The same source `roundtrip.test.ts` reads, and deliberately not `readdirSync`:
 * a manifest entry is the thing `pnpm corpus` has already checked a licence on,
 * and counting a file that no manifest claims would put an unlicensed deck into
 * the number on the README.
 */
export function committedDecks(root: string): { id: string; collection: string; file: string }[] {
  const decks: { id: string; collection: string; file: string }[] = [];
  for (const collection of COLLECTIONS) {
    const dir = join(root, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as {
      entries: ManifestEntry[];
    };
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      decks.push({ id: entry.id, collection, file: join(dir, entry.path ?? '') });
    }
  }
  return decks;
}

export function roundTripCorpus(root: string, roundTrip: RoundTrip): CorpusRoundTrip {
  const decks: DeckResult[] = [];
  for (const deck of committedDecks(root)) {
    const bytes = new Uint8Array(readFileSync(deck.file));
    const result = roundTrip(bytes);
    const { counts, parts, differences } = result.comparison;
    decks.push({
      id: deck.id,
      collection: deck.collection,
      bytes: bytes.length,
      parts: parts.length,
      xml: counts.xml,
      relationships: counts.relationships,
      binary: counts.binary,
      ok: result.ok,
      differences: differences.map((difference) => difference.kind + ' ' + difference.part),
    });
  }

  const sum = (pick: (deck: DeckResult) => number): number =>
    decks.reduce((total, deck) => total + pick(deck), 0);

  return {
    ok: decks.every((deck) => deck.ok),
    passed: decks.filter((deck) => deck.ok).length,
    total: decks.length,
    parts: sum((deck) => deck.parts),
    xml: sum((deck) => deck.xml),
    relationships: sum((deck) => deck.relationships),
    binary: sum((deck) => deck.binary),
    decks,
  };
}

/** A Shields endpoint document: `schemaVersion`, `label`, `message`, `color`. */
export interface Badge {
  readonly schemaVersion: 1;
  readonly label: string;
  readonly message: string;
  readonly color: string;
  readonly cacheSeconds: number;
}

export function badgeFor(report: CorpusRoundTrip): Badge {
  return {
    schemaVersion: 1,
    label: 'round trip',
    message: String(report.passed) + '/' + String(report.total) + ' decks',
    // `brightgreen` and `red`, and nothing in between. A round trip that loses
    // one deck out of fifty-two is not 98% working: it means a part of some
    // package does not survive being read and written, and an amber badge would
    // invite reading that as a score rather than as a broken invariant.
    color: report.ok ? 'brightgreen' : 'red',
    // Shields caches aggressively and a stale green badge over a red main is
    // the failure mode worth avoiding. Five minutes is short enough to notice.
    cacheSeconds: 300,
  };
}

/** Exactly how the file is written: two-space JSON, one trailing newline. */
export function badgeText(report: CorpusRoundTrip): string {
  return JSON.stringify(badgeFor(report), null, 2) + '\n';
}

/** GitHub-flavoured markdown for the job summary on a pull request. */
export function summaryFor(report: CorpusRoundTrip): string {
  const lines: string[] = [
    '## Round trip',
    '',
    report.ok
      ? '**' +
        String(report.passed) +
        '/' +
        String(report.total) +
        ' decks** read, written back, and still the same document.'
      : '**' +
        String(report.total - report.passed) +
        ' of ' +
        String(report.total) +
        ' decks did not survive the round trip.**',
    '',
    '| collection | decks | parts | xml | rels | binary |',
    '| --- | ---: | ---: | ---: | ---: | ---: |',
  ];

  for (const collection of COLLECTIONS) {
    const decks = report.decks.filter((deck) => deck.collection === collection);
    if (decks.length === 0) continue;
    const sum = (pick: (deck: DeckResult) => number): string =>
      String(decks.reduce((total, deck) => total + pick(deck), 0));
    lines.push(
      '| `' +
        collection +
        '` | ' +
        String(decks.length) +
        ' | ' +
        sum((deck) => deck.parts) +
        ' | ' +
        sum((deck) => deck.xml) +
        ' | ' +
        sum((deck) => deck.relationships) +
        ' | ' +
        sum((deck) => deck.binary) +
        ' |',
    );
  }

  lines.push(
    '| **total** | **' +
      String(report.total) +
      '** | **' +
      String(report.parts) +
      '** | ' +
      String(report.xml) +
      ' | ' +
      String(report.relationships) +
      ' | ' +
      String(report.binary) +
      ' |',
    '',
  );

  const failures = report.decks.filter((deck) => !deck.ok);
  if (failures.length > 0) {
    lines.push('### What moved', '');
    for (const deck of failures) {
      lines.push('- **' + deck.id + '** — ' + deck.differences.slice(0, 8).join(', '));
    }
    lines.push('');
  }

  lines.push(
    'Parts are compared as documents, never as bytes: canonical XML per XML part,',
    'the relationship graph with ids treated as opaque labels, SHA-256 on everything',
    'else. This does not say the decks open in PowerPoint — no hosted runner has',
    'Office on it — only that reading and writing them changes nothing.',
    '',
  );

  return lines.join('\n');
}

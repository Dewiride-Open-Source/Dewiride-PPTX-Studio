import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { describe, expect, it } from 'vitest';
import { CENSUS_FEATURE_KEYS } from '../census/keys.gen.ts';
import type { UncoveredFeature } from '../manifest/schema.ts';

/**
 * `C-COV`, as a number rather than as a rule.
 *
 * The rule itself is `C019` in `check.ts`, and it runs in `pnpm corpus` on a
 * bare clone with nothing built - which is where it belongs, because the thing
 * it catches is somebody adding a census feature and no deck to exercise it.
 * This file is the other half: what the answer currently **is**.
 *
 * Worth separating for the reason `C-LEX` keeps `what C-LEX found` apart from
 * the rule that enforces it. A gate that passes tells you nobody broke the
 * invariant; it does not tell you the corpus covers forty-six of forty-seven
 * keys, that the one it misses is blocked on a download nobody approved, or
 * that fourteen of the forty-six rest on a single deck each. Those are the
 * facts a reader wants, and a fact nobody wrote down is a fact that changes
 * without anybody noticing.
 */

const COLLECTIONS = ['decks', 'authored', 'written', 'reject', 'bench', 'ground-truth'];

interface Envelope {
  readonly uncovered?: readonly UncoveredFeature[];
  readonly entries: readonly {
    readonly id: string;
    readonly kind: string;
    readonly storage: string;
    readonly features?: Readonly<Record<string, number>>;
  }[];
}

/** census key -> the decks exercising it, on the same terms `C019` uses. */
function coverage(): Map<string, string[]> {
  const found = new Map<string, string[]>();
  for (const collection of COLLECTIONS) {
    const path = join(ROOT, 'corpus', collection, 'manifest.json');
    const envelope = JSON.parse(readFileSync(path, 'utf8')) as Envelope;
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage === 'pinned') continue;
      for (const [key, count] of Object.entries(entry.features ?? {})) {
        if (count <= 0) continue;
        const decks = found.get(key);
        if (decks === undefined) found.set(key, [entry.id]);
        else decks.push(entry.id);
      }
    }
  }
  return found;
}

function declared(): UncoveredFeature[] {
  return COLLECTIONS.flatMap((collection) => {
    const path = join(ROOT, 'corpus', collection, 'manifest.json');
    return [...((JSON.parse(readFileSync(path, 'utf8')) as Envelope).uncovered ?? [])];
  });
}

const COVERED = coverage();
const UNCOVERED = declared();

describe('what C-COV says', () => {
  it('accounts for every census key, one way or the other', () => {
    const names = new Set(UNCOVERED.map((row) => row.key));
    const unaccounted = CENSUS_FEATURE_KEYS.filter(
      (key) => (COVERED.get(key)?.length ?? 0) === 0 && !names.has(key),
    );

    // The same assertion `C019` makes, from the other side of the build. It is
    // here as well because this file is what a reader opens to find out what
    // the corpus covers, and a summary that could be stale is worse than none.
    expect(unaccounted).toEqual([]);
  });

  it('covers forty-six of forty-seven', () => {
    const covered = CENSUS_FEATURE_KEYS.filter((key) => (COVERED.get(key)?.length ?? 0) > 0);

    expect(CENSUS_FEATURE_KEYS).toHaveLength(47);
    expect(covered).toHaveLength(46);
    expect(UNCOVERED).toHaveLength(1);
  });

  it('misses `model3d`, and says why in a form somebody can act on', () => {
    const gap = UNCOVERED[0]!;

    expect(gap.key).toBe('model3d');
    expect(COVERED.has('model3d')).toBe(false);
    // Not a spelling test. The gap is only useful if it names the experiment
    // that closes it and the slot that would hold the deck, because the next
    // person to read it will be deciding whether to run E4.
    expect(gap.closes).toContain('E4');
    expect(gap.closes).toContain('a28-model3d');
    expect(gap.why).toContain('no public specification');
  });

  it('rests thirteen of the forty-six on a single deck each', () => {
    // Not a violation and not a target: a probe corpus is built one feature at
    // a time, so most features have exactly one probe. It is here because
    // "46 of 47" reads like breadth and this is the shape underneath it - and
    // because it is the list to check first when a deck is about to be deleted.
    //
    // Fourteen until Gate 1. `a43-kitchen-sink` is the corpus's second deck
    // with a `vbaProject.bin` in it, so `macros` came off this list - which is
    // most of what a second probe for a feature is worth.
    const alone = [...COVERED]
      .filter(([, decks]) => decks.length === 1)
      .map(([key, decks]) => key + ' <- ' + decks[0]!)
      .sort();

    expect(alone).toEqual([
      'audio <- a24-media',
      'chartEx <- a22-chartex',
      'connector <- a06-lines',
      'contentPart <- a27-ink',
      'decorative <- a19-decorative',
      'groupFill <- a03-fills',
      'ink <- a27-ink',
      'innerShadow <- a04-effects',
      'math <- a29-math',
      'media <- a24-media',
      'scene3d <- a04-effects',
      'svgBlip <- a25-svg-blips',
      'video <- a24-media',
    ]);
  });

  it('claims nothing for a deck whose bytes are in no clone', () => {
    // `pinned` entries are excluded from the count above. There are none in the
    // corpus today, and this is what notices if one arrives and is quietly
    // allowed to satisfy a coverage claim with a file only one machine has.
    const pinnedDecks = COLLECTIONS.flatMap((collection) => {
      const path = join(ROOT, 'corpus', collection, 'manifest.json');
      const envelope = JSON.parse(readFileSync(path, 'utf8')) as Envelope;
      return envelope.entries.filter(
        (entry) => entry.kind === 'deck' && entry.storage === 'pinned',
      );
    });

    expect(pinnedDecks).toEqual([]);
  });
});

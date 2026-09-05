import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { describe, expect, it } from 'vitest';
import { inventory, type FormEvidence, type LexicalDeck } from './inventory.ts';
import { LEXICAL_FORMS, SERIALIZER, type LexicalForm } from './forms.ts';

/**
 * `C-LEX`, the last of the four named corpus rules.
 *
 * The rule: **every lexical convention is exercised by at least two decks from
 * at least two distinct serializers, or carries a written gap saying why not.**
 *
 * It waited for a second producer and now has three, and the reason it waited
 * is worth restating. Sub-phase 0.5's gate is "parse, serialize, byte-identical
 * for 100% of parts". Run against decks a single generator wrote, that gate
 * proves the generator and the serializer agree with each other - which they
 * would even if both were wrong. Only a second producer turns it into evidence
 * about the format.
 *
 * Two halves, and the second is the one that earns the rule its keep:
 *
 * 1. Nothing in the corpus is uncatalogued. Every form a fresh scan finds has a
 *    row, so a deck introducing a convention nobody wrote down fails here the
 *    way an unclaimed file fails `C008`.
 * 2. Nothing claims coverage it does not have. `producers` is pinned per row
 *    and compared against the scan, so a row cannot quietly gain or lose a
 *    producer, and a row below two producers must say why in prose.
 *
 * The table also carries forms this project measured **outside** the corpus -
 * ADR 0004's 2834 real parts, ADR 0002's 37 real packages - which is how a rule
 * about coverage says anything at all about what is missing. Those rows pin
 * `producers: []`, so if a deck ever grows one the assertion fails and somebody
 * gets to delete a gap.
 */

interface Envelope {
  readonly collection: string;
  readonly serializers?: { readonly xml: string; readonly container: string };
  readonly entries: readonly { id: string; kind: string; storage: string; path?: string }[];
}

/** Every committed deck in the corpus, with the serializers its manifest declares. */
function committedDecks(): LexicalDeck[] {
  const decks: LexicalDeck[] = [];
  for (const collection of ['decks', 'authored', 'written']) {
    const dir = join(ROOT, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as Envelope;
    const serializers = envelope.serializers;
    if (serializers === undefined) {
      throw new Error(collection + '/manifest.json declares no `serializers` (see C018)');
    }
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      decks.push({
        id: entry.id,
        collection,
        serializers,
        bytes: new Uint8Array(readFileSync(join(dir, entry.path ?? ''))),
      });
    }
  }
  return decks;
}

const DECKS = committedDecks();
const MEASURED = inventory(DECKS);

const key = (row: { dimension: string; form: string }): string => row.dimension + '=' + row.form;
const declared = new Map<string, LexicalForm>(LEXICAL_FORMS.map((row) => [key(row), row]));
const measured = new Map<string, FormEvidence>(MEASURED.map((row) => [key(row), row]));

describe('the corpus is a corpus of producers', () => {
  it('reads every committed deck of all three collections', () => {
    expect(DECKS).toHaveLength(52);
    expect(new Set(DECKS.map((deck) => deck.collection))).toEqual(
      new Set(['decks', 'authored', 'written']),
    );
  });

  it('names only serializers this table knows, so a typo cannot invent a producer', () => {
    const known = new Set<string>(Object.values(SERIALIZER));
    for (const deck of DECKS) {
      expect(known, deck.id + ' xml').toContain(deck.serializers.xml);
      expect(known, deck.id + ' container').toContain(deck.serializers.container);
    }
  });

  it('counts Tier B and Tier C as one XML producer and two container writers', () => {
    // The caveat, enforced rather than written down and trusted. `c01` is
    // `b01-blank` passed through `PartStore`: its ZIP headers are ours and
    // every byte inside every part is still Microsoft's, so counting files
    // instead of serializers would score a copy as independent evidence.
    const authored = DECKS.find((deck) => deck.collection === 'authored')!;
    const written = DECKS.find((deck) => deck.collection === 'written')!;

    expect(written.serializers.xml).toBe(authored.serializers.xml);
    expect(written.serializers.container).not.toBe(authored.serializers.container);
  });

  it('has three container writers and two XML serializers, and that is the ceiling', () => {
    // Everything below is bounded by these two numbers: no XML form can reach
    // three producers, and no `C-LEX` gap closes without a deck from somewhere
    // this corpus does not yet have.
    expect(new Set(DECKS.map((deck) => deck.serializers.xml)).size).toBe(2);
    expect(new Set(DECKS.map((deck) => deck.serializers.container)).size).toBe(3);
  });
});

describe('C-LEX', () => {
  it('catalogues every form in the corpus', () => {
    const uncatalogued = MEASURED.filter((row) => !declared.has(key(row))).map(
      (row) => key(row) + '  (' + row.producers.join(' + ') + ', ' + row.decks.join(', ') + ')',
    );

    // The `C008` of lexical form: a deck may not introduce a convention that
    // nobody wrote down. Adding the row is the work - saying what the form is,
    // whether it occurs in the wild, and who else writes it.
    expect(uncatalogued).toEqual([]);
  });

  it('claims exactly the producers it has, in both directions', () => {
    const wrong: string[] = [];
    for (const row of LEXICAL_FORMS) {
      const found = measured.get(key(row))?.producers ?? [];
      const claimed = [...row.producers].sort((a, b) => a.localeCompare(b));
      if (JSON.stringify(found) !== JSON.stringify(claimed)) {
        wrong.push(
          key(row) +
            '\n    declared ' +
            claimed.join(' + ') +
            '\n    measured ' +
            found.join(' + '),
        );
      }
    }
    expect(wrong).toEqual([]);
  });

  it('has every covered form in at least two decks, not just two producers', () => {
    // Both halves of the rule as written. Two producers and one deck would
    // mean a form that exists in exactly one file, which is a fixture rather
    // than a convention.
    const thin = LEXICAL_FORMS.filter((row) => row.producers.length >= 2)
      .map((row) => ({ row, decks: measured.get(key(row))?.decks ?? [] }))
      .filter((each) => each.decks.length < 2)
      .map((each) => key(each.row) + ' is in ' + String(each.decks.length) + ' deck(s)');

    expect(thin).toEqual([]);
  });

  it('says in prose why every uncovered form is uncovered', () => {
    const silent = LEXICAL_FORMS.filter(
      (row) => row.producers.length < 2 && (row.gap ?? '').trim() === '',
    ).map((row) => key(row));

    // A gap that is written down is a different thing from one nobody noticed.
    // This is `C-COV`'s `uncovered` array in another shape: a signed statement
    // that somebody looked, rather than a number quietly short of the target.
    expect(silent).toEqual([]);
  });

  it('gives every row evidence about the world outside this repository', () => {
    const unsupported = LEXICAL_FORMS.filter((row) => row.wild.trim() === '').map((row) =>
      key(row),
    );

    // Two producers that agree prove nothing on their own, which is why the
    // rule is not "two distinct producers" alone. Every row cites either a
    // measurement of real files or the counter-measurement saying the form was
    // looked for and not found - `a36`'s single-quoted attribute is 0 of
    // 170 019, and admitting that is the point of the field.
    expect(unsupported).toEqual([]);
  });

  it('has no row twice', () => {
    expect(declared.size).toBe(LEXICAL_FORMS.length);
  });
});

describe('what C-LEX found', () => {
  const covered = LEXICAL_FORMS.filter((row) => row.producers.length >= 2);
  const single = LEXICAL_FORMS.filter((row) => row.producers.length === 1);
  const absent = LEXICAL_FORMS.filter((row) => row.producers.length === 0);

  it('covers the conventions both producers share', () => {
    expect(covered.length).toBeGreaterThanOrEqual(29);
  });

  it('leaves every entity reference on one producer, which is the headline', () => {
    // The nine PowerPoint-authored decks contain no ampersand at all across
    // their 387 XML parts, so every escape in this corpus was written by the
    // escaper it is meant to test. E8 measured PowerPoint writing `&amp;`,
    // `&lt;`, `&gt;` and `&quot;`; the decks it measured were never committed.
    const entities = LEXICAL_FORMS.filter(
      (row) => row.dimension === 'xml.entity' && row.producers.length > 0,
    );
    expect(entities).not.toHaveLength(0);
    for (const row of entities) expect(row.producers).toEqual([SERIALIZER.GEN]);
  });

  it('leaves the dominant self-closing form on one producer, and it is ours', () => {
    const spaced = declared.get('xml.selfClosing=spaced')!;
    expect(spaced.producers).toEqual([SERIALIZER.GEN]);
    expect(measured.get('xml.selfClosing=spaced')?.decks).toEqual(['a36-spaced-tags']);
  });

  it('records the four parts PowerPoint writes with no declaration at all', () => {
    // Found by building this rule, and contradicting nothing that was written
    // down: E8 measured 38 parts and ADR 0004 measured 2834, and no part in
    // either lacked a declaration. `b05-chart`'s chart style and colour parts
    // begin at their root element.
    expect(measured.get('xml.declaration=absent')?.decks).toEqual(['b05-chart']);
    expect(declared.get('xml.declaration=absent')?.producers).toEqual([SERIALIZER.PPT]);
  });

  it('accounts for every form it cannot cover', () => {
    // Not a number to drive to zero. Some of these close with one more deck;
    // some cannot close at all, because a third XML producer would have to be
    // a file this project has no licence to commit.
    expect(single.length + absent.length).toBe(LEXICAL_FORMS.length - covered.length);
    for (const row of [...single, ...absent]) expect(row.gap).toBeTruthy();
  });
});

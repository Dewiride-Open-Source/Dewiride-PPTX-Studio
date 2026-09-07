import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { isRelationshipPartName, readZip, REL_TYPE, type PartStore } from '@pptx-studio/opc';
import { isValidateError } from '@pptx-studio/validate';
import { exportPackage, openPackage } from '@pptx-studio/writer';
import { describe, expect, it } from 'vitest';

/**
 * Sub-phase 1.3's verification: a no-op export of every committed deck.
 *
 * The plan words it as one line and it carries the whole of Phase 1. If a deck
 * can be opened and written back with nothing re-serialised, then the first
 * architectural bet holds - preservation is the *default state* of the export
 * path rather than a feature with a coverage percentage - and everything built
 * on top of it inherits that. If it does not hold, the architecture is wrong,
 * and the point of putting this before a single renderer line is to find out
 * now.
 *
 * ## What "no-op" is allowed to mean
 *
 * Not raw ZIP byte equality. The plan rules that out for 1.4 and the reason
 * applies here: `passthroughEntry` deliberately does not copy an entry's DOS
 * timestamp, external attributes or extra fields, because those describe where
 * a file came from rather than what is in it. An output that differed in a
 * modification date would fail a byte-equality test on day one, and a test that
 * is red on day one is disabled on day two.
 *
 * What is asserted instead is stronger where it matters and is checked by
 * `assertPreserved` inside every export here: for each entry, the **stored**
 * bytes - still compressed - together with the method, the CRC and the
 * uncompressed size. Comparing compressed bytes catches a re-compression that
 * an inflate-and-compare would wave through, which is the failure mode this is
 * looking for.
 *
 * The one number this file exists to produce is `preservation.checked`. It is
 * the count of entries compared and found identical, and summing it across the
 * corpus says how much of Phase 1's claim has actually been tested.
 */

const COLLECTIONS = ['decks', 'authored', 'written'];

interface Deck {
  readonly id: string;
  readonly bytes: Uint8Array;
  readonly entries: number;
}

function committedDecks(): Deck[] {
  const decks: Deck[] = [];
  for (const collection of COLLECTIONS) {
    const dir = join(ROOT, 'corpus', collection);
    const envelope = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as {
      entries: { id: string; kind: string; storage: string; path?: string }[];
    };
    for (const entry of envelope.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      const bytes = new Uint8Array(readFileSync(join(dir, entry.path ?? '')));
      decks.push({ id: entry.id, bytes, entries: readZip(bytes).entries.length });
    }
  }
  return decks;
}

const DECKS = committedDecks();

describe('a no-op export of every committed deck', () => {
  const results = DECKS.map((deck) => ({
    deck,
    result: exportPackage(openPackage(deck.bytes)),
  }));

  it('exports all fifty-four without the firewall refusing', () => {
    expect(results).toHaveLength(54);
    for (const { deck, result } of results) {
      expect(result.report?.ok, deck.id).toBe(true);
      expect(result.report?.blocking, deck.id).toBe(0);
    }
  });

  it('re-serialises not one part of not one deck', () => {
    // Reported as a list rather than a count, because when this fails the only
    // useful information is which part of which deck stopped being streamed.
    const rewritten = results.flatMap(({ deck, result }) =>
      result.rewritten.map((part) => deck.id + ' ' + part),
    );
    expect(rewritten).toEqual([]);
  });

  it('compares all 1570 entries of all 54 archives and finds them identical', () => {
    // 1474 parts plus the content-type stream of each deck. The export threw if
    // any single one differed; this asserts that none was quietly *skipped*,
    // which is the failure mode a passing preservation check can hide - and the
    // total is pinned rather than derived, so that a deck losing half its parts
    // is a failure here rather than a smaller number nobody reads.
    for (const { deck, result } of results) {
      expect(result.preservation.skipped, deck.id).toBeNull();
      expect(result.preservation.checked, deck.id).toBe(deck.entries);
    }
    const total = results.reduce((sum, { result }) => sum + result.preservation.checked, 0);
    expect(total).toBe(1570);
  });

  it('collects nothing and keeps nothing back', () => {
    for (const { deck, result } of results) {
      expect(result.collection.collect, deck.id).toEqual([]);
      expect(result.collection.kept, deck.id).toEqual([]);
      expect(result.collection.safe, deck.id).toBe(true);
    }
  });

  it('runs all twenty-nine rules, because a baseline was supplied', () => {
    // The contrast with `validate.test.ts`, which runs twenty-six: a file on
    // the command line has no history, so the three preservation rules are
    // skipped there and named as skipped. Here there is a baseline, and they
    // run against real markup - which is the first time in the project that
    // they have.
    for (const { deck, result } of results) {
      expect(result.report?.checked, deck.id).toHaveLength(29);
      expect(result.report?.skipped, deck.id).toEqual([]);
    }
  });

  it('produces an archive with the same entries in the same order', () => {
    for (const { deck, result } of results) {
      const before = readZip(deck.bytes).entries.map((entry) => entry.name);
      const after = readZip(result.bytes).entries.map((entry) => entry.name);
      expect(after, deck.id).toEqual(before);
    }
  });
});

describe('an export with a real edit in it', () => {
  /** The first slide part of a deck, which every deck has. */
  function firstSlide(store: PartStore): string {
    const slide = store.partNames.find((name) => /^\/ppt\/slides\/slide\d+\.xml$/.test(name));
    expect(slide).toBeDefined();
    return slide!;
  }

  it('re-serialises the edited part and streams the other twenty-something', () => {
    // The edit is a byte-for-byte replacement, which sounds like a no-op and is
    // not: `replacePart` clears `fromArchive`, so the part leaves the streaming
    // path and the three preservation rules switch on for it. That makes this
    // the cheapest way to run `V027`, `V028` and `V029` over real slide markup -
    // `mc:AlternateContent` branches, `extLst` blocks, `a:fld` caches and all -
    // and assert they find nothing, which until now nothing had done.
    const deck = DECKS.find((entry) => entry.id === 'a21-charts') ?? DECKS[0]!;
    const pkg = openPackage(deck.bytes);
    const slide = firstSlide(pkg.store);
    pkg.store.replacePart(slide, pkg.store.read(slide));

    const result = exportPackage(pkg);
    expect(result.rewritten).toEqual([slide]);
    expect(result.streamed).toBe(pkg.store.partNames.length - 1);
    expect(result.report?.ok).toBe(true);
    expect(result.preservation.checked).toBe(deck.entries - 1);
  });

  it('has no deck that writes xml:space="preserve" - so V029 has to be given one', () => {
    // Recorded rather than worked around. Sub-phase 1.2 measured 120 `<a:t>`
    // elements across the corpus carrying leading or trailing whitespace
    // *without* the attribute, which is why `V029` treats it as preserved
    // rather than required. The other half of that measurement is here: not one
    // deck writes it either, so nothing in the corpus can exercise the rule and
    // the next test has to build the file it needs.
    const writers = DECKS.filter((deck) => findPreservedSpace(deck) !== null).map(
      (deck) => deck.id,
    );
    expect(writers).toEqual([]);
  });

  it('refuses when the edit drops an xml:space it was given', () => {
    // A negative test, because "the rules found nothing" is only worth
    // something if they can find something. The baseline is a real deck with
    // the attribute added to one `<a:t>` - always valid, and what PowerPoint
    // itself writes for a run with edge whitespace. Taking it away again leaves
    // the XML well-formed and the deck openable, and changes what the text is:
    // the exact class of silent loss the rule exists for.
    const deck = DECKS.find((entry) => entry.id === 'a01-minimal') ?? DECKS[0]!;
    const seeded = openPackage(deck.bytes);
    const slide = firstSlide(seeded.store);
    const original = new TextDecoder().decode(seeded.store.read(slide));
    expect(original, deck.id + ' has no <a:t> to seed').toContain('<a:t>');
    seeded.store.replacePart(
      slide,
      new TextEncoder().encode(original.replace('<a:t>', '<a:t xml:space="preserve">')),
    );
    const withSpace = exportPackage({ ...seeded, verifyPreservation: false }).bytes;

    const pkg = openPackage(withSpace);
    pkg.store.replacePart(
      slide,
      new TextEncoder().encode(
        new TextDecoder().decode(pkg.store.read(slide)).replace(' xml:space="preserve"', ''),
      ),
    );

    try {
      exportPackage(pkg);
      expect.unreachable('should have refused: ' + deck.id + ' ' + slide);
    } catch (error) {
      expect(isValidateError(error) && error.code).toBe('ERR_VALIDATION_FAILED');
      expect((error as Error).message).toContain('V029');
    }
  });
});

/**
 * Deleting a picture, on a deck that opens in PowerPoint.
 *
 * `b09-picture` has two `<p:pic>` on its one slide; `rId3` names the second and
 * nothing else in the package relates to `image2.jpeg`. Pinned by name and by
 * anchor rather than searched for, the way the reject fixtures are: an edit
 * that silently stops matching is an edit that silently stops testing anything.
 */
describe('an export that orphans an image', () => {
  const DECK = 'b09-picture';
  const SLIDE = '/ppt/slides/slide1.xml';
  const RELS = '/ppt/slides/_rels/slide1.xml.rels';
  const REL_ID = 'rId3';
  const IMAGE = '/ppt/media/image2.jpeg';

  function deckBytes(): Uint8Array {
    const deck = DECKS.find((entry) => entry.id === DECK);
    expect(deck, DECK + ' is no longer in the corpus').toBeDefined();
    return deck!.bytes;
  }

  it('refuses when only the relationship is removed', () => {
    // The division of labour, made falsifiable. The sweep is a safety net, not
    // a delete: on its own, dropping the relationship leaves eight bytes of
    // markup still naming it, and that is a dangling reference rather than a
    // tidier file. `V006` catches it, and the firewall would rather refuse than
    // hand back a deck PowerPoint opens with a red X where a photograph was.
    const pkg = openPackage(deckBytes());
    const rels = pkg.store.relationships(SLIDE);
    expect(rels.byId(REL_ID)?.type).toBe(REL_TYPE.image);
    rels.remove(REL_ID);

    try {
      exportPackage(pkg);
      expect.unreachable('should have refused');
    } catch (error) {
      expect(isValidateError(error) && error.code).toBe('ERR_VALIDATION_FAILED');
      expect((error as Error).message).toContain('V006');
    }
  });

  it('collects the image once the markup that used it is gone', () => {
    const bytes = deckBytes();
    const pkg = openPackage(bytes);
    const before = new TextDecoder().decode(pkg.store.read(SLIDE));
    const after = removePicture(before, REL_ID);
    expect(after, 'no <p:pic> in ' + SLIDE + ' uses ' + REL_ID).not.toBe(before);

    pkg.store.replacePart(SLIDE, new TextEncoder().encode(after));
    pkg.store.relationships(SLIDE).remove(REL_ID);
    const result = exportPackage(pkg);

    expect(result.report?.ok).toBe(true);
    expect(result.collection.collect).toEqual([IMAGE]);
    expect(result.collection.kept).toEqual([]);
    // Two parts re-serialised: the slide we edited and its relationship part.
    // Everything else - the other picture's image included - streamed.
    expect([...result.rewritten].sort()).toEqual([RELS, SLIDE]);
    expect(result.preservation.checked).toBe(readZip(bytes).entries.length - 3);

    // The other picture keeps its image: a sweep that took both would be the
    // shared-media bug, and it is the one that eats a logo off every slide.
    const mediaOf = (archive: Uint8Array): string[] =>
      readZip(archive)
        .entries.map((entry) => entry.name)
        .filter((name) => /^ppt\/media\//i.test(name));
    expect(mediaOf(result.bytes)).toEqual(
      mediaOf(bytes).filter((name) => name !== 'ppt/media/image2.jpeg'),
    );
    expect(mediaOf(result.bytes).length).toBeGreaterThan(0);
  });
});

/**
 * Cut out the `<p:pic>` that uses one relationship.
 *
 * Index arithmetic rather than a regular expression, because a lazy
 * `<p:pic>[\s\S]*?</p:pic>` matches from the *first* picture to the *first*
 * close tag after the id, which on a slide with two pictures silently deletes
 * both. `p:pic` cannot nest, so scanning tag to tag is exact.
 */
function removePicture(xml: string, relId: string): string {
  const embed = 'r:embed="' + relId + '"';
  for (let at = 0; ;) {
    const start = xml.indexOf('<p:pic>', at);
    if (start < 0) return xml;
    const end = xml.indexOf('</p:pic>', start) + '</p:pic>'.length;
    if (xml.slice(start, end).includes(embed)) return xml.slice(0, start) + xml.slice(end);
    at = end;
  }
}

/** The first part in a deck whose markup declares `xml:space="preserve"`. */
function findPreservedSpace(deck: Deck): { part: string } | null {
  const store = openPackage(deck.bytes).store;
  for (const part of store.partNames) {
    if (!part.endsWith('.xml') || isRelationshipPartName(part)) continue;
    const text = new TextDecoder().decode(store.read(part));
    if (text.includes(' xml:space="preserve"')) return { part };
  }
  return null;
}

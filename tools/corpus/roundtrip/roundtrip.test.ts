import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { REPO_ROOT as ROOT } from '../../repo/root.ts';
import { isRelationshipPartName, isXmlContentType, PartStore } from '@pptx-studio/opc';
import { attributeNamespaceOf, descendantElements, NS, parseXml } from '@pptx-studio/xml';
import { comparePackages, roundTripPackage } from '@pptx-studio/writer';
import { describe, expect, it } from 'vitest';

/**
 * Sub-phase 1.4's verification: read every committed deck, write it back, and
 * prove nothing moved.
 *
 * The plan writes it as "50/50 green" and spends most of its sentence on what
 * the test must **not** be:
 *
 * > Never raw ZIP byte equality - entry order, deflate level, DOS timestamps
 * > and attribute order all legitimately differ, so a byte-equality test is red
 * > on day one and disabled on day two.
 *
 * Sub-phase 1.3 already compares stored bytes, and that comparison is worth
 * having: on a no-op export nothing should be re-serialised at all, so an
 * entry whose bytes moved is a defect. But it can only ever answer for an
 * export that changed nothing. The moment a real edit lands - the moment this
 * project does anything - the question becomes whether the deck is the same
 * *document*, and no comparison of archives can answer that.
 *
 * So this file exercises the comparator from three directions:
 *
 * 1. It agrees, on all fifty-four decks.
 * 2. It keeps agreeing when the archive is deliberately rebuilt differently -
 *    a normalised entry order and a different deflate level - which is where a
 *    byte-equality gate would be red.
 * 3. It keeps agreeing when every relationship id in a deck is renumbered and
 *    the markup follows, which is what PowerPoint does on save and the one
 *    clause of the plan that needed real markup to be believed.
 *
 * And, because a comparator that never disagrees is not a comparator, three
 * negatives: one character of a slide, one byte of an image, and one
 * relationship removed.
 */

const COLLECTIONS = ['decks', 'authored', 'written'];

interface Deck {
  readonly id: string;
  readonly bytes: Uint8Array;
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
      decks.push({
        id: entry.id,
        bytes: new Uint8Array(readFileSync(join(dir, entry.path ?? ''))),
      });
    }
  }
  return decks;
}

const DECKS = committedDecks();
const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);
const encode = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('a round trip of every committed deck', () => {
  const results = DECKS.map((deck) => ({ deck, result: roundTripPackage(deck.bytes) }));

  it('reads and writes all fifty-four with nothing to report', () => {
    expect(results).toHaveLength(54);
    const failures = results
      .filter(({ result }) => !result.ok)
      .map(({ deck, result }) => deck.id + ': ' + JSON.stringify(result.comparison.differences));
    expect(failures).toEqual([]);
  });

  it('compares every part of every deck, and says how it compared each', () => {
    // Pinned rather than derived, for the same reason 1.3 pins its entry count:
    // a deck that silently lost half its parts would still report zero
    // differences, and only a total can catch that.
    const totals = results.reduce(
      (sum, { result }) => ({
        xml: sum.xml + result.comparison.counts.xml,
        binary: sum.binary + result.comparison.counts.binary,
        relationships: sum.relationships + result.comparison.counts.relationships,
        same: sum.same + result.comparison.counts.same,
      }),
      { xml: 0, binary: 0, relationships: 0, same: 0 },
    );
    const parts = results.reduce((sum, { result }) => sum + result.comparison.parts.length, 0);

    expect(totals.xml + totals.binary + totals.relationships).toBe(parts);
    expect(totals.same).toBe(parts);
    expect(parts).toBe(1516);

    // 553 relationship parts, 46 compared as bytes, and the remaining 875 as
    // canonical XML. The 46 is wider than the 23 parts under `/ppt/media/` that
    // the media sweep can touch: it also counts five `.fntdata`, five embedded
    // workbooks, eleven `docProps/thumbnail.jpeg` and two `vbaProject.bin`.
    //
    // Counted, not recited. The previous spelling of this comment claimed
    // twenty-four media parts against a total of forty-three, which does not
    // add up - and nothing in the suite would ever have said so, because the
    // breakdown was prose beside three numbers rather than a fourth number.
    //
    // What is *not* in it any more is the three `vmlDrawing` parts. VML is XML
    // whose content type does not say so - no `+xml` anywhere in it - and the
    // spelling Office writes has a capital D in the middle. This split is what
    // caught a lookup table that had copied that spelling and then lower-cased
    // its input, which matched nothing; before the fix these three compared as
    // opaque bytes, which is right about the answer and wrong about everything
    // else, including where a difference would be reported.
    expect(totals.relationships).toBe(569);
    expect(totals.binary).toBe(46);
    expect(totals.xml).toBe(901);
  });

  it('renames no relationship id, because this writer never renumbers', () => {
    // The mechanism is there for a file PowerPoint saved, not for one of ours.
    // Asserting it stays unused is what makes the corpus test below - where the
    // renumbering is done on purpose - mean something.
    for (const { deck, result } of results) {
      expect(result.comparison.relabelled, deck.id).toEqual([]);
    }
  });

  it('agrees when every part is re-compressed and the entries reordered', () => {
    // The claim the plan makes about byte equality, made falsifiable on real
    // decks. A no-op export cannot show it, and that is worth knowing rather
    // than working around: nothing is re-serialised, so `deflateLevel` has
    // nothing to act on and the archive comes out identical. Marking every part
    // rewritten is what gives the two knobs something to change - every entry
    // deflated at a different level, in a different order - and the document is
    // the same document throughout.
    for (const deck of DECKS) {
      const result = roundTripPackage(deck.bytes, {
        normalizeEntryOrder: true,
        deflateLevel: 1,
        verifyPreservation: false,
        prepare: [
          {
            name: 'rewrite every part with its own bytes',
            run: (context) => {
              for (const part of context.store.partNames) {
                // Except the embeddings, and not because of anything to do with
                // this comparison: `V027` refuses to let an export replace one
                // at all, whatever the replacement is, because `ppt/embeddings`
                // holds OLE2 compound files and rewriting one is a guaranteed
                // repair prompt. The firewall stopped this test the first time
                // it ran, which is the rule working.
                if (part.startsWith('/ppt/embeddings/')) continue;
                context.store.replacePart(part, context.store.read(part));
              }
            },
          },
        ],
      });
      expect(result.comparison.differences, deck.id).toEqual([]);
      expect(result.exported.rewritten.length, deck.id).toBeGreaterThan(0);
      expect(result.exported.bytes, deck.id).not.toEqual(deck.bytes);
    }
    // Fifty-one decks re-compressed part by part, so a minute rather than the
    // default five seconds: a timeout tuned to a quiet machine is a test that
    // fails on a busy one.
  }, 60_000);

  it('agrees when the attributes of real markup are written in another order', () => {
    // Attribute order is the difference the plan names first, and the one a
    // byte comparison is least able to forgive. `<a:off>` carries `x` and `y`
    // in that order in every deck that has one; swapping them changes the part
    // and changes nothing about the slide.
    let swapped = 0;
    for (const deck of DECKS) {
      const store = PartStore.open(deck.bytes);
      for (const part of store.partNames) {
        if (!/^\/ppt\/(slides|slideLayouts|slideMasters)\//.test(part)) continue;
        const markup = decode(store.read(part));
        const next = markup.replace(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g, '<a:off y="$2" x="$1"/>');
        if (next === markup) continue;
        store.replacePart(part, encode(next));
        swapped++;
      }
      const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(store.write()));
      expect(report.differences, deck.id).toEqual([]);
    }
    expect(swapped).toBeGreaterThan(40);
  }, 60_000);
});

/**
 * Renumber every relationship id in a package, in the `.rels` and in the markup
 * that refers to them.
 *
 * `+ 1000` rather than a shuffle, so that no new id collides with an old one
 * inside the same collection - which would make the package say something
 * different rather than say the same thing differently, and would be testing
 * the wrong claim.
 *
 * The markup pass is anchored on a `r:`-prefixed attribute name. Nothing else
 * in an OOXML part holds a relationship id, and a looser pattern would rewrite
 * any string that happened to look like one.
 */
function renumber(bytes: Uint8Array): { bytes: Uint8Array; changed: number } {
  const store = PartStore.open(bytes);
  let changed = 0;
  for (const part of store.partNames) {
    const text = decode(store.read(part));
    const next = isRelationshipPartName(part)
      ? text.replace(
          /Id="rId(\d+)"/g,
          (_all, id: string) => 'Id="rId' + String(Number(id) + 1000) + '"',
        )
      : text.replace(
          /(\sr:[A-Za-z]+=")rId(\d+)"/g,
          (_all, lead: string, id: string) => lead + 'rId' + String(Number(id) + 1000) + '"',
        );
    if (next === text) continue;
    changed++;
    store.replacePart(part, encode(next));
  }
  return { bytes: store.write(), changed };
}

describe('a deck whose relationship ids have all been renamed', () => {
  // Pinned to a deck with pictures, so the renumbering touches `r:embed` inside
  // slide markup and not only the `r:id` of the slide list. Falling back keeps
  // the test honest rather than silently skipping if the id ever moves.
  const deck = DECKS.find((entry) => entry.id === 'b09-picture') ?? DECKS[0]!;

  it('renames ids in the relationship parts and in the markup that uses them', () => {
    const { bytes, changed } = renumber(deck.bytes);
    expect(changed).toBeGreaterThan(2);
    const store = PartStore.open(bytes);
    const rels = store.partNames.filter((name) => isRelationshipPartName(name));
    const ids = rels.flatMap((part) => [...decode(store.read(part)).matchAll(/Id="(rId\d+)"/g)]);
    expect(ids.length).toBeGreaterThan(0);
    expect(ids.every(([, id]) => Number(id!.slice(3)) > 1000)).toBe(true);
  });

  it('is still the same deck', () => {
    // The plan's "relationship-graph isomorphism with rIds treated as opaque
    // labels", on a real package rather than a fixture. Every id in the file
    // changed; nothing a reader can observe did.
    const { bytes } = renumber(deck.bytes);
    const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(bytes));
    expect(report.differences).toEqual([]);
    expect(report.relabelled.length).toBeGreaterThan(2);
    for (const entry of report.relabelled) {
      expect(Number(entry.to.slice(3))).toBe(Number(entry.from.slice(3)) + 1000);
    }
  });

  it('is not the same deck once one of those ids points somewhere else', () => {
    // The other half. A renumbering is invisible; a re-pointing is not, and the
    // difference between the two is the whole reason the ids are matched by
    // what they resolve to rather than assumed to line up.
    const store = PartStore.open(deck.bytes);
    const slide = '/ppt/slides/slide1.xml';
    const markup = decode(store.read(slide));
    const embeds = [...markup.matchAll(/r:embed="(rId\d+)"/g)].map(([, id]) => id!);
    expect(new Set(embeds).size).toBeGreaterThan(1);

    store.replacePart(slide, encode(markup.replace(embeds[0]!, embeds[1]!)));
    const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(store.write()));
    expect(report.differences.map((entry) => entry.kind)).toContain('xml');
  });
});

describe('what the comparison has to be able to see', () => {
  const deck = DECKS.find((entry) => entry.id === 'b09-picture') ?? DECKS[0]!;

  it('sees one character of a slide', () => {
    const store = PartStore.open(deck.bytes);
    const slide = '/ppt/slides/slide1.xml';
    const markup = decode(store.read(slide));
    // `<p:sp>` has a name attribute on every shape; changing one character of
    // it changes nothing a viewer would notice, which is the point - the
    // comparison is not allowed to depend on the change being visible.
    expect(markup).toContain('name="');
    store.replacePart(slide, encode(markup.replace('name="', 'name="x')));

    const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(store.write()));
    const difference = report.differences.find((entry) => entry.kind === 'xml');
    expect(difference?.part).toBe(slide);
    expect(difference?.detail).toContain('character');
  });

  it('sees one byte of an image, and names both digests', () => {
    const store = PartStore.open(deck.bytes);
    const image = store.partNames.find((name) => name.startsWith('/ppt/media/'));
    expect(image).toBeDefined();
    const pixels = Uint8Array.from(store.read(image!));
    pixels[pixels.length - 1] = (pixels[pixels.length - 1]! + 1) & 0xff;
    store.replacePart(image!, pixels);

    const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(store.write()));
    const difference = report.differences.find((entry) => entry.kind === 'binary');
    expect(difference?.part).toBe(image);
    const compared = report.parts.find((part) => part.part === image);
    expect(compared?.how).toBe('binary');
    expect(compared?.digest).toMatch(/^[0-9a-f]{64}$/);
    expect(compared?.writtenDigest).not.toBe(compared?.digest);
  });

  it('sees a relationship that is gone', () => {
    const store = PartStore.open(deck.bytes);
    const rels = '/ppt/slides/_rels/slide1.xml.rels';
    const markup = decode(store.read(rels));
    const stripped = markup.replace(/<Relationship [^>]*Id="rId3"[^>]*\/>/, '');
    expect(stripped).not.toBe(markup);
    store.replacePart(rels, encode(stripped));

    const report = comparePackages(PartStore.open(deck.bytes), PartStore.open(store.write()));
    const difference = report.differences.find((entry) => entry.kind === 'relationship');
    expect(difference?.part).toBe('/ppt/slides/slide1.xml');
    expect(difference?.detail).toContain('rId3');
  });
});

describe('the premise the relabelling rests on', () => {
  /**
   * Every attribute in the relationships namespace holds a relationship id.
   *
   * The comparator relabels by **namespace** rather than by a list of attribute
   * names, and this is the assertion that says why. The plan names eight:
   * `r:id`, `r:embed`, `r:link`, `r:pict`, `r:dm`, `r:lo`, `r:qs`, `r:cs`. The
   * corpus has eight too - and they are not the same eight. `r:blip` is here, on
   * forty-eight attributes, and `r:pict` is not.
   *
   * A list of names would therefore already have been one entry short, on a
   * corpus we built ourselves, before meeting a single deck from outside. That
   * is the whole argument, and it is measured rather than asserted.
   */
  it('has no attribute in the relationships namespace that is not an rId', () => {
    const names = new Map<string, number>();
    const odd: string[] = [];

    for (const deck of DECKS) {
      const store = PartStore.open(deck.bytes);
      for (const part of store.partNames) {
        if (isRelationshipPartName(part)) continue;
        if (!isXmlContentType(store.contentTypeOf(part))) continue;
        for (const element of descendantElements(parseXml(store.read(part)).root)) {
          for (const attr of element.attributes) {
            if (attributeNamespaceOf(element, attr) !== NS.r) continue;
            names.set(attr.local, (names.get(attr.local) ?? 0) + 1);
            // Empty is legal: `a:blip` with neither an embed nor a link is a
            // picture placeholder with no picture in it yet.
            if (!/^(rId\d+)?$/.test(attr.value)) odd.push(part + ' ' + attr.qname);
          }
        }
      }
    }

    expect(odd).toEqual([]);
    expect([...names.keys()].sort()).toEqual([
      'blip',
      'cs',
      'dm',
      'embed',
      'id',
      'link',
      'lo',
      'qs',
    ]);
    expect(names.get('blip')).toBe(48);
  });
});

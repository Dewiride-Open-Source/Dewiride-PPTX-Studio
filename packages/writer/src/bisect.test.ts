import {
  CONTENT_TYPE,
  PartStore,
  REL_TYPE,
  deflatedEntry,
  passthroughEntry,
  readZip,
  writeZip,
  type ZipEntryInput,
} from '@pptx-studio/opc';
import { parseXml } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';
import {
  bisectPackages,
  collectDelta,
  flattenChanges,
  leafChanges,
  type Verdict,
} from './bisect.js';

/**
 * A bisector is judged on two things, and only one of them is that it finds the
 * answer.
 *
 * The other is what it does when it cannot. A reducer that always returns
 * *something* is worse than useless, because the something looks exactly like a
 * finding: hand it a package that was already broken and it will confidently
 * blame an innocent change. So the outcomes that are not "localized" get as
 * many tests here as the one that is.
 *
 * Every reduction below is run against a delta of **many** changes, which is
 * not incidental. A pair of packages differing in one place reduces in zero
 * oracle runs - `ddmin` returns immediately on a set of one - so a test built
 * that way asserts the answer without ever exercising the search. The first
 * draft of this file was built that way, and four of its tests passed for that
 * reason and no other.
 *
 * The oracle here is synchronous and deterministic - "does this text appear in
 * this entry" - because what is under test is the search, not the question. The
 * real oracles live in `packages/cli`, where there is a process to spawn and a
 * PowerPoint to ask.
 */

const enc = (text: string): Uint8Array => new TextEncoder().encode(text);
const dec = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

const MAIN = '/doc.xml';
const MAIN_ENTRY = 'doc.xml';

/** A package holding one XML part, plus whatever extra entries are asked for. */
function pack(body: string, extra: Record<string, Uint8Array> = {}): Uint8Array {
  const store = PartStore.create();
  store.addPart(MAIN, CONTENT_TYPE.xml, enc(body));
  store.rootRelationships().addTo(REL_TYPE.officeDocument, MAIN);
  for (const [name, bytes] of Object.entries(extra)) {
    store.addPart(name, CONTENT_TYPE.png, bytes);
  }
  return store.write();
}

/** Rewrite one ZIP entry, leaving every other entry byte-identical. */
function rewriteEntry(bytes: Uint8Array, name: string, to: (was: string) => string): Uint8Array {
  const archive = readZip(bytes);
  const entries: ZipEntryInput[] = archive.entries.map((entry) =>
    entry.name === name
      ? deflatedEntry(entry.name, enc(to(dec(archive.read(entry)))))
      : passthroughEntry(archive, entry),
  );
  return writeZip(entries);
}

/** The text of an entry inside a package. */
function entryText(bytes: Uint8Array, name: string): string {
  const archive = readZip(bytes);
  const entry = archive.get(name);
  return entry === undefined ? '' : dec(archive.read(entry));
}

/** `count` shapes, each replaceable by `edit`. */
function deck(count: number, edit: (i: number) => string | undefined = () => undefined): string {
  const parts: string[] = [];
  for (let i = 0; i < count; i += 1) {
    parts.push(edit(i) ?? '<sp name="shape' + String(i) + '" x="' + String(i * 10) + '"/>');
  }
  return '<doc>' + parts.join('') + '</doc>';
}

/**
 * Every shape edited, and exactly one of them poisoned.
 *
 * The shape of a real export gone wrong: many changes, one of them fatal. The
 * poisoned shape keeps its `x`, so the only difference that can matter is the
 * name - which is what makes the expected answer a single attribute.
 */
const poisonAmongMany =
  (at: number) =>
  (i: number): string =>
    i === at
      ? '<sp name="POISON" x="' + String(i * 10) + '"/>'
      : '<sp name="renamed' + String(i) + '" x="' + String(i * 10 + 1) + '"/>';

/** An oracle that fails when a string appears in an entry, and a log of what it saw. */
function poisonOracle(needle: string, entry = MAIN_ENTRY) {
  const seen: Uint8Array[] = [];
  return {
    seen,
    oracle: (bytes: Uint8Array): Verdict => {
      seen.push(bytes);
      return entryText(bytes, entry).includes(needle) ? 'fails' : 'passes';
    },
  };
}

describe('what the delta is taken over', () => {
  it('says two identical packages are identical, and asks the oracle nothing', () => {
    const bytes = pack('<doc><a/></doc>');
    const probe = poisonOracle('never');
    const result = bisectPackages(bytes, bytes, { oracle: probe.oracle });

    expect(result.outcome).toBe('identical');
    expect(result.changes).toEqual([]);
    expect(result.runs).toBe(0);
  });

  it('descends into every entry that parses as XML, and no further', () => {
    // A PNG is one atom because it is not markup; the XML part is a tree. The
    // rule is "did it parse", not "what does the content type say" - which is
    // what lets `[Content_Types].xml` and every `.rels` be descended into
    // without this file knowing that either of them exists.
    const original = pack('<doc><a x="1"/></doc>', { '/media/a.png': enc('the original pixels') });
    const broken = rewriteEntry(
      rewriteEntry(original, MAIN_ENTRY, () => '<doc><a x="2"/></doc>'),
      'media/a.png',
      () => 'quite different pixels',
    );

    const roots = collectDelta(readZip(original), readZip(broken)).changes;
    expect(roots).toHaveLength(2);

    const image = roots.find((change) => change.entry === 'media/a.png')!;
    expect(image.kind).toBe('entry');
    expect(image.children).toEqual([]);

    const doc = roots.find((change) => change.entry === MAIN_ENTRY)!;
    const leaves = flattenChanges([doc]).filter((change) => change.children.length === 0);
    // ...down to the attribute, which is the finest thing it can name.
    expect(leaves.map((change) => change.where)).toContain('/doc/a/@x');
  });

  it('reaches [Content_Types].xml, which is not a part and never will be', () => {
    // The reason the delta is over ZIP entries rather than over `PartStore`
    // parts. `partNames` excludes the content-type stream by design, so a
    // bisector built on parts could not vary the one file the plan names as the
    // canonical cause of a repair prompt.
    const original = pack('<doc><a/></doc>');
    const broken = rewriteEntry(original, '[Content_Types].xml', (was) =>
      was.replace('ContentType="application/xml"', 'ContentType="application/poisoned"'),
    );

    const probe = poisonOracle('application/poisoned', '[Content_Types].xml');
    const result = bisectPackages(original, broken, { oracle: probe.oracle });

    expect(result.outcome).toBe('localized');
    expect(result.minimal).toHaveLength(1);
    expect(result.minimal[0]!.entry).toBe('[Content_Types].xml');
    expect(result.minimal[0]!.kind).toBe('attribute');
    expect(result.minimal[0]!.where).toMatch(/@ContentType$/);
  });

  it('names the one child that was removed, not the parent that lost it', () => {
    // The case that forced the two lists to be matched from both ends. With
    // straight positional pairing a deletion shifts every child after it,
    // nothing lines up, and the answer degrades to "something in this element" -
    // which is what this reported for `[Content_Types].xml` on the first run
    // against the real PowerPoint, and it is the headline failure of the whole
    // sub-phase.
    const original = pack('<doc><a/><b/><c/><d/><e/></doc>');
    const broken = pack('<doc><a/><b/><d/><e/></doc>');

    const leaves = leafChanges(collectDelta(readZip(original), readZip(broken)).changes);
    expect(leaves).toHaveLength(1);
    expect(leaves[0]!.kind).toBe('children');
    expect(leaves[0]!.where).toBe('/doc/c');
    expect(leaves[0]!.was).toBe('<c/>');
    expect(leaves[0]!.text).toBe('');
  });

  it('names a child that was inserted, with nothing on the original side', () => {
    const original = pack('<doc><a/><b/></doc>');
    const broken = pack('<doc><a/><inserted/><b/></doc>');

    const leaves = leafChanges(collectDelta(readZip(original), readZip(broken)).changes);
    expect(leaves).toHaveLength(1);
    expect(leaves[0]!.kind).toBe('children');
    expect(leaves[0]!.was).toBe('');
    expect(leaves[0]!.text).toBe('<inserted/>');
  });

  it('still pairs one to one when the two sides are the same length', () => {
    // Trimming from both ends must not cost the finer answer in the ordinary
    // case: three children, the middle one edited, still resolves to the
    // attribute rather than to a run.
    const original = pack('<doc><a/><b v="1"/><c/></doc>');
    const broken = pack('<doc><a/><b v="2"/><c/></doc>');

    const leaves = leafChanges(collectDelta(readZip(original), readZip(broken)).changes);
    expect(leaves).toHaveLength(1);
    expect(leaves[0]!.kind).toBe('attribute');
    expect(leaves[0]!.where).toBe('/doc/b/@v');
  });

  it('names an entry that was added, and one that was taken away', () => {
    const original = pack('<doc/>', { '/media/gone.png': enc('pixels') });
    const broken = pack('<doc/>', { '/media/new.png': enc('other pixels') });

    const kinds = new Map(
      collectDelta(readZip(original), readZip(broken)).changes.map((c) => [c.entry, c.kind]),
    );
    expect(kinds.get('media/gone.png')).toBe('entry-removed');
    expect(kinds.get('media/new.png')).toBe('entry-added');
  });
});

describe('finding the one change that matters', () => {
  it('picks one poisoned shape out of thirty that were all edited', () => {
    const probe = poisonOracle('POISON');
    const result = bisectPackages(pack(deck(30)), pack(deck(30, poisonAmongMany(17))), {
      oracle: probe.oracle,
    });

    expect(result.outcome).toBe('localized');
    expect(result.minimal).toHaveLength(1);
    expect(result.minimal[0]!.entry).toBe(MAIN_ENTRY);
    expect(result.minimal[0]!.where).toBe('/doc/sp[18]/@name');
  });

  it('leaves the twenty-nine innocent changes out of the package it hands back', () => {
    // A bisector that returned all thirty would be a diff, and there is one of
    // those already. What makes this the answer is that the other twenty-nine
    // came back the way the original wrote them.
    const probe = poisonOracle('POISON');
    const result = bisectPackages(pack(deck(30)), pack(deck(30, poisonAmongMany(17))), {
      oracle: probe.oracle,
    });

    const text = entryText(result.bytes, MAIN_ENTRY);
    expect(text).toContain('POISON');
    expect(text).toContain('name="shape0"');
    expect(text).toContain('name="shape29"');
    expect(text).not.toContain('renamed');
    // ...and the poisoned shape kept the original's `x`, because only its name
    // was load-bearing. Reverting a whole element would have been an answer
    // too, and a coarser one.
    expect(text).toContain('<sp name="POISON" x="170"/>');
  });

  it('keeps both changes when neither is fatal on its own', () => {
    // 1-minimal is not the same as smallest, and this is where the difference
    // shows. The oracle fails only when both markers are present, so no single
    // change can be dropped and no single change is the answer.
    const original = deck(20);
    const broken = deck(20, (i) =>
      i === 4
        ? '<sp name="LEFT" x="40"/>'
        : i === 15
          ? '<sp name="RIGHT" x="150"/>'
          : '<sp name="renamed' + String(i) + '" x="' + String(i * 10 + 1) + '"/>',
    );

    const result = bisectPackages(pack(original), pack(broken), {
      oracle: (bytes) => {
        const text = entryText(bytes, MAIN_ENTRY);
        return text.includes('LEFT') && text.includes('RIGHT') ? 'fails' : 'passes';
      },
    });

    expect(result.outcome).toBe('localized');
    expect(result.minimal).toHaveLength(2);
    expect(result.minimal.map((change) => change.where).sort()).toEqual([
      '/doc/sp[16]/@name',
      '/doc/sp[5]/@name',
    ]);
  });

  it('localizes a change in a relationship part', () => {
    // `.rels` is XML, so it decomposes like anything else - which is what makes
    // "the one relationship that broke it" an answer this can give.
    const original = pack('<doc/>');
    const broken = rewriteEntry(original, '_rels/.rels', (was) =>
      was.replace('Id="rId1"', 'Id="rId9999"'),
    );

    const probe = poisonOracle('rId9999', '_rels/.rels');
    const result = bisectPackages(original, broken, { oracle: probe.oracle });

    expect(result.minimal).toHaveLength(1);
    expect(result.minimal[0]!.entry).toBe('_rels/.rels');
    expect(result.minimal[0]!.kind).toBe('attribute');
    expect(result.minimal[0]!.where).toMatch(/@Id$/);
  });
});

describe('what it does when there is nothing to find', () => {
  it('says so when the broken package is not broken', () => {
    const result = bisectPackages(pack(deck(10)), pack(deck(10, poisonAmongMany(3))), {
      oracle: () => 'passes',
    });

    expect(result.outcome).toBe('broken-passes');
    expect(result.minimal).toEqual(result.changes);
    // Two runs and no more: the empty configuration and the full one. There is
    // nothing to be gained by reducing a set whose full form does not fail.
    expect(result.runs).toBe(2);
  });

  it('says so when the original was already broken', () => {
    // The failure this outcome exists to prevent: with no check, `ddmin` finds
    // that the empty configuration fails, reduces to nothing, and reports that
    // no change at all was needed - phrased as though it had localized one.
    const result = bisectPackages(pack(deck(10)), pack(deck(10, poisonAmongMany(3))), {
      oracle: () => 'fails',
    });

    expect(result.outcome).toBe('original-fails');
    expect(result.minimal).toEqual([]);
    expect(result.runs).toBe(1);
  });

  it('stops at the ceiling and says it stopped, rather than reporting a minimum', () => {
    const probe = poisonOracle('POISON');
    const result = bisectPackages(pack(deck(40)), pack(deck(40, poisonAmongMany(21))), {
      oracle: probe.oracle,
      maxRuns: 4,
    });

    expect(result.exhausted).toBe(true);
    expect(result.runs).toBe(4);
    // Twenty, and the number is the point: the partial reduction is kept. Two
    // of the four runs go on the preconditions - the empty configuration and
    // the full one - and the other two halve forty shapes down to the twenty
    // holding the poisoned one, the first half having to be tried and rejected
    // before the second is tried and kept. The fifth run is the one the ceiling
    // refuses. An earlier version discarded that progress and reported the
    // whole entry instead, which is a true answer and a useless one.
    expect(result.minimal).toHaveLength(20);
    // Still a failing configuration, just not a minimal one - and the flag is
    // what stops a caller reading it as though it were.
    expect(entryText(result.bytes, MAIN_ENTRY)).toContain('POISON');
  });
});

describe('the properties the search depends on', () => {
  it('shows the oracle only well-formed XML, in every configuration', () => {
    // The reason for isolating rather than simplifying. Because the splice unit
    // is a whole node replaced by a whole node, no configuration can come out
    // unbalanced - so no oracle run is spent on markup no editor could produce.
    // One shape changes shape entirely, to make sure a structural difference
    // goes the same way.
    const broken = deck(20, (i) =>
      i === 11
        ? '<sp name="POISON"><t>and now it has children</t></sp>'
        : '<sp name="renamed' + String(i) + '" x="' + String(i * 10 + 1) + '"/>',
    );

    const probe = poisonOracle('POISON');
    bisectPackages(pack(deck(20)), pack(broken), { oracle: probe.oracle });

    expect(probe.seen.length).toBeGreaterThan(3);
    for (const bytes of probe.seen) {
      expect(() => parseXml(enc(entryText(bytes, MAIN_ENTRY)))).not.toThrow();
    }
  });

  it('rebuilds the original exactly when no change is applied', () => {
    // `ddmin`'s precondition is that the empty configuration passes, and the
    // empty configuration is this. If it were not the original, every verdict
    // after it would be about a package nobody asked about.
    const original = pack('<doc><a x="1"/><b>text</b></doc>');
    const broken = pack('<doc><a x="2"/><b>other</b></doc>');

    let first: Uint8Array | undefined;
    bisectPackages(original, broken, {
      oracle: (bytes) => {
        first ??= bytes;
        return 'passes';
      },
    });

    expect(entryText(first!, MAIN_ENTRY)).toBe('<doc><a x="1"/><b>text</b></doc>');
  });

  it('never asks the oracle the same question twice', () => {
    const probe = poisonOracle('POISON');
    bisectPackages(pack(deck(24)), pack(deck(24, poisonAmongMany(5))), { oracle: probe.oracle });

    const asked = probe.seen.map((bytes) => entryText(bytes, MAIN_ENTRY));
    expect(new Set(asked).size).toBe(asked.length);
  });

  it('answers a repeated configuration from the cache instead of the oracle', () => {
    // Worth its own test because the straightforward reduction never repeats a
    // question: when a subset fails on its own, `ddmin` takes it and never
    // reaches the complements. The repetition happens exactly when no subset
    // fails alone - here, two changes that are only fatal together - and then
    // the complement of the first half *is* the second half, already asked. At
    // a second and a half per run against PowerPoint that is not a rounding
    // error, so the saving is counted.
    const broken = deck(20, (i) =>
      i === 4
        ? '<sp name="LEFT" x="40"/>'
        : i === 15
          ? '<sp name="RIGHT" x="150"/>'
          : '<sp name="renamed' + String(i) + '" x="' + String(i * 10 + 1) + '"/>',
    );

    const result = bisectPackages(pack(deck(20)), pack(broken), {
      oracle: (bytes) => {
        const text = entryText(bytes, MAIN_ENTRY);
        return text.includes('LEFT') && text.includes('RIGHT') ? 'fails' : 'passes';
      },
    });

    expect(result.cached).toBeGreaterThan(0);
  });

  it('counts an unresolved run instead of reading it as a pass', () => {
    // An oracle that cannot answer must not be able to shrink the result. Here
    // it never answers after the two preconditions, so nothing can be dropped
    // and the whole delta comes back - the safe direction, and the one that
    // says to go and look somewhere else.
    let call = 0;
    const result = bisectPackages(pack(deck(12)), pack(deck(12, poisonAmongMany(3))), {
      oracle: () => {
        call += 1;
        if (call === 1) return 'passes';
        if (call === 2) return 'fails';
        return 'unresolved';
      },
    });

    expect(result.unresolved).toBeGreaterThan(0);
    // Every leaf, which is the whole delta expressed as finely as it can be.
    // `changes` is the forest's *roots* - one per entry - so comparing against
    // that would have compared 23 against 1 and looked like a bug in the search
    // rather than a difference between two words.
    expect(result.minimal).toHaveLength(leafChanges(result.changes).length);
  });
});

describe('the delta, as a report', () => {
  it('paths a change the way the file writes it, prefixes and all', () => {
    const original = pack('<p:doc xmlns:p="urn:p"><p:a v="1"/><p:a v="2"/></p:doc>');
    const broken = pack('<p:doc xmlns:p="urn:p"><p:a v="1"/><p:a v="9"/></p:doc>');

    const leaves = flattenChanges(collectDelta(readZip(original), readZip(broken)).changes).filter(
      (change) => change.children.length === 0,
    );
    // Indexed because there are two `p:a`, prefixed because the document is.
    expect(leaves.map((change) => change.where)).toContain('/p:doc/p:a[2]/@v');
  });

  it('carries both sides of every change, so a report can show them', () => {
    const original = pack('<doc><a v="before"/></doc>');
    const broken = pack('<doc><a v="after"/></doc>');

    const attribute = flattenChanges(collectDelta(readZip(original), readZip(broken)).changes).find(
      (change) => change.kind === 'attribute',
    )!;
    expect(attribute.was).toBe(' v="before"');
    expect(attribute.text).toBe(' v="after"');
  });
});

import { describe, expect, it } from 'vitest';
import { checkCorpus, humanBytes, type CheckInput, type ManifestFile } from './check.ts';
import { CAPS, type CorpusEntry, type RuleId } from './schema.ts';

/**
 * One `it` per rule, against synthetic manifests.
 *
 * `checkCorpus` takes the file table rather than reading the disk, so every
 * rule - including the two that are about the filesystem - is exercised here
 * without a temporary directory, and a rule can be tested against a corpus
 * shape this repository does not have yet.
 *
 * The pattern throughout: start from something clean, break exactly one thing,
 * assert exactly one rule fires. A test that asserts "some violation" would
 * pass for the wrong reason as soon as a second rule started firing too.
 */

const SHA_A = 'a'.repeat(64);
const SHA_B = 'b'.repeat(64);

/**
 * A clean entry, with anything overridden.
 *
 * `overrides` is deliberately untyped. Most of these tests exist to feed the
 * checker something a `CorpusEntry` could not hold - a missing licence, a
 * British spelling, a path that escapes the tree - and typing the parameter
 * would mean writing `as never` on nearly every call.
 */
function entry(overrides: Record<string, unknown> = {}): CorpusEntry {
  return {
    id: 'a01-example',
    kind: 'deck',
    storage: 'committed',
    path: 'a01-example.pptx',
    license: 'CC0-1.0',
    source: 'synthetic',
    sha256: SHA_A,
    bytes: 1024,
    features: { shape: 3 },
    description: 'A synthetic deck.',
    addedIn: '1.1',
    ...overrides,
  };
}

/**
 * A corpus, with the file table derived from the entries.
 *
 * `censusKeys` is `['shape']` and the default entry reports three of them, so
 * the baseline corpus is fully covered and needs no `uncovered` array at all.
 * That is deliberate: `C019`'s clean state is a corpus with no declared gaps,
 * and a baseline that carried one would put an apology in every other test.
 *
 * `envelope` merges into the manifest object itself, for the handful of rules
 * that are about the envelope rather than an entry.
 */
function input(
  entries: readonly CorpusEntry[],
  overrides: Partial<CheckInput> = {},
  envelope: Record<string, unknown> = {},
): CheckInput {
  const files = entries
    .filter((item) => item.storage === 'committed' && typeof item.path === 'string')
    .map((item) => ({
      path: 'corpus/decks/' + String(item.path),
      bytes: item.bytes,
      sha256: item.sha256,
    }));
  return {
    manifests: [
      {
        path: 'corpus/decks/manifest.json',
        dir: 'corpus/decks',
        json: {
          manifestVersion: 1,
          collection: 'decks',
          generatedIn: '1.1',
          adr: 'docs/adr/0009-the-corpus.md',
          serializers: { xml: 'tools/corpus/gen', container: 'tools/ground-truth/zip.ts' },
          entries,
          ...envelope,
        },
      },
    ],
    files,
    binaryExtensions: new Set(['pptx', 'pptm', 'png']),
    toolPaths: new Set(['tools/bench/make-deck.ts']),
    censusKeys: ['shape'],
    ...overrides,
  };
}

/** A well-formed declared gap, with anything overridden. */
function gap(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    key: 'model3d',
    why: 'The extension GUID that carries it is in no public specification.',
    closes: 'Experiment E4, and the a28-model3d slot it would unblock.',
    ...overrides,
  };
}

/** The two-key corpus C019 needs: one covered, one that has to be accounted for. */
const TWO_KEYS: Partial<CheckInput> = { censusKeys: ['model3d', 'shape'] };

/**
 * An empty census table, for the tests whose corpus makes no coverage claim.
 *
 * `C019` asks that every census key be exercised by some deck, so a corpus of
 * one fixture - or one deck that is `pinned`, whose bytes are in nobody's
 * clone - fails it for every key in the table. That is the rule working, and it
 * has its own tests below; here it would just be a second violation in a test
 * about something else, which is what the one-broken-thing discipline exists to
 * avoid.
 */
const NOT_ABOUT_COVERAGE: Partial<CheckInput> = { censusKeys: [] };

function rules(result: readonly { rule: RuleId }[]): RuleId[] {
  return [...new Set(result.map((violation) => violation.rule))].sort();
}

describe('a corpus that follows the rules', () => {
  it('reports nothing', () => {
    expect(checkCorpus(input([entry()]))).toEqual([]);
  });

  it('accepts a generated entry, which has a recipe and no path', () => {
    const generated = entry({
      id: 'big-deck',
      storage: 'generated',
      path: undefined,
      outputName: 'big-deck.pptx',
      recipe: { tool: 'tools/bench/make-deck.ts', args: ['big-deck', '<out>'] },
      bytes: 200_000_000,
    });
    expect(checkCorpus(input([generated]))).toEqual([]);
  });

  it('accepts a pinned entry, which can be located by nothing', () => {
    const pinned = entry({
      id: 'local-deck-01',
      storage: 'pinned',
      path: undefined,
      features: { shape: 12 },
    });
    // The gap has to be declared, and that is C019 rather than an accident of
    // this fixture: a pinned deck is a hash of bytes no clone has, so a corpus
    // whose only deck is pinned covers nothing at all.
    const declared = { uncovered: [gap({ key: 'shape' })] };
    expect(checkCorpus(input([pinned], {}, declared))).toEqual([]);
  });
});

describe('C001-no-licence', () => {
  it('fires when the field is missing', () => {
    const result = checkCorpus(input([entry({ license: undefined })]));
    expect(rules(result)).toEqual(['C001-no-licence']);
    expect(result[0]?.message).toContain('no `license` field');
  });

  it('rejects NOASSERTION and NONE by name', () => {
    // Both are real SPDX spellings meaning "nobody established this", and
    // accepting either turns the rule into a formality.
    for (const value of ['NONE', 'NOASSERTION', '']) {
      const result = checkCorpus(input([entry({ license: value })]));
      expect(rules(result), value).toEqual(['C001-no-licence']);
      expect(result[0]?.message).toContain('cannot be redistributed');
    }
  });

  it('rejects a licence outside the closed set', () => {
    const result = checkCorpus(input([entry({ license: 'MIT' })]));
    expect(rules(result)).toEqual(['C001-no-licence']);
  });
});

describe('C002-no-source', () => {
  it('fires when the field is missing', () => {
    expect(rules(checkCorpus(input([entry({ source: undefined })])))).toEqual(['C002-no-source']);
  });

  it('fires on a source outside the closed set', () => {
    expect(rules(checkCorpus(input([entry({ source: 'the internet' })])))).toEqual([
      'C002-no-source',
    ]);
  });
});

describe('C003-source-licence', () => {
  it('refuses to relicense a third-party CC work as Apache-2.0', () => {
    // The rule that does the actual legal work. C001 is satisfied by typing
    // six characters; this one says whose work it is.
    const result = checkCorpus(
      input([
        entry({
          source: 'cc',
          license: 'Apache-2.0',
          attribution: {
            creator: 'Someone Else',
            title: 'A deck',
            sourceUrl: 'https://example.org/deck',
            licenseUrl: 'https://creativecommons.org/licenses/by/4.0/',
            modified: false,
          },
        }),
      ]),
    );
    expect(rules(result)).toEqual(['C003-source-licence']);
    expect(result[0]?.message).toContain('Allowed for source "cc"');
  });

  it('refuses anything but CC0 for a file our own tooling wrote', () => {
    expect(
      rules(checkCorpus(input([entry({ source: 'synthetic', license: 'Apache-2.0' })]))),
    ).toEqual(['C003-source-licence']);
  });

  it('allows Apache-2.0 for something we authored', () => {
    expect(checkCorpus(input([entry({ source: 'self-authored', license: 'Apache-2.0' })]))).toEqual(
      [],
    );
  });
});

describe('C004-attribution', () => {
  const cc = (attribution: unknown): CorpusEntry =>
    entry({ source: 'cc', license: 'CC-BY-4.0', attribution: attribution });

  it('requires the block at all', () => {
    expect(rules(checkCorpus(input([cc(undefined)])))).toEqual(['C004-attribution']);
  });

  it('requires every field CC BY 3(a)(1) names', () => {
    const result = checkCorpus(
      input([cc({ creator: 'A', title: '', sourceUrl: '', modified: false })]),
    );
    expect(rules(result)).toEqual(['C004-attribution']);
    expect(result.length).toBe(3); // title, sourceUrl, licenseUrl
  });

  it('requires a description of the change when modified is true', () => {
    const result = checkCorpus(
      input([
        cc({
          creator: 'A',
          title: 'T',
          sourceUrl: 'https://example.org',
          licenseUrl: 'https://creativecommons.org/licenses/by/4.0/',
          modified: true,
        }),
      ]),
    );
    expect(rules(result)).toEqual(['C004-attribution']);
    expect(result[0]?.message).toContain('3(a)(1)(B)');
  });
});

describe('C005-unknown-key', () => {
  it('catches the British spelling, which is the one that would happen here', () => {
    // Every other document in this repository writes "licence". A manifest
    // that does parses, validates and means the entry has no licence.
    const bad = { ...entry({ license: undefined }), licence: 'CC0-1.0' };
    const result = checkCorpus(input([bad]));
    expect(rules(result)).toEqual(['C001-no-licence', 'C005-unknown-key']);
  });

  it('allows $-prefixed commentary', () => {
    const commented = { ...entry(), $note: 'why this deck exists' };
    expect(checkCorpus(input([commented]))).toEqual([]);
  });

  it('catches an unknown key in the envelope', () => {
    const result = checkCorpus({
      ...input([entry()]),
      manifests: [
        {
          path: 'corpus/decks/manifest.json',
          dir: 'corpus/decks',
          json: {
            manifestVersion: 1,
            collection: 'decks',
            generatedIn: '1.1',
            adr: 'docs/adr/0009-the-corpus.md',
            serializers: { xml: 'tools/corpus/gen', container: 'tools/ground-truth/zip.ts' },
            entries: [entry()],
            licence: 'CC0-1.0',
          },
        },
      ],
    });
    expect(rules(result)).toEqual(['C005-unknown-key']);
  });

  it('rejects a collection that is not its own directory name', () => {
    // Every collection is named after the directory it lives in. Left
    // unchecked, two manifests could both call themselves `bench`, and every
    // message about either would name the wrong directory.
    const one = input([entry()]);
    const result = checkCorpus({
      ...one,
      manifests: [
        {
          path: 'corpus/decks/manifest.json',
          dir: 'corpus/decks',
          json: {
            manifestVersion: 1,
            collection: 'bench',
            generatedIn: '1.1',
            adr: 'docs/adr/0009-the-corpus.md',
            serializers: { xml: 'tools/corpus/gen', container: 'tools/ground-truth/zip.ts' },
            entries: [entry()],
          },
        },
      ],
    });
    expect(rules(result)).toEqual(['C005-unknown-key']);
    expect(result[0]?.message).toContain('in a directory named "decks"');
  });
});

describe('C006-id', () => {
  it('rejects an id that is not kebab-case', () => {
    for (const id of ['A01', 'a01_example', 'a01 example', '-a01', '']) {
      expect(rules(checkCorpus(input([entry({ id })]))), id).toContain('C006-id');
    }
  });

  it('rejects a duplicate across two manifests, not just within one', () => {
    const one = input([entry()]);
    const result = checkCorpus({
      ...one,
      manifests: [
        ...one.manifests,
        {
          path: 'corpus/bench/manifest.json',
          dir: 'corpus/bench',
          json: {
            manifestVersion: 1,
            collection: 'bench',
            generatedIn: '0.8',
            adr: 'docs/adr/0008-inspect-worker-and-the-benchmark.md',
            entries: [entry({ storage: 'pinned', path: undefined })],
          },
        },
      ],
    });
    expect(rules(result)).toContain('C006-id');
  });
});

describe('C007-storage', () => {
  it('requires a path when the bytes are committed', () => {
    expect(rules(checkCorpus(input([entry({ path: undefined })])))).toContain('C007-storage');
  });

  it('forbids a path on a generated entry', () => {
    // There is no file to point at, which is the whole reason for the
    // distinction: "(not committed - 188.5 MiB)" as a path was prose in a
    // field that is supposed to be machine-checkable.
    const result = checkCorpus(
      input([
        entry({
          storage: 'generated',
          outputName: 'x.pptx',
          recipe: { tool: 'tools/bench/make-deck.ts', args: [] },
        }),
      ]),
    );
    expect(rules(result)).toEqual(['C007-storage']);
  });

  it('forbids anything locatable on a pinned entry', () => {
    const result = checkCorpus(input([entry({ storage: 'pinned', path: 'decks/somewhere.pptx' })]));
    expect(rules(result)).toContain('C007-storage');
  });

  it('refuses a path that escapes corpus/', () => {
    const escaping = entry({ path: '../../etc/passwd' });
    const result = checkCorpus({ ...input([]), files: [], manifests: input([escaping]).manifests });
    expect(rules(result)).toContain('C007-storage');
    expect(result.some((violation) => violation.message.includes('outside corpus/'))).toBe(true);
  });

  it('refuses an absolute path', () => {
    const result = checkCorpus({
      ...input([]),
      manifests: input([entry({ path: '/etc/passwd' })]).manifests,
    });
    expect(rules(result)).toContain('C007-storage');
  });
});

describe('C008-orphan', () => {
  it('fires on a file no manifest claims', () => {
    // The only rule that catches the thing that actually goes wrong: a deck
    // dropped into corpus/ without anyone touching a manifest.
    const result = checkCorpus({
      ...input([entry()]),
      files: [
        { path: 'corpus/decks/a01-example.pptx', bytes: 1024, sha256: SHA_A },
        { path: 'corpus/decks/board-review.pptx', bytes: 2_000_000, sha256: SHA_B },
      ],
    });
    expect(rules(result)).toEqual(['C008-orphan']);
    expect(result[0]?.where).toContain('board-review.pptx');
    expect(result[0]?.message).toContain('establishing its provenance');
  });
});

describe('C009-double-claim', () => {
  it('fires when two entries claim one file', () => {
    const result = checkCorpus(input([entry(), entry({ id: 'a02-other' })]));
    expect(rules(result)).toEqual(['C009-double-claim']);
    expect(result[0]?.message).toContain('One file, one provenance');
  });
});

describe('C010-sha256 and C011-bytes', () => {
  it('fires when the committed bytes are not the bytes the manifest describes', () => {
    const result = checkCorpus({
      ...input([entry()]),
      files: [{ path: 'corpus/decks/a01-example.pptx', bytes: 1024, sha256: SHA_B }],
    });
    expect(rules(result)).toEqual(['C010-sha256']);
    expect(result[0]?.message).toContain('licence a statement about specific');
  });

  it('fires when the size disagrees', () => {
    const result = checkCorpus({
      ...input([entry()]),
      files: [{ path: 'corpus/decks/a01-example.pptx', bytes: 999, sha256: SHA_A }],
    });
    expect(rules(result)).toEqual(['C011-bytes']);
  });
});

describe('C012-size-cap', () => {
  it('fires on a committed file over the per-file cap', () => {
    const big = CAPS.perFile + 1;
    const result = checkCorpus({
      ...input([entry({ bytes: big })]),
      files: [{ path: 'corpus/decks/a01-example.pptx', bytes: big, sha256: SHA_A }],
    });
    expect(rules(result)).toEqual(['C012-size-cap']);
    expect(result[0]?.message).toContain('storage: "generated"');
  });

  it('does not fire on the same size when it is generated rather than committed', () => {
    const generated = entry({
      storage: 'generated',
      path: undefined,
      outputName: 'huge.pptx',
      recipe: { tool: 'tools/bench/make-deck.ts', args: [] },
      bytes: 200_000_000,
    });
    expect(checkCorpus(input([generated]))).toEqual([]);
  });

  it('fires on the total even when every file is under the per-file cap', () => {
    const count = Math.ceil(CAPS.total / CAPS.perFile) + 1;
    const entries = Array.from({ length: count }, (_, index) =>
      entry({
        id: 'a' + String(index).padStart(3, '0') + '-filler',
        path: 'decks/a' + String(index).padStart(3, '0') + '.pptx',
        bytes: CAPS.perFile,
      }),
    );
    const result = checkCorpus(input(entries));
    expect(rules(result)).toEqual(['C012-size-cap']);
    expect(result[0]?.where).toBe('corpus/');
  });
});

describe('C013-feature-key', () => {
  it('rejects a key that is not in the census table', () => {
    const result = checkCorpus(input([entry({ features: { smartart: 2 } })], NOT_ABOUT_COVERAGE));
    expect(rules(result)).toEqual(['C013-feature-key']);
    expect(result[0]?.message).toContain('not a census feature key');
  });

  it('accepts a count of zero, which is a statement and not an omission', () => {
    // A zero says the census looked and found none, which is a fact worth
    // recording and is not coverage. `C019` reads it the same way, so the key
    // still has to be declared - that agreement is the point of the test.
    const result = checkCorpus(
      input([entry({ features: { shape: 3, model3d: 0 } })], TWO_KEYS, { uncovered: [gap()] }),
    );
    expect(result).toEqual([]);
  });

  it('requires a features map on a deck', () => {
    expect(rules(checkCorpus(input([entry({ features: undefined })], NOT_ABOUT_COVERAGE)))).toEqual(
      ['C013-feature-key'],
    );
  });

  it('lets a non-package fixture carry tags instead', () => {
    const fixture = entry({
      id: 'color-transforms',
      kind: 'fixture',
      path: 'color-transforms.json',
      features: undefined,
      tags: ['tint', 'shade'],
    });
    expect(checkCorpus(input([fixture], NOT_ABOUT_COVERAGE))).toEqual([]);
  });

  it('refuses both at once', () => {
    const both = entry({ features: { shape: 1 }, tags: ['x'] });
    expect(rules(checkCorpus(input([both])))).toEqual(['C013-feature-key']);
  });
});

describe('C014-recipe-tool', () => {
  it('fires when the recipe names a tool that does not exist', () => {
    const result = checkCorpus(
      input([entry({ recipe: { tool: 'tools/bench/no-such-thing.ts', args: [] } })]),
    );
    expect(rules(result)).toEqual(['C014-recipe-tool']);
    expect(result[0]?.message).toContain('is a comment');
  });

  it('fires when the tool is outside tools/', () => {
    expect(
      rules(checkCorpus(input([entry({ recipe: { tool: 'scripts/build.sh', args: [] } })]))),
    ).toEqual(['C014-recipe-tool']);
  });
});

describe('C015-derivation', () => {
  it('requires derivedFrom for a redacted derivative', () => {
    expect(
      rules(checkCorpus(input([entry({ source: 'redacted-derivative', license: 'CC0-1.0' })]))),
    ).toEqual(['C015-derivation']);
  });

  it('refuses a derivation whose original hashes the same as what is committed', () => {
    const result = checkCorpus(
      input([
        entry({
          source: 'redacted-derivative',
          derivedFrom: {
            sha256: SHA_A,
            bytes: 1024,
            filename: 'original.pptx',
            redaction: 'nothing',
          },
        }),
      ]),
    );
    expect(rules(result)).toEqual(['C015-derivation']);
    expect(result[0]?.message).toContain('nothing was');
  });

  it('refuses derivedFrom on an entry that is not a derivative', () => {
    const result = checkCorpus(
      input([
        entry({
          derivedFrom: { sha256: SHA_B, bytes: 1, filename: 'x', redaction: 'y' },
        }),
      ]),
    );
    expect(rules(result)).toEqual(['C015-derivation']);
  });
});

describe('C016-gitattributes', () => {
  it('fires on a committed binary with no `binary` attribute', () => {
    const result = checkCorpus({
      ...input(
        [
          entry({
            id: 'probe-font',
            kind: 'asset',
            path: 'fonts/probe.eot',
            features: undefined,
            tags: ['eot'],
          }),
        ],
        NOT_ABOUT_COVERAGE,
      ),
      files: [{ path: 'corpus/decks/fonts/probe.eot', bytes: 1024, sha256: SHA_A }],
    });
    expect(rules(result)).toEqual(['C016-gitattributes']);
    expect(result[0]?.message).toContain('byte fidelity is the product');
  });

  it('does not ask for a binary attribute on genuinely textual files', () => {
    const fixture = entry({
      id: 'color-transforms',
      kind: 'fixture',
      path: 'color-transforms.json',
      features: undefined,
      tags: ['tint'],
    });
    expect(checkCorpus(input([fixture], NOT_ABOUT_COVERAGE))).toEqual([]);
  });
});

describe('C017-sorted', () => {
  it('fires when entries are out of order', () => {
    const result = checkCorpus(
      input([entry({ id: 'z01-last' }), entry({ id: 'a01-first', path: 'decks/other.pptx' })]),
    );
    expect(rules(result)).toContain('C017-sorted');
  });
});

describe('C018-serializers', () => {
  /** The default envelope with `serializers` replaced, or removed entirely. */
  function withSerializers(serializers: unknown): CheckInput {
    const base = input([entry()]);
    const json = { ...(base.manifests[0]!.json as Record<string, unknown>) };
    if (serializers === undefined) delete json['serializers'];
    else json['serializers'] = serializers;
    return { ...base, manifests: [{ ...base.manifests[0]!, json }] };
  }

  it('fires when a collection with committed decks declares none', () => {
    expect(rules(checkCorpus(withSerializers(undefined)))).toEqual(['C018-serializers']);
  });

  it('fires when either layer is missing or empty', () => {
    for (const bad of [
      { xml: 'tools/corpus/gen' },
      { container: 'tools/ground-truth/zip.ts' },
      { xml: '', container: 'tools/ground-truth/zip.ts' },
      { xml: 'tools/corpus/gen', container: '   ' },
      { xml: 42, container: 'tools/ground-truth/zip.ts' },
      'tools/corpus/gen',
    ]) {
      expect(rules(checkCorpus(withSerializers(bad))), JSON.stringify(bad)).toEqual([
        'C018-serializers',
      ]);
    }
  });

  it('does not ask it of a collection whose decks are generated', () => {
    // `corpus/bench` names a generator and commits no deck bytes, so there is
    // nothing for `C-LEX` to attribute and requiring the field there would be
    // a value nothing consumes.
    const generated = entry({
      id: 'big-deck',
      storage: 'generated',
      path: undefined,
      outputName: 'big-deck.pptx',
      recipe: { tool: 'tools/bench/make-deck.ts', args: ['big-deck', '<out>'] },
    });
    const base = input([generated]);
    const json = { ...(base.manifests[0]!.json as Record<string, unknown>) };
    delete json['serializers'];
    expect(checkCorpus({ ...base, manifests: [{ ...base.manifests[0]!, json }] })).toEqual([]);
  });

  it('lets two collections name the same serializer, which is the point', () => {
    // `corpus/authored` and `corpus/written` both declare PowerPoint as their
    // XML serializer because Tier C's parts *are* Tier B's parts. `C-LEX`
    // counts the collision as one producer; nothing here may forbid it.
    const base = input([entry()]);
    const second = {
      path: 'corpus/written/manifest.json',
      dir: 'corpus/written',
      json: {
        manifestVersion: 1,
        collection: 'written',
        generatedIn: '1.1',
        adr: 'docs/adr/0009-the-corpus.md',
        serializers: { xml: 'tools/corpus/gen', container: 'packages/opc' },
        entries: [entry({ id: 'c01-copy', path: undefined, storage: 'pinned' })],
      },
    };
    expect(checkCorpus({ ...base, manifests: [...base.manifests, second] })).toEqual([]);
  });
});

describe('C019-coverage', () => {
  /** The default corpus, with a census key nothing in it exercises. */
  const short = (envelope: Record<string, unknown> = {}): CheckInput =>
    input([entry()], TWO_KEYS, envelope);

  it('fires when a census key is covered by nothing and declared nowhere', () => {
    const result = checkCorpus(short());
    expect(rules(result)).toEqual(['C019-coverage']);
    expect(result[0]?.where).toBe('corpus/');
    expect(result[0]?.message).toContain('no deck exercises `model3d`');
  });

  it('accepts the gap once somebody has signed for it', () => {
    expect(checkCorpus(short({ uncovered: [gap()] }))).toEqual([]);
  });

  it('fires when a declared gap has been filled, so the array cannot rot', () => {
    // The direction that matters six months from now. Without it, the entry
    // outlives the hole and the file goes on apologising for something that
    // was fixed - which is worse than silence, because it reads as current.
    const covers = entry({ features: { shape: 1, model3d: 2 } });
    const result = checkCorpus(input([covers], TWO_KEYS, { uncovered: [gap()] }));
    expect(rules(result)).toEqual(['C019-coverage']);
    expect(result[0]?.message).toContain('a01-example exercises it');
  });

  it('will not take a key on its own', () => {
    for (const field of ['why', 'closes']) {
      const result = checkCorpus(short({ uncovered: [gap({ [field]: undefined })] }));
      expect(rules(result), field).toEqual(['C019-coverage']);
      expect(result[0]?.message, field).toContain('`' + field + '` is required');
    }
  });

  it('will not take a blank one either', () => {
    const result = checkCorpus(short({ uncovered: [gap({ why: '   ' })] }));
    expect(rules(result)).toEqual(['C019-coverage']);
  });

  it('rejects a gap in something that is not a census key', () => {
    // A typo here is the worst failure the rule has: it looks like a signed
    // statement, and the key it was meant to cover is still undeclared.
    const result = checkCorpus(short({ uncovered: [gap({ key: 'model3D' })] }));
    expect(rules(result)).toEqual(['C019-coverage']);
    expect(result.map((violation) => violation.message).join('\n')).toContain(
      'declares a gap in nothing',
    );
    expect(result.map((violation) => violation.message).join('\n')).toContain(
      'no deck exercises `model3d`',
    );
  });

  it("rejects an unrecognised field, for C005's reason", () => {
    const result = checkCorpus(short({ uncovered: [gap({ reason: 'blocked' })] }));
    expect(rules(result)).toEqual(['C019-coverage']);
    expect(result[0]?.message).toContain('unrecognised key `reason`');
  });

  it('rejects the same gap declared twice', () => {
    const result = checkCorpus(short({ uncovered: [gap(), gap()] }));
    expect(rules(result)).toEqual(['C019-coverage']);
    expect(result[0]?.message).toContain('One gap, one statement');
  });

  it("wants the array sorted, for C017's reason", () => {
    const result = checkCorpus(
      input(
        [entry({ features: {} })],
        { censusKeys: ['chart', 'model3d', 'shape'] },
        {
          uncovered: [gap(), gap({ key: 'chart' }), gap({ key: 'shape' })],
        },
      ),
    );
    expect(result.map((violation) => violation.message)).toContain(
      'not sorted by key, for the same reason `C017` sorts entries',
    );
  });

  it('does not count a pinned deck as coverage', () => {
    // Its bytes are on one machine and in no clone, so a coverage claim
    // resting on it rests on a file nobody else has.
    const pinned = entry({ storage: 'pinned', path: undefined, features: { shape: 1 } });
    const result = checkCorpus(input([pinned], { censusKeys: ['shape'] }));
    expect(rules(result)).toEqual(['C019-coverage']);
  });

  it('counts a generated deck, whose bytes anyone can reproduce', () => {
    const generated = entry({
      id: 'big-deck',
      storage: 'generated',
      path: undefined,
      outputName: 'big-deck.pptx',
      recipe: { tool: 'tools/bench/make-deck.ts', args: [] },
      features: { shape: 1 },
    });
    expect(checkCorpus(input([generated], { censusKeys: ['shape'] }))).toEqual([]);
  });

  it('lets one collection cover what another says nothing about', () => {
    // The gap is corpus-wide, so it is settled only after every manifest has
    // been read. A per-manifest rule would fire on whichever collection
    // happens not to hold the deck, which is every collection but one.
    const base = input([entry()], TWO_KEYS);
    const second: ManifestFile = {
      path: 'corpus/written/manifest.json',
      dir: 'corpus/written',
      json: {
        manifestVersion: 1,
        collection: 'written',
        generatedIn: '1.1',
        adr: 'docs/adr/0009-the-corpus.md',
        serializers: { xml: 'tools/corpus/gen', container: 'packages/opc' },
        entries: [entry({ id: 'c01-copy', path: 'c01-copy.pptx', features: { model3d: 1 } })],
      },
    };
    expect(
      checkCorpus({
        ...base,
        manifests: [...base.manifests, second],
        files: [
          ...base.files,
          { path: 'corpus/written/c01-copy.pptx', bytes: 1024, sha256: SHA_A },
        ],
      }),
    ).toEqual([]);
  });
});

describe('humanBytes', () => {
  it('is readable at every magnitude a corpus reaches', () => {
    expect(humanBytes(512)).toBe('512 B');
    expect(humanBytes(2048)).toBe('2.0 KiB');
    expect(humanBytes(197_631_762)).toBe('188.5 MiB');
  });
});

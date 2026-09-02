/**
 * The corpus's legal posture, executable.
 *
 * `LEGAL.md` says which sources a fixture may come from and which licences may
 * be claimed for them. Prose in a repository decays: this is the same policy in
 * a form that fails a build. If the two ever disagree, `LEGAL.md` is the one
 * that is wrong, because this is the one a contributor actually runs into.
 *
 * The important rule is not `C001-no-licence` on its own. Requiring a `license`
 * field is satisfied by typing six characters. `C003-source-licence` is the one
 * that does the legal work: it says a deck somebody else wrote and released
 * CC-BY is not ours to relicense as Apache-2.0, whatever the field says.
 */

// ------------------------------------------------------------------ licences

/**
 * The licences a corpus entry may claim. Deliberately a closed set and not the
 * SPDX list: an entry needing a licence that is not here is an entry that needs
 * a conversation, not a wider enum.
 */
export const LICENSES = {
  'CC0-1.0': 'public-domain dedication; the default for anything we authored',
  'Apache-2.0': 'when the fixture is really code - a recipe, a probe generator',
  'CC-BY-4.0': 'third-party work under CC BY; requires a complete `attribution` block',
  'LicenseRef-US-Gov-PublicDomain':
    'a US federal work, 17 U.S.C. 105. SPDX has no identifier for this; ' +
    '`LicenseRef-` is the sanctioned form for a licence outside the list',
} as const;

export type License = keyof typeof LICENSES;

/**
 * Where a fixture came from. This is provenance, not licence: the two are
 * separate questions and conflating them is how a third-party deck ends up
 * labelled CC0 because somebody generated a variant of it.
 */
export const SOURCES = [
  'self-authored',
  'synthetic',
  'cc',
  'gov-public-domain',
  'redacted-derivative',
] as const;

export type Source = (typeof SOURCES)[number];

/**
 * Which licences each source may claim.
 *
 * `synthetic` is CC0 only, because a file our own tooling wrote from nothing has
 * no other honest answer. `cc` may not be Apache-2.0, because relicensing
 * somebody else's work is exactly the thing this table exists to stop.
 */
export const ALLOWED: Readonly<Record<Source, readonly License[]>> = {
  'self-authored': ['CC0-1.0', 'Apache-2.0'],
  synthetic: ['CC0-1.0'],
  cc: ['CC0-1.0', 'CC-BY-4.0'],
  'gov-public-domain': ['LicenseRef-US-Gov-PublicDomain', 'CC0-1.0'],
  'redacted-derivative': ['CC0-1.0', 'Apache-2.0'],
};

/**
 * Spellings that look like an answer and are not.
 *
 * SPDX defines both, and both mean "nobody established this". An entry whose
 * licence is unknown is an entry that cannot be redistributed, so accepting
 * either would turn the whole check into a formality.
 */
export const NON_ANSWERS = ['NONE', 'NOASSERTION', 'UNKNOWN', 'TBD', ''] as const;

// --------------------------------------------------------------------- sizes

/**
 * Byte budgets, because a corpus is cloned by everyone who contributes.
 *
 * A file over the per-file cap is not forbidden - it just has to be
 * `storage: "generated"`, so what is committed is the recipe and the hash
 * rather than 188 MiB of noise.
 */
export const CAPS = {
  /** One committed file. */
  perFile: 512 * 1024,
  /** Everything committed under `corpus/`, across every collection. */
  total: 12 * 1024 * 1024,
} as const;

// -------------------------------------------------------------------- shapes

export type Kind = 'deck' | 'asset' | 'fixture' | 'record';
export type Storage = 'committed' | 'generated' | 'pinned';

/**
 * Which directory under `corpus/` an entry lives in. `C005` checks that a
 * manifest's `collection` equals its own directory's name, so this enum and the
 * filesystem cannot drift apart.
 *
 * There is one directory per producer, and that is the whole taxonomy: `decks`
 * is what `tools/corpus/gen` writes, `authored` is what PowerPoint writes, and
 * `written` is what `packages/opc`'s own writer writes. The split was forced
 * rather than chosen: `build-probes.ts --manifest` rewrites
 * `corpus/decks/manifest.json` in full from `PROBE_DECKS`, so an entry from any
 * other producer parked there would survive exactly until the next time a Tier A
 * deck changed.
 *
 * The three also make different claims about reproducibility, and one directory
 * per producer lets each manifest state only the claim it can honour:
 *
 * | collection | recipe | why                                                    |
 * | ---------- | ------ | ------------------------------------------------------ |
 * | `decks`    | yes    | the generator is a pure function of its deck module     |
 * | `authored` | no     | PowerPoint stamps `dcterms:created` on every save       |
 * | `written`  | yes    | a no-op write recompresses nothing, so it is bit-stable |
 *
 * `written`'s recipe is the surprising one, because a deflating writer normally
 * cannot promise reproducible bytes at all - deflated output depends on which
 * zlib produced it, which is why Tier A stores its entries rather than
 * compressing them. It holds here because a no-op write touches no part:
 * `passthroughEntry` moves every already-compressed stream across without
 * inflating it, so nothing in the output depends on which version of `fflate`
 * happened to be installed.
 */
export type Collection =
  'decks' | 'authored' | 'written' | 'reject' | 'bench' | 'ground-truth' | 'local';

export const KINDS: readonly Kind[] = ['deck', 'asset', 'fixture', 'record'];
export const STORAGES: readonly Storage[] = ['committed', 'generated', 'pinned'];
export const COLLECTIONS: readonly Collection[] = [
  'decks',
  'authored',
  'written',
  'reject',
  'bench',
  'ground-truth',
  'local',
];

export interface Producer {
  readonly application: string;
  readonly version: string;
  readonly build: string | null;
  readonly platform: string;
}

/**
 * Who wrote each layer of every deck in a collection, for `C-LEX` (`C018`).
 *
 * Two names rather than one, because a package has two lexical surfaces and
 * they do not have to come from the same place. `corpus/written` is the case
 * that forces it: `c01-opc-writer` is `b01-blank` read through `PartStore` and
 * written straight back out, so its thirty-seven ZIP headers are ours and every
 * byte inside every part is still Microsoft's. A single `producer` field cannot
 * say that, and `C-LEX` counting one deck as one producer would score a copy as
 * independent evidence about XML lexical form.
 *
 * Declared per collection rather than per entry because within a collection it
 * is true by construction, and fifty-one repetitions of the same pair is fifty
 * chances for one of them to be wrong.
 *
 * The values are free-form on purpose: `tools/corpus/gen` names a directory,
 * `Microsoft PowerPoint 16.0.20326` names a build. What matters is only that
 * two collections sharing a serializer spell it identically, which is what
 * makes `written` and `authored` count once for XML and twice for containers.
 */
export interface Serializers {
  /** What serialized the bytes inside the parts. */
  readonly xml: string;
  /** What wrote the ZIP headers. */
  readonly container: string;
}

/**
 * A census feature key no deck exercises, and the statement that says so.
 *
 * `C-COV` is the rule that every census key is covered by at least one deck
 * **or** named here. The array is what makes the difference between a corpus
 * that is 46 keys wide and a corpus that claims to be 47: an undeclared gap is
 * how a coverage badge becomes a lie, and a declared one is a signature.
 *
 * Three fields rather than a bare key, for the same reason `C-LEX`'s rows carry
 * a `gap` string. "`model3d`" names a hole; it does not say whether anybody
 * looked, whether closing it is a morning's work or blocked on something
 * outside this repository, or what the next person should do about it. A gap
 * nobody can act on is barely better than one nobody declared.
 *
 * `C019` also fires when a declared key **is** covered, which is what stops the
 * array rotting: the entry has to be deleted in the same commit that adds the
 * deck, so the file cannot go on apologising for a hole that was filled.
 */
export interface UncoveredFeature {
  /** A census feature key, spelled as `census-keys.gen.ts` spells it. */
  readonly key: string;
  /** Why no deck exercises it. */
  readonly why: string;
  /** What would close it - a deck slot, an experiment, a decision. */
  readonly closes: string;
}

export const UNCOVERED_KEYS: readonly string[] = ['key', 'why', 'closes'];

export interface Recipe {
  readonly tool: string;
  readonly args: readonly string[];
}

/** CC BY 4.0 section 3(a)(1): creator, title, licence, link, and what changed. */
export interface Attribution {
  readonly creator: string;
  readonly title: string;
  readonly sourceUrl: string;
  readonly licenseUrl: string;
  readonly modified: boolean;
  readonly modifications?: string;
}

/**
 * A redacted derivative's upstream.
 *
 * `LEGAL.md`: "Anything else is committed only as a redacted derivative
 * together with the SHA-256 of the original, so a report can be reproduced
 * without redistributing the source deck."
 */
export interface DerivedFrom {
  readonly sha256: string;
  readonly bytes: number;
  readonly filename: string;
  readonly redaction: string;
}

export interface CorpusEntry {
  readonly id: string;
  readonly kind: Kind;
  readonly storage: Storage;
  readonly license: License;
  readonly source: Source;
  readonly sha256: string;
  readonly bytes: number;
  readonly description: string;
  readonly addedIn: string;

  /** `storage: "committed"` only. Relative to the manifest's own directory. */
  readonly path?: string;
  /** `storage: "generated"` only. What the recipe writes; never a location. */
  readonly outputName?: string;
  readonly recipe?: Recipe;

  readonly sourceNote?: string;
  readonly producer?: Producer;
  /** Census feature keys to counts. Required for `kind: "deck"`. */
  readonly features?: Readonly<Record<string, number>>;
  /** Free-form labels, for fixtures that are not packages and have no census. */
  readonly tags?: readonly string[];
  readonly attribution?: Attribution;
  readonly derivedFrom?: DerivedFrom;

  /** Whatever else a collection records about itself, e.g. `slides`, `zipEntries`. */
  readonly [extra: string]: unknown;
}

export interface CorpusManifest {
  readonly manifestVersion: 1;
  readonly collection: Collection;
  readonly generatedIn: string;
  readonly adr: string;
  readonly entries: readonly CorpusEntry[];
  readonly serializers?: Serializers;
  readonly uncovered?: readonly UncoveredFeature[];
  readonly [extra: string]: unknown;
}

/**
 * Keys an entry may carry.
 *
 * A closed set, and this is `C005`'s whole point: without it, `"licence"` -
 * the British spelling, which the rest of this repository uses in prose -
 * parses fine, validates fine, and means the entry has no licence.
 */
export const ENTRY_KEYS: readonly string[] = [
  'id',
  'kind',
  'storage',
  'license',
  'source',
  'sha256',
  'bytes',
  'description',
  'addedIn',
  'path',
  'outputName',
  'recipe',
  'sourceNote',
  'producer',
  'features',
  'tags',
  'attribution',
  'derivedFrom',
  // Collection-specific facts. Descriptive, not load-bearing.
  'slides',
  'zipEntries',
  'format',
  // `corpus/reject` only, and load-bearing there: `C-REJECT` reads `rule` to
  // know which of the validator's twenty-nine each fixture has to trip, and
  // `refusal` is the sentence PowerPoint gave when the markup was measured -
  // the only diagnostic that exists, and the thing somebody will want if the
  // rule ever has to be argued with.
  'rule',
  'refusal',
];

export const MANIFEST_KEYS: readonly string[] = [
  'manifestVersion',
  'collection',
  'generatedIn',
  'adr',
  'entries',
  'measuredOn',
  'generator',
  'targetCount',
  'uncovered',
  'serializers',
];

// --------------------------------------------------------------------- rules

export type RuleId =
  | 'C001-no-licence'
  | 'C002-no-source'
  | 'C003-source-licence'
  | 'C004-attribution'
  | 'C005-unknown-key'
  | 'C006-id'
  | 'C007-storage'
  | 'C008-orphan'
  | 'C009-double-claim'
  | 'C010-sha256'
  | 'C011-bytes'
  | 'C012-size-cap'
  | 'C013-feature-key'
  | 'C014-recipe-tool'
  | 'C015-derivation'
  | 'C016-gitattributes'
  | 'C017-sorted'
  | 'C018-serializers'
  | 'C019-coverage';

export interface Violation {
  readonly rule: RuleId;
  /**
   * Where the problem is, as a path plus a JSON Pointer where there is one.
   * A package has one manifest; a corpus manifest has fifty entries, so
   * `tools/layering`'s "name the directory" is not enough resolution here.
   */
  readonly where: string;
  readonly message: string;
}

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
export type Collection = 'decks' | 'bench' | 'ground-truth' | 'local';

export const KINDS: readonly Kind[] = ['deck', 'asset', 'fixture', 'record'];
export const STORAGES: readonly Storage[] = ['committed', 'generated', 'pinned'];
export const COLLECTIONS: readonly Collection[] = ['decks', 'bench', 'ground-truth', 'local'];

export interface Producer {
  readonly application: string;
  readonly version: string;
  readonly build: string | null;
  readonly platform: string;
}

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
  | 'C017-sorted';

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

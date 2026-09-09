import {
  ALLOWED,
  CAPS,
  COLLECTIONS,
  ENTRY_KEYS,
  KINDS,
  LICENSES,
  MANIFEST_KEYS,
  NON_ANSWERS,
  SOURCES,
  STORAGES,
  UNCOVERED_KEYS,
  type CorpusEntry,
  type CorpusManifest,
  type RuleId,
  type Violation,
} from './schema.ts';

/**
 * The corpus rules, as a pure function.
 *
 * Takes parsed manifests **and** a pre-computed table of the files on disk, so
 * that every rule - including the two that are about the filesystem - is
 * testable with synthetic input and no temporary directories. Same shape as
 * `checkLayering({ manifests, catalog })`, for the same reason.
 */

export interface ManifestFile {
  /** Workspace-relative path, e.g. `corpus/bench/manifest.json`. */
  readonly path: string;
  /** Workspace-relative directory the manifest's own `path` fields resolve against. */
  readonly dir: string;
  readonly json: unknown;
}

export interface CorpusFile {
  /** Workspace-relative path, e.g. `corpus/bench/results.json`. */
  readonly path: string;
  readonly bytes: number;
  readonly sha256: string;
}

export interface CheckInput {
  readonly manifests: readonly ManifestFile[];
  /** Every file under `corpus/`, manifests themselves excluded. */
  readonly files: readonly CorpusFile[];
  /** Extensions `.gitattributes` marks `binary`, without the dot, lowercased. */
  readonly binaryExtensions: ReadonlySet<string>;
  /** Paths that exist under `tools/`, for `C014`. */
  readonly toolPaths: ReadonlySet<string>;
  /**
   * Every census feature key, for `C013` and `C019`.
   *
   * Passed in rather than imported from `census-keys.gen.ts` so that `C019` is
   * a pure function of its input like every other rule. Its denominator is the
   * whole census table, so importing the real one would make the rule fire for
   * forty-six keys against any synthetic corpus, and the only way to test it
   * would be to stop testing it.
   */
  readonly censusKeys: readonly string[];
}

const SHA256 = /^[0-9a-f]{64}$/;
const ID = /^[a-z0-9][a-z0-9-]*$/;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function extensionOf(path: string): string {
  const dot = path.lastIndexOf('.');
  return dot === -1 ? '' : path.slice(dot + 1).toLowerCase();
}

/**
 * A file's owner, or null.
 *
 * `path` is relative to the manifest's directory and must not escape `corpus/`,
 * which is the same containment question `PackURI` answers for part names and
 * is asked here for the same reason.
 */
function resolveClaim(dir: string, path: string): string | null {
  if (path.startsWith('/') || path.includes('\\') || /^[a-zA-Z]:/.test(path)) return null;
  const segments = (dir + '/' + path).split('/');
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      if (out.length === 0) return null;
      out.pop();
      continue;
    }
    out.push(segment);
  }
  const joined = out.join('/');
  return joined.startsWith('corpus/') ? joined : null;
}

export function checkCorpus(input: CheckInput): Violation[] {
  const violations: Violation[] = [];
  const add = (rule: RuleId, where: string, message: string): void => {
    violations.push({ rule, where, message });
  };

  const licenseList = Object.keys(LICENSES).join(', ');
  const seenIds = new Map<string, string>();
  /** claimed corpus-relative path -> the entries claiming it */
  const claims = new Map<string, string[]>();
  const byPath = new Map(input.files.map((file) => [file.path, file]));
  /** census key -> the decks exercising it, for `C019`. */
  const coverage = new Map<string, string[]>();
  /** census key -> where it was declared uncovered. */
  const declaredUncovered = new Map<string, string>();
  let committedBytes = 0;

  for (const manifest of input.manifests) {
    if (!isRecord(manifest.json)) {
      add('C005-unknown-key', manifest.path, 'not a JSON object');
      continue;
    }
    const envelope = manifest.json as unknown as CorpusManifest;

    for (const key of Object.keys(manifest.json)) {
      if (key.startsWith('$')) continue;
      if (!MANIFEST_KEYS.includes(key)) {
        add(
          'C005-unknown-key',
          manifest.path + '#/' + key,
          'unrecognised manifest key `' + key + '`. Known: ' + MANIFEST_KEYS.join(', '),
        );
      }
    }
    if (envelope.manifestVersion !== 1) {
      add('C005-unknown-key', manifest.path + '#/manifestVersion', 'must be 1');
    }
    if (!COLLECTIONS.includes(envelope.collection)) {
      add(
        'C005-unknown-key',
        manifest.path + '#/collection',
        'must be one of ' + COLLECTIONS.join(', '),
      );
    } else if (envelope.collection !== manifest.dir.split('/').pop()) {
      // Every collection name is its directory's name. Left unchecked, two
      // manifests could both call themselves `bench` and every message about
      // either would name the wrong directory.
      add(
        'C005-unknown-key',
        manifest.path + '#/collection',
        '`collection` is "' +
          String(envelope.collection) +
          '" in a directory named "' +
          String(manifest.dir.split('/').pop()) +
          '". They must match.',
      );
    }

    const entries = Array.isArray(envelope.entries) ? envelope.entries : [];
    if (!Array.isArray(envelope.entries)) {
      add('C005-unknown-key', manifest.path + '#/entries', '`entries` must be an array');
      continue;
    }

    // --- C018: serializers --------------------------------------------------
    // `C-LEX` counts distinct serializers, not distinct files, and a manifest
    // is the only durable place that fact can live. Required of any collection
    // holding a deck whose bytes are committed, because those are the decks
    // `C-LEX` can read; a collection whose decks are `generated` has no bytes
    // to attribute, so demanding it there would be a field nothing consumes.
    const hasCommittedDeck = entries.some(
      (entry) => isRecord(entry) && entry['kind'] === 'deck' && entry['storage'] === 'committed',
    );
    const serializers = envelope.serializers as unknown;
    if (serializers === undefined) {
      if (hasCommittedDeck) {
        add(
          'C018-serializers',
          manifest.path + '#/serializers',
          'a collection with committed decks must declare `serializers.xml` and ' +
            '`serializers.container`. C-LEX counts producers, and two decks that share a ' +
            'serializer are one piece of evidence however many files they are.',
        );
      }
    } else if (!isRecord(serializers)) {
      add('C018-serializers', manifest.path + '#/serializers', '`serializers` must be an object');
    } else {
      for (const layer of ['xml', 'container']) {
        const named = serializers[layer];
        if (typeof named !== 'string' || named.trim() === '') {
          add(
            'C018-serializers',
            manifest.path + '#/serializers/' + layer,
            '`serializers.' + layer + '` must be a non-empty string naming what wrote that layer',
          );
        }
      }
    }

    // --- C019: the declared gaps -------------------------------------------
    // Shape only; whether each one is still true is settled at the end, once
    // every manifest has been read. A gap is a corpus-wide claim, and no single
    // manifest can know that no other collection covers the key.
    const uncovered = envelope.uncovered as unknown;
    if (uncovered !== undefined) {
      if (!Array.isArray(uncovered)) {
        add('C019-coverage', manifest.path + '#/uncovered', '`uncovered` must be an array');
      } else {
        const keys: string[] = [];
        for (const [index, raw] of uncovered.entries()) {
          const at = manifest.path + '#/uncovered/' + String(index);
          if (!isRecord(raw)) {
            add('C019-coverage', at, 'each entry is an object with key, why and closes');
            continue;
          }
          for (const key of Object.keys(raw)) {
            if (key.startsWith('$')) continue;
            if (!UNCOVERED_KEYS.includes(key)) {
              add(
                'C019-coverage',
                at,
                'unrecognised key `' + key + '`. Known: ' + UNCOVERED_KEYS.join(', '),
              );
            }
          }
          const key = raw['key'];
          if (typeof key !== 'string' || !input.censusKeys.includes(key)) {
            add(
              'C019-coverage',
              at,
              '`key` must be a census feature key. `' +
                String(key) +
                '` is not one, so this declares a gap in nothing.',
            );
            continue;
          }
          for (const field of ['why', 'closes']) {
            const value = raw[field];
            if (typeof value !== 'string' || value.trim() === '') {
              add(
                'C019-coverage',
                at + '  (' + key + ')',
                '`' +
                  field +
                  '` is required. A key on its own names a hole; it does not say ' +
                  (field === 'why'
                    ? 'whether anybody looked.'
                    : 'what the next person should do about it.'),
              );
            }
          }
          const first = declaredUncovered.get(key);
          if (first !== undefined) {
            add(
              'C019-coverage',
              at + '  (' + key + ')',
              'already declared uncovered by ' + first + '. One gap, one statement.',
            );
          } else {
            declaredUncovered.set(key, at);
          }
          keys.push(key);
        }
        const sortedKeys = [...keys].sort();
        if (keys.join(' ') !== sortedKeys.join(' ')) {
          add(
            'C019-coverage',
            manifest.path + '#/uncovered',
            'not sorted by key, for the same reason `C017` sorts entries',
          );
        }
      }
    }

    // --- C017: sorted -------------------------------------------------------
    const ids = entries.map((entry) => String((entry as CorpusEntry).id ?? ''));
    const sorted = [...ids].sort();
    if (ids.join('\u0000') !== sorted.join('\u0000')) {
      add(
        'C017-sorted',
        manifest.path + '#/entries',
        'entries are not sorted by id. A stable order keeps a diff to the entry ' +
          'that changed rather than to everything after it.',
      );
    }

    for (const [index, raw] of entries.entries()) {
      const at = manifest.path + '#/entries/' + String(index);
      if (!isRecord(raw)) {
        add('C005-unknown-key', at, 'entry is not an object');
        continue;
      }
      const entry = raw as unknown as CorpusEntry;
      const id = typeof entry.id === 'string' ? entry.id : '';
      const where = at + (id === '' ? '' : '  (id: ' + id + ')');

      // --- C005: unknown key -----------------------------------------------
      for (const key of Object.keys(raw)) {
        if (key.startsWith('$')) continue;
        if (!ENTRY_KEYS.includes(key)) {
          add(
            'C005-unknown-key',
            where,
            'unrecognised key `' +
              key +
              '`. Known: ' +
              ENTRY_KEYS.join(', ') +
              '. A misspelling here validates and means nothing.',
          );
        }
      }

      // --- C006: id ---------------------------------------------------------
      if (!ID.test(id)) {
        add('C006-id', where, 'id `' + id + '` must match ' + String(ID));
      } else {
        const first = seenIds.get(id);
        if (first !== undefined) {
          add(
            'C006-id',
            where,
            'id `' +
              id +
              '` is already used by ' +
              first +
              '. Ids are unique across all manifests.',
          );
        } else {
          seenIds.set(id, at);
        }
      }

      // --- C001 / C002 / C003: the legal core --------------------------------
      const license = entry.license;
      if (typeof license !== 'string' || NON_ANSWERS.includes(license as never)) {
        add(
          'C001-no-licence',
          where,
          (license === undefined ? 'no `license` field' : 'license is `' + String(license) + '`') +
            '. Accepted: ' +
            licenseList +
            '. NONE and NOASSERTION are not accepted: an entry whose licence ' +
            'is unknown is an entry that cannot be redistributed.',
        );
      } else if (!(license in LICENSES)) {
        add(
          'C001-no-licence',
          where,
          'unknown license `' + license + '`. Accepted: ' + licenseList,
        );
      }

      const source = entry.source;
      if (typeof source !== 'string' || !SOURCES.includes(source)) {
        add(
          'C002-no-source',
          where,
          (source === undefined ? 'no `source` field' : 'unknown source `' + String(source) + '`') +
            '. Accepted: ' +
            SOURCES.join(', '),
        );
      } else if (typeof license === 'string' && license in LICENSES) {
        const allowed = ALLOWED[source];
        if (!allowed.includes(license)) {
          add(
            'C003-source-licence',
            where,
            'source is "' +
              source +
              '" but license is "' +
              license +
              '". Allowed for source "' +
              source +
              '": ' +
              allowed.join(', ') +
              '.',
          );
        }
      }

      // --- C004: attribution -------------------------------------------------
      if (source === 'cc' || source === 'gov-public-domain') {
        const attribution = entry.attribution;
        if (!isRecord(attribution)) {
          add(
            'C004-attribution',
            where,
            'source "' +
              source +
              '" requires an `attribution` block naming creator, ' +
              'title, sourceUrl, licenseUrl and whether it was modified.',
          );
        } else {
          for (const field of ['creator', 'title', 'sourceUrl', 'licenseUrl']) {
            if (typeof attribution[field] !== 'string' || attribution[field] === '') {
              add('C004-attribution', where, '`attribution.' + field + '` is missing or empty');
            }
          }
          if (attribution['modified'] === true && !attribution['modifications']) {
            add(
              'C004-attribution',
              where,
              'modified is true, so `attribution.modifications` must say what changed. ' +
                'CC BY 4.0 section 3(a)(1)(B) requires indicating modifications.',
            );
          }
        }
      }

      // --- C015: derivation --------------------------------------------------
      if (source === 'redacted-derivative') {
        const derived = entry.derivedFrom;
        if (!isRecord(derived)) {
          add(
            'C015-derivation',
            where,
            'source "redacted-derivative" requires a `derivedFrom` block with the ' +
              "original's sha256, bytes, filename and what was redacted.",
          );
        } else {
          if (typeof derived['sha256'] !== 'string' || !SHA256.test(String(derived['sha256']))) {
            add('C015-derivation', where, '`derivedFrom.sha256` must be 64 lowercase hex digits');
          } else if (derived['sha256'] === entry.sha256) {
            add(
              'C015-derivation',
              where,
              "`derivedFrom.sha256` equals this entry's own sha256, so nothing was " +
                'redacted and the original is what is committed.',
            );
          }
          if (!derived['redaction']) {
            add('C015-derivation', where, '`derivedFrom.redaction` must say what was removed');
          }
        }
      } else if (entry.derivedFrom !== undefined) {
        add(
          'C015-derivation',
          where,
          '`derivedFrom` is only meaningful when source is "redacted-derivative"',
        );
      }

      // --- C007: storage -----------------------------------------------------
      const storage = entry.storage;
      if (typeof storage !== 'string' || !STORAGES.includes(storage)) {
        add(
          'C007-storage',
          where,
          (storage === undefined
            ? 'no `storage` field'
            : 'unknown storage `' + String(storage) + '`') +
            '. Accepted: ' +
            STORAGES.join(', '),
        );
      }
      if (typeof entry.kind !== 'string' || !KINDS.includes(entry.kind)) {
        add('C007-storage', where, 'kind must be one of ' + KINDS.join(', '));
      }
      if (typeof entry.description !== 'string' || entry.description === '') {
        add('C007-storage', where, '`description` is required and must not be empty');
      }
      if (typeof entry.addedIn !== 'string' || entry.addedIn === '') {
        add('C007-storage', where, '`addedIn` must name the sub-phase that added the entry');
      }
      if (typeof entry.sha256 !== 'string' || !SHA256.test(entry.sha256)) {
        add('C007-storage', where, '`sha256` must be 64 lowercase hex digits');
      }
      if (typeof entry.bytes !== 'number' || !Number.isInteger(entry.bytes) || entry.bytes < 0) {
        add('C007-storage', where, '`bytes` must be a non-negative integer');
      }

      if (storage === 'committed') {
        if (typeof entry.path !== 'string' || entry.path === '') {
          add('C007-storage', where, 'storage "committed" requires a `path`');
        }
        if (entry.outputName !== undefined) {
          add('C007-storage', where, 'storage "committed" must not carry `outputName`');
        }
      } else if (storage === 'generated') {
        if (typeof entry.outputName !== 'string' || entry.outputName === '') {
          add('C007-storage', where, 'storage "generated" requires an `outputName`');
        }
        if (entry.path !== undefined) {
          add(
            'C007-storage',
            where,
            'storage "generated" must not carry `path`: there is no file to point at, ' +
              'which is the whole reason for the distinction.',
          );
        }
        if (!isRecord(entry.recipe)) {
          add('C007-storage', where, 'storage "generated" requires a `recipe`');
        }
      } else if (storage === 'pinned') {
        if (entry.path !== undefined || entry.outputName !== undefined) {
          add(
            'C007-storage',
            where,
            'storage "pinned" records a hash and nothing that can locate the file. ' +
              'That is the point: nothing in the repository may find it.',
          );
        }
      }

      // --- C014: recipe tool -------------------------------------------------
      if (isRecord(entry.recipe)) {
        const tool = entry.recipe['tool'];
        if (typeof tool !== 'string' || !tool.startsWith('tools/')) {
          add('C014-recipe-tool', where, '`recipe.tool` must be a path under tools/');
        } else if (!input.toolPaths.has(tool)) {
          add(
            'C014-recipe-tool',
            where,
            '`recipe.tool` is ' +
              tool +
              ', which does not exist. A regenerate command ' +
              'that names a file nobody can run is a comment.',
          );
        }
        if (!Array.isArray(entry.recipe['args'])) {
          add('C014-recipe-tool', where, '`recipe.args` must be an array');
        }
      }

      // --- C013: feature keys ------------------------------------------------
      if (entry.kind === 'deck' && !isRecord(entry.features)) {
        add(
          'C013-feature-key',
          where,
          'kind "deck" requires a `features` map. It is produced by the census, ' +
            'not curated by hand: `pptx-studio inspect <deck> --json`.',
        );
      }
      if (isRecord(entry.features)) {
        for (const [key, count] of Object.entries(entry.features)) {
          if (!input.censusKeys.includes(key)) {
            add(
              'C013-feature-key',
              where,
              '`features.' +
                key +
                '` is not a census feature key. The keys are the `key` column of ' +
                'packages/census/src/features.ts.',
            );
          }
          if (typeof count !== 'number' || !Number.isInteger(count) || count < 0) {
            add('C013-feature-key', where, '`features.' + key + '` must be a non-negative integer');
          }
        }
      }
      // `C019` counts a key as covered when some deck reports a **non-zero**
      // count for it. A zero is a real statement - `a01-minimal` uses one to
      // say the census looked and found none - but it is the opposite of
      // coverage, and reading it as coverage would let one deck declaring every
      // key at zero score the corpus at 100%.
      //
      // `pinned` is excluded with it. A pinned entry is the hash of a file that
      // stays on the machine that made it, so a coverage claim resting on one
      // rests on bytes no contributor has.
      if (entry.kind === 'deck' && storage !== 'pinned' && isRecord(entry.features)) {
        for (const [featureKey, count] of Object.entries(entry.features)) {
          if (typeof count !== 'number' || count <= 0) continue;
          const decks = coverage.get(featureKey);
          if (decks === undefined) coverage.set(featureKey, [id]);
          else decks.push(id);
        }
      }

      if (entry.features !== undefined && entry.tags !== undefined) {
        add(
          'C013-feature-key',
          where,
          'an entry carries `features` (a package the census can read) or `tags` ' +
            '(anything else), never both.',
        );
      }

      // --- the filesystem: C008 / C009 / C010 / C011 / C012 / C016 -----------
      if (storage === 'committed' && typeof entry.path === 'string' && entry.path !== '') {
        const claimed = resolveClaim(manifest.dir, entry.path);
        if (claimed === null) {
          add(
            'C007-storage',
            where,
            '`path` ' +
              entry.path +
              ' resolves outside corpus/. Targets are relative ' +
              "to the manifest's own directory and may not escape it.",
          );
        } else {
          const already = claims.get(claimed);
          if (already === undefined) claims.set(claimed, [at]);
          else already.push(at);

          const file = byPath.get(claimed);
          if (file === undefined) {
            add('C007-storage', where, '`path` names ' + claimed + ', which is not on disk');
          } else {
            if (already === undefined) committedBytes += file.bytes;
            if (typeof entry.sha256 === 'string' && SHA256.test(entry.sha256)) {
              if (file.sha256 !== entry.sha256) {
                add(
                  'C010-sha256',
                  where,
                  claimed +
                    '\n      manifest ' +
                    entry.sha256 +
                    '\n      on disk  ' +
                    file.sha256 +
                    '\n      The hash is what makes the licence a statement about specific ' +
                    'bytes rather than about whatever is at that path today.',
                );
              }
            }
            if (typeof entry.bytes === 'number' && file.bytes !== entry.bytes) {
              add(
                'C011-bytes',
                where,
                claimed +
                  ': manifest says ' +
                  String(entry.bytes) +
                  ' bytes, on disk it is ' +
                  String(file.bytes),
              );
            }
            if (file.bytes > CAPS.perFile) {
              add(
                'C012-size-cap',
                where,
                claimed +
                  ' is ' +
                  humanBytes(file.bytes) +
                  ', over the ' +
                  humanBytes(CAPS.perFile) +
                  ' per-file cap. A fixture this size is `storage: "generated"`: ' +
                  'commit the recipe and the hash, not the bytes.',
              );
            }
            const extension = extensionOf(claimed);
            if (extension !== '' && !isTextExtension(extension)) {
              if (!input.binaryExtensions.has(extension)) {
                add(
                  'C016-gitattributes',
                  claimed,
                  '.gitattributes has no `*.' +
                    extension +
                    ' binary` line. Without one a checkout on a machine with different ' +
                    'line-ending settings can rewrite the file, and byte fidelity is the product.',
                );
              }
            }
          }
        }
      }
    }
  }

  // --- C009: two entries claiming one file ---------------------------------
  for (const [path, claimants] of claims) {
    if (claimants.length > 1) {
      add(
        'C009-double-claim',
        path,
        'claimed by ' + claimants.join(' and ') + '. One file, one provenance.',
      );
    }
  }

  // --- C008: files nobody claims -------------------------------------------
  for (const file of input.files) {
    if (!claims.has(file.path)) {
      add(
        'C008-orphan',
        file.path + '  (' + humanBytes(file.bytes) + ')',
        'no manifest entry claims this file. Two ways out, and only two: add an ' +
          'entry establishing its provenance and licence, or delete the file. ' +
          'Scratch output does not belong under corpus/ - every generator here ' +
          'takes an output path.',
      );
    }
  }

  // --- C019: coverage -------------------------------------------------------
  // `C-COV`, the last of sub-phase 1.1's four named rules. `C013` checks that
  // every key a deck claims is a real one; this is the converse, and the harder
  // direction: every key the census can report is exercised by some deck, or
  // somebody has signed a statement saying why not.
  //
  // Both directions matter. A key covered by nothing and declared nowhere is
  // how a coverage badge becomes a lie. A key declared uncovered that some deck
  // now exercises is how the declaration rots into an apology for a hole that
  // was filled - so that fires too, and the entry has to be deleted in the same
  // commit that adds the deck.
  for (const key of input.censusKeys) {
    const decks = coverage.get(key);
    const declared = declaredUncovered.get(key);
    if (decks === undefined || decks.length === 0) {
      if (declared === undefined) {
        add(
          'C019-coverage',
          'corpus/',
          'no deck exercises `' +
            key +
            "`. Either add one, or declare it in a manifest's `uncovered` array " +
            'with why nothing covers it and what would close it. A gap nobody ' +
            'wrote down is indistinguishable from one nobody noticed.',
        );
      }
    } else if (declared !== undefined) {
      add(
        'C019-coverage',
        declared + '  (' + key + ')',
        '`' +
          key +
          '` is declared uncovered, but ' +
          decks.join(', ') +
          ' exercise' +
          (decks.length === 1 ? 's' : '') +
          ' it. Delete the declaration.',
      );
    }
  }

  // --- C012: the total ------------------------------------------------------
  if (committedBytes > CAPS.total) {
    add(
      'C012-size-cap',
      'corpus/',
      humanBytes(committedBytes) +
        ' committed, over the ' +
        humanBytes(CAPS.total) +
        ' total cap. Everyone who contributes clones this.',
    );
  }

  return violations;
}

/** Extensions that are genuinely text and must not be forced binary. */
function isTextExtension(extension: string): boolean {
  return ['json', 'xml', 'md', 'txt', 'svg', 'csv', 'rels'].includes(extension);
}

export function humanBytes(bytes: number): string {
  if (bytes < 1024) return String(bytes) + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KiB';
  return (bytes / 1024 / 1024).toFixed(1) + ' MiB';
}

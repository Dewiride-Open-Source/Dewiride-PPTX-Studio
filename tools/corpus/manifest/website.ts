/**
 * `C020-website-copy`: the decks the public site serves are corpus decks, byte for byte.
 *
 * The site ships copies under `website/public/decks/` with a manifest of its own, because it
 * builds from its directory alone. A copy is allowed only when a corpus file with the same
 * SHA-256 exists, so the corpus's provenance rules cover the site without a second policy.
 */

import type { CorpusFile } from './check.ts';
import type { Violation } from './schema.ts';

export interface WebsiteDeck {
  readonly file: string;
  readonly sha256: string;
  readonly license: string;
  readonly source: string;
  readonly about: string;
}

/** The two provenance values a public site may serve; anything else needs a policy first. */
const SERVABLE = { license: 'CC0-1.0', source: 'self-authored' } as const;

export function readWebsiteManifest(json: unknown, where: string): readonly WebsiteDeck[] {
  if (
    typeof json !== 'object' ||
    json === null ||
    !Array.isArray((json as { decks?: unknown }).decks)
  ) {
    throw new Error(`${where}: expected { decks: [...] }`);
  }
  return (json as { decks: unknown[] }).decks.map((entry, at) => {
    const record = entry as Record<string, unknown>;
    for (const key of ['file', 'sha256', 'license', 'source', 'about'] as const) {
      if (typeof record[key] !== 'string' || record[key].length === 0) {
        throw new Error(`${where}#/decks/${String(at)}/${key}: expected a non-empty string`);
      }
    }
    return record as unknown as WebsiteDeck;
  });
}

export function checkWebsiteCopies(input: {
  readonly manifestPath: string;
  readonly decks: readonly WebsiteDeck[];
  /** The files under `website/public/decks/`, hashed. */
  readonly copies: readonly CorpusFile[];
  /** Every file under `corpus/`, hashed. */
  readonly corpus: readonly CorpusFile[];
}): Violation[] {
  const violations: Violation[] = [];
  const add = (where: string, message: string): void => {
    violations.push({ rule: 'C020-website-copy', where, message });
  };
  const originals = new Map(input.corpus.map((file) => [file.sha256, file.path]));
  const copies = new Map(input.copies.map((file) => [file.path.split('/').pop() ?? '', file]));
  const claimed = new Set<string>();

  input.decks.forEach((deck, at) => {
    const where = `${input.manifestPath}#/decks/${String(at)}`;
    claimed.add(deck.file);
    const copy = copies.get(deck.file);
    if (copy === undefined) {
      add(where, `${deck.file} is not under website/public/decks/`);
      return;
    }
    if (copy.sha256 !== deck.sha256) {
      add(where, `sha256 is ${deck.sha256} but the file is ${copy.sha256}`);
    }
    if (!originals.has(copy.sha256)) {
      add(
        where,
        `${deck.file} is byte-identical to no file under corpus/; the site serves corpus decks only`,
      );
    }
    if (deck.license !== SERVABLE.license || deck.source !== SERVABLE.source) {
      add(
        where,
        `license ${deck.license} from ${deck.source}; a public site serves ${SERVABLE.license} ${SERVABLE.source} decks only`,
      );
    }
  });

  for (const [name, file] of copies) {
    if (!claimed.has(name)) add(file.path, 'not claimed by website/public/decks/manifest.json');
  }
  return violations;
}

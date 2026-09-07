/**
 * Claim what the harness writes, in the corpus manifest.
 *
 * `pnpm corpus` refuses a file under `corpus/` that no entry claims, and it is
 * right to: a fixture with no provenance is a fixture nobody can redistribute.
 * Both generators write through here so that neither can add a file and forget
 * to say where it came from. ADR 0035.
 */

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';

import { repoPath } from '../repo/root.ts';

/** Where the fidelity fixtures live, relative to `corpus/ground-truth/`. */
export const FIXTURE_PREFIX = 'render/fidelity/';

export interface FixtureEntry {
  readonly id: string;
  readonly path: string;
  readonly tags: readonly string[];
  readonly description: string;
  readonly recipe: { tool: string; args: readonly string[] };
}

const MANIFEST = repoPath('corpus/ground-truth/manifest.json');

/**
 * What a fidelity fixture is, legally.
 *
 * The decks are ours and CC0. What is committed is a measurement of how
 * PowerPoint drew them, reduced to 8x8-pixel cell means - a resolution that
 * retains no glyph outline and no image. That is the same basis on which
 * sub-phase 0.7 committed measurements of Microsoft's fonts and never their
 * bytes. `LEGAL.md`, "Corpus licensing".
 */
const SOURCE_NOTE =
  'PowerPoint 365 rendered the corpus deck this entry names and exported each slide as a PNG; ' +
  'tools/ground-truth/render/fidelity/analyse.ts reduced it through tools/fidelity/metric/reduce.ts. ' +
  'The deck is ours and CC0-1.0; what is committed is a measurement of how PowerPoint drew it, at a ' +
  'resolution that retains no glyph outline and no image.';

/**
 * Add or replace entries, leaving every other entry alone.
 *
 * `sweepPrefix` additionally drops entries under that prefix that this call did
 * not claim, and only a caller that knows the complete set may pass it - a
 * caller that claims one file and sweeps a directory deletes the provenance of
 * every other file in it.
 */
export function claimFixtures(entries: readonly FixtureEntry[], sweepPrefix?: string): void {
  const manifest = JSON.parse(readFileSync(MANIFEST, 'utf8')) as {
    entries: readonly { id: string; path: string }[];
  };

  const claimed = entries.map((entry) => {
    const bytes = readFileSync(repoPath('corpus/ground-truth', entry.path));
    return {
      id: entry.id,
      kind: 'fixture',
      storage: 'committed',
      path: entry.path,
      license: 'CC0-1.0',
      source: 'self-authored',
      sourceNote: SOURCE_NOTE,
      sha256: createHash('sha256').update(bytes).digest('hex'),
      bytes: bytes.length,
      tags: [...entry.tags],
      description: entry.description,
      recipe: { tool: entry.recipe.tool, args: [...entry.recipe.args] },
      addedIn: '3.9',
    };
  });

  const ids = new Set(claimed.map((entry) => entry.id));
  const kept = manifest.entries.filter((entry) => {
    if (ids.has(entry.id)) return false;
    return sweepPrefix === undefined || !entry.path.startsWith(sweepPrefix);
  });
  // Sorted by id, which is `C017-sorted`: a diff stays on the entry that moved.
  const merged = [...kept, ...claimed].sort((a, b) => (a.id < b.id ? -1 : 1));

  writeFileSync(MANIFEST, `${JSON.stringify({ ...manifest, entries: merged }, null, 2)}\n`);
}

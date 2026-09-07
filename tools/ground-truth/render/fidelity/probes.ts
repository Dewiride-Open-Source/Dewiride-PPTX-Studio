/**
 * Experiment F1 - which slides the fidelity harness scores, and at what size.
 *
 * Unlike every other experiment in this directory there is no deck to author:
 * the subject is the corpus we already have. So there is no `build-deck.ts`
 * here, and its absence is the finding rather than a gap - a generator that
 * invented slides would be scoring our own idea of a slide against PowerPoint's
 * rendering of it. ADR 0035.
 */

import { readFileSync } from 'node:fs';

import { repoPath } from '../../../repo/root.ts';

/** The collections whose decks are scored, in the order they are reported. */
export const COLLECTIONS: readonly string[] = ['decks', 'authored', 'written'];

export interface FidelityProbe {
  readonly id: string;
  /** Repository-relative, so it is both a served URL and a path. */
  readonly path: string;
  readonly slides: number;
  readonly license: string;
}

interface ManifestEntry {
  readonly id: string;
  readonly kind: string;
  readonly storage: string;
  readonly path: string;
  readonly format?: string;
  readonly slides?: number;
  readonly license: string;
}

/**
 * Every committed `.pptx` in the scored collections, with its slide count.
 *
 * `.pptm` is excluded, and not because macros are hard: PowerPoint prompts on
 * opening one unless macro security is lowered, and a capture that depends on a
 * machine's security policy is not a capture anyone else can reproduce.
 */
export function fidelityProbes(): readonly FidelityProbe[] {
  const out: FidelityProbe[] = [];
  for (const collection of COLLECTIONS) {
    const manifest = JSON.parse(
      readFileSync(repoPath('corpus', collection, 'manifest.json'), 'utf8'),
    ) as { entries: readonly ManifestEntry[] };
    for (const entry of manifest.entries) {
      if (entry.kind !== 'deck' || entry.storage !== 'committed') continue;
      if (!entry.path.endsWith('.pptx')) continue;
      const slides = entry.slides;
      if (slides === undefined || slides <= 0) {
        throw new Error(`${entry.id} has no slide count in corpus/${collection}/manifest.json`);
      }
      out.push({
        id: entry.id,
        path: `corpus/${collection}/${entry.path}`,
        slides,
        license: entry.license,
      });
    }
  }
  return out;
}

/** `<deck>-<NN>`, the name a slide's grid and its oracle PNG both carry. */
export function slideKey(deckId: string, slide: number): string {
  return `${deckId}-${String(slide).padStart(2, '0')}`;
}

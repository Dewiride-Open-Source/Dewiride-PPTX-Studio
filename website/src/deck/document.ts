'use client';

import { useMemo } from 'react';

import { PartStore } from '@pptx-studio/opc';
import { loadDocument, type Document } from '@pptx-studio/model';

import { useDeck } from './provider';

export interface ParsedDeck {
  readonly store: PartStore;
  readonly document: Document;
}

export interface ParseState {
  readonly parsed: ParsedDeck | null;
  readonly error: string | null;
  readonly loading: boolean;
  readonly name: string | null;
  /** Bumped whenever an edit rewrites a part, so views recompute. */
  readonly bytes: Uint8Array | null;
}

/**
 * The deck, opened and resolved, memoised on the bytes.
 *
 * `PartStore.open` inflates and `loadDocument` walks every sheet's own `.rels`,
 * so doing it once per tool rather than once per render is the difference
 * between a tab that responds and one that does not.
 */
export function useDocument(): ParseState {
  const { deck, error, loading } = useDeck();

  const result = useMemo(() => {
    if (deck === null) return { parsed: null, parseError: null };
    try {
      const store = PartStore.open(deck.bytes);
      return { parsed: { store, document: loadDocument(store) }, parseError: null };
    } catch (cause) {
      return {
        parsed: null,
        parseError: cause instanceof Error ? cause.message : String(cause),
      };
    }
  }, [deck]);

  return {
    parsed: result.parsed,
    error: error ?? result.parseError,
    loading,
    name: deck?.name ?? null,
    bytes: deck?.bytes ?? null,
  };
}

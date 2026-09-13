'use client';

import { useMemo } from 'react';

import { loadDocument, type Document } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import {
  createTextEngine,
  mediaFromStore,
  type MediaResolver,
  type TextEngine,
} from '@pptx-studio/render-svg';

import type { Failure } from '@/design/state';
import { describeFailure } from './failures';
import { useDeck } from './provider';

export interface ParsedDeck {
  readonly store: PartStore;
  readonly document: Document;
  /** Pictures by `r:embed`, resolved through the part they were written in. */
  readonly media: MediaResolver;
  /** One measurer and one font probe, shared by every slide the deck mounts. */
  readonly text: TextEngine;
}

export interface ParseState {
  readonly parsed: ParsedDeck | null;
  readonly failure: Failure | null;
  readonly loading: boolean;
  readonly name: string | null;
  readonly bytes: Uint8Array | null;
}

/**
 * The deck, opened and resolved, memoised on the bytes.
 *
 * `PartStore.open` inflates and `loadDocument` walks every sheet's own `.rels`,
 * so doing it once per deck rather than once per render is the difference
 * between a tab that responds and one that does not.
 */
export function useDocument(): ParseState {
  const { deck, failure, loading } = useDeck();

  const result = useMemo<{ parsed: ParsedDeck | null; failure: Failure | null }>(() => {
    if (deck === null) return { parsed: null, failure: null };
    try {
      const store = PartStore.open(deck.bytes);
      const document = loadDocument(store);
      return {
        parsed: {
          store,
          document,
          media: mediaFromStore(store),
          text: createTextEngine({ defaultTextStyle: document.defaultTextStyle }),
        },
        failure: null,
      };
    } catch (cause) {
      return { parsed: null, failure: describeFailure(cause) };
    }
  }, [deck]);

  return {
    parsed: result.parsed,
    failure: failure ?? result.failure,
    loading,
    name: deck?.name ?? null,
    bytes: deck?.bytes ?? null,
  };
}

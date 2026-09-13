'use client';

import {
  createContext,
  use,
  useCallback,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from 'react';

import { DEFAULT_SAMPLE } from './samples';

export interface LoadedDeck {
  readonly name: string;
  readonly bytes: Uint8Array;
}

interface DeckState {
  readonly deck: LoadedDeck | null;
  readonly error: string | null;
  readonly loading: boolean;
  load(file: File): void;
  loadSample(file: string): void;
}

const DeckContext = createContext<DeckState | null>(null);

/**
 * One deck, held in the root layout so it survives navigating between tools.
 *
 * The bytes never leave the tab except on the one route that says it is a
 * server route, and nothing is uploaded anywhere else.
 */
export function DeckProvider({ children }: { children: ReactNode }) {
  const [deck, setDeck] = useState<LoadedDeck | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  const loadSample = useCallback((file: string) => {
    setLoading(true);
    setError(null);
    fetch(`/decks/${file}`)
      .then(async (response) => {
        if (!response.ok) throw new Error(`${response.status} fetching ${file}`);
        return response.arrayBuffer();
      })
      .then((buffer) => setDeck({ name: file, bytes: new Uint8Array(buffer) }))
      .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)))
      .finally(() => setLoading(false));
  }, []);

  const load = useCallback((file: File) => {
    setLoading(true);
    setError(null);
    file
      .arrayBuffer()
      .then((buffer) => setDeck({ name: file.name, bytes: new Uint8Array(buffer) }))
      .catch((cause: unknown) => setError(cause instanceof Error ? cause.message : String(cause)))
      .finally(() => setLoading(false));
  }, []);

  // Open on a real deck rather than an empty dropzone: a tool that shows
  // nothing until you feed it demonstrates nothing.
  useEffect(() => {
    loadSample(DEFAULT_SAMPLE.file);
  }, [loadSample]);

  const value = useMemo<DeckState>(
    () => ({ deck, error, loading, load, loadSample }),
    [deck, error, loading, load, loadSample],
  );

  return <DeckContext value={value}>{children}</DeckContext>;
}

export function useDeck(): DeckState {
  const found = use(DeckContext);
  if (found === null) throw new Error('useDeck outside a DeckProvider');
  return found;
}

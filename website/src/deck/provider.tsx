'use client';

import {
  createContext,
  use,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';

import type { Failure } from '@/design/state';
import { publicUrl } from '@/site/base-path';
import { describeFailure } from './failures';
import { DEFAULT_SAMPLE } from './samples';

export interface LoadedDeck {
  readonly name: string;
  readonly bytes: Uint8Array;
  /** True for a deck the visitor opened, false for a shipped sample. */
  readonly own: boolean;
}

interface DeckState {
  readonly deck: LoadedDeck | null;
  readonly failure: Failure | null;
  readonly loading: boolean;
  readonly load: (file: File) => void;
  readonly loadSample: (file: string) => void;
}

const DeckContext = createContext<DeckState | null>(null);

async function fetchSample(file: string): Promise<LoadedDeck> {
  const response = await fetch(publicUrl(`/decks/${file}`));
  if (!response.ok) throw new Error(`${String(response.status)} fetching ${file}`);
  return { name: file, bytes: new Uint8Array(await response.arrayBuffer()), own: false };
}

/**
 * One deck, held in the root layout so it survives navigating between pages.
 *
 * The bytes never leave the tab; nothing on this site uploads anything.
 */
export function DeckProvider({ children }: { children: ReactNode }) {
  const [deck, setDeck] = useState<LoadedDeck | null>(null);
  const [failure, setFailure] = useState<Failure | null>(null);
  const [loading, setLoading] = useState(true);
  // A slow sample must not overwrite a deck the visitor dropped meanwhile.
  const generation = useRef(0);

  const settle = useCallback((pending: Promise<LoadedDeck>) => {
    const mine = ++generation.current;
    pending
      .then((loaded) => {
        if (mine !== generation.current) return;
        setDeck(loaded);
        setFailure(null);
      })
      .catch((cause: unknown) => {
        if (mine === generation.current) setFailure(describeFailure(cause));
      })
      .finally(() => {
        if (mine === generation.current) setLoading(false);
      });
  }, []);

  const loadSample = useCallback(
    (file: string) => {
      setLoading(true);
      settle(fetchSample(file));
    },
    [settle],
  );

  const load = useCallback(
    (file: File) => {
      setLoading(true);
      settle(
        file
          .arrayBuffer()
          .then((buffer) => ({ name: file.name, bytes: new Uint8Array(buffer), own: true })),
      );
    },
    [settle],
  );

  // Open on a real deck rather than an empty dropzone: a demo that shows
  // nothing until you feed it demonstrates nothing.
  useEffect(() => {
    settle(fetchSample(DEFAULT_SAMPLE.file));
  }, [settle]);

  const value = useMemo<DeckState>(
    () => ({ deck, failure, loading, load, loadSample }),
    [deck, failure, loading, load, loadSample],
  );

  return <DeckContext value={value}>{children}</DeckContext>;
}

export function useDeck(): DeckState {
  const found = use(DeckContext);
  if (found === null) throw new Error('useDeck outside a DeckProvider');
  return found;
}

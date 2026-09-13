'use client';

import { useEffect, useRef } from 'react';

import type { Failure } from '@/design/state';
import type {
  CensusDone,
  DeckRequest,
  DeckResponse,
  ExportDone,
  ReplacedPart,
  ValidateDone,
  WorkerEnvironment,
} from './protocol';

/** A request before it is numbered; `Omit` over the union would keep only the shared keys. */
type Unnumbered = DeckRequest extends infer R
  ? R extends DeckRequest
    ? Omit<R, 'id'>
    : never
  : never;

interface Pending {
  readonly resolve: (response: DeckResponse) => void;
  readonly reject: (failure: Failure) => void;
  readonly onProgress?: (done: number, total: number) => void;
}

export class DeckWorkerError extends Error {
  readonly failure: Failure;

  constructor(failure: Failure) {
    super(failure.message);
    this.name = 'DeckWorkerError';
    this.failure = failure;
  }
}

/** One Worker, one request at a time per id, answered as promises. */
export class DeckWorker {
  private readonly worker: Worker;
  private readonly pending = new Map<number, Pending>();
  private next = 1;
  readonly ready: Promise<WorkerEnvironment>;

  constructor() {
    this.worker = new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' });
    this.ready = new Promise((resolve) => {
      const onReady = (event: MessageEvent<DeckResponse>) => {
        if (event.data.kind !== 'ready') return;
        this.worker.removeEventListener('message', onReady);
        resolve(event.data.environment);
      };
      this.worker.addEventListener('message', onReady);
    });
    this.worker.addEventListener('message', (event: MessageEvent<DeckResponse>) => {
      const message = event.data;
      if (message.kind === 'ready') return;
      const waiting = this.pending.get(message.id);
      if (waiting === undefined) return;
      if (message.kind === 'progress') {
        waiting.onProgress?.(message.done, message.total);
        return;
      }
      this.pending.delete(message.id);
      if (message.kind === 'failed') waiting.reject(message.failure);
      else waiting.resolve(message);
    });
  }

  private ask(
    request: Unnumbered,
    onProgress?: (done: number, total: number) => void,
  ): Promise<DeckResponse> {
    const id = this.next++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, {
        resolve,
        reject: (failure) => reject(new DeckWorkerError(failure)),
        onProgress,
      });
      // The deck is shared with every other demo on the page, so it is copied
      // in; only what comes back is transferred.
      this.worker.postMessage({ ...request, id });
    });
  }

  async census(
    bytes: Uint8Array,
    onProgress?: (done: number, total: number) => void,
  ): Promise<CensusDone> {
    const answer = await this.ask({ kind: 'census', bytes: copy(bytes) }, onProgress);
    if (answer.kind !== 'census') throw new Error(`unexpected ${answer.kind}`);
    return answer;
  }

  async validate(bytes: Uint8Array, replaced: readonly ReplacedPart[] = []): Promise<ValidateDone> {
    const answer = await this.ask({ kind: 'validate', bytes: copy(bytes), replaced });
    if (answer.kind !== 'validate') throw new Error(`unexpected ${answer.kind}`);
    return answer;
  }

  async export(bytes: Uint8Array, replaced: readonly ReplacedPart[] = []): Promise<ExportDone> {
    const answer = await this.ask({ kind: 'export', bytes: copy(bytes), replaced });
    if (answer.kind !== 'export') throw new Error(`unexpected ${answer.kind}`);
    return answer;
  }

  terminate(): void {
    this.worker.terminate();
    for (const waiting of this.pending.values()) {
      waiting.reject({ message: 'the worker was stopped' });
    }
    this.pending.clear();
  }
}

function copy(bytes: Uint8Array): ArrayBuffer {
  return bytes.slice().buffer;
}

/** A worker that lives as long as the component that asked for it. */
export function useDeckWorker(): { current: DeckWorker | null } {
  const ref = useRef<DeckWorker | null>(null);
  useEffect(() => {
    const worker = new DeckWorker();
    ref.current = worker;
    return () => {
      worker.terminate();
      ref.current = null;
    };
  }, []);
  return ref;
}

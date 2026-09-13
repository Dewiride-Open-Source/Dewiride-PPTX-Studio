/**
 * The census, off the main thread.
 *
 * A 200 MB deck inflates to several hundred megabytes of XML and there is no
 * yield point inside a tokenizer, so doing this on the main thread does not
 * make the tab slow - it freezes it. `@pptx-studio/census` returns plain JSON
 * for exactly this reason: a `Map`, a `Set` or a class instance would not
 * survive `postMessage`.
 *
 * It also answers a question the project's architecture rests on, in the
 * environment the claim is about: whether `DOMParser` exists in a Worker.
 */

import { censusPackage, type PackageCensus } from '@pptx-studio/census';

export interface WorkerEnvironment {
  readonly hasDOMParser: boolean;
  readonly hasOffscreenCanvas: boolean;
  readonly hardwareConcurrency: number;
}

export type CensusRequest = { readonly bytes: ArrayBuffer };

export type CensusResponse =
  | { readonly type: 'ready'; readonly environment: WorkerEnvironment }
  | { readonly type: 'progress'; readonly done: number; readonly total: number }
  | {
      readonly type: 'done';
      readonly census: PackageCensus;
      readonly ms: number;
      readonly environment: WorkerEnvironment;
    }
  | { readonly type: 'failed'; readonly message: string };

const scope = self as unknown as {
  postMessage(message: CensusResponse): void;
  addEventListener(type: 'message', listener: (event: MessageEvent<CensusRequest>) => void): void;
  navigator: { hardwareConcurrency: number };
};

function probe(): WorkerEnvironment {
  const has = (name: string): boolean =>
    (globalThis as unknown as Record<string, unknown>)[name] !== undefined;
  return {
    hasDOMParser: has('DOMParser'),
    hasOffscreenCanvas: has('OffscreenCanvas'),
    hardwareConcurrency: scope.navigator.hardwareConcurrency,
  };
}

const environment = probe();

scope.addEventListener('message', (event) => {
  const startedAt = performance.now();
  try {
    const census = censusPackage(new Uint8Array(event.data.bytes), {
      onProgress: (done, total) => {
        if (done % 16 === 0 || done === total) scope.postMessage({ type: 'progress', done, total });
      },
    });
    scope.postMessage({
      type: 'done',
      census,
      ms: performance.now() - startedAt,
      environment,
    });
  } catch (cause) {
    scope.postMessage({
      type: 'failed',
      message: cause instanceof Error ? cause.message : String(cause),
    });
  }
});

scope.postMessage({ type: 'ready', environment });

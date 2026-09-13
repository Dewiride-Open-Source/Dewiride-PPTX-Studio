/**
 * Census, validation and export, off the main thread.
 *
 * None of the three has a yield point, and a visitor's deck can be tens of
 * megabytes, so on the main thread they would not make the tab slow - they
 * would freeze it. Every answer is plain JSON, which is why it survives
 * `postMessage`; the exported bytes are transferred, not copied.
 */

import { censusPackage } from '@pptx-studio/census';
import { PartStore } from '@pptx-studio/opc';
import { validatePackage } from '@pptx-studio/validate';
import { comparePackages, exportPackage, openPackage } from '@pptx-studio/writer';

import { describeFailure } from '../failures';
import type { DeckRequest, DeckResponse, ReplacedPart, WorkerEnvironment } from './protocol';

const scope = self as unknown as {
  postMessage(message: DeckResponse, transfer?: Transferable[]): void;
  addEventListener(type: 'message', listener: (event: MessageEvent<DeckRequest>) => void): void;
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

function opened(bytes: ArrayBuffer, replaced: readonly ReplacedPart[]) {
  const pkg = openPackage(new Uint8Array(bytes));
  for (const { part, bytes: content } of replaced) {
    pkg.store.replacePart(part, new Uint8Array(content));
  }
  return pkg;
}

scope.addEventListener('message', (event) => {
  const request = event.data;
  const startedAt = performance.now();
  try {
    if (request.kind === 'census') {
      const census = censusPackage(new Uint8Array(request.bytes), {
        onProgress: (done, total) => {
          if (done % 16 === 0 || done === total) {
            scope.postMessage({ kind: 'progress', id: request.id, done, total });
          }
        },
      });
      scope.postMessage({
        kind: 'census',
        id: request.id,
        census,
        ms: performance.now() - startedAt,
      });
    } else if (request.kind === 'validate') {
      const pkg = opened(request.bytes, request.replaced);
      const report = validatePackage({
        store: pkg.store,
        baseline: pkg.baseline,
        baselineBytes: pkg.baselineBytes,
        ...(request.replaced.length === 0 ? { bytes: new Uint8Array(request.bytes) } : {}),
      });
      scope.postMessage({
        kind: 'validate',
        id: request.id,
        report,
        ms: performance.now() - startedAt,
      });
    } else {
      const pkg = opened(request.bytes, request.replaced);
      const result = exportPackage(pkg);
      const comparison = comparePackages(
        PartStore.open(pkg.baselineBytes),
        PartStore.open(result.bytes),
      );
      const out = result.bytes.buffer.slice(
        result.bytes.byteOffset,
        result.bytes.byteOffset + result.bytes.byteLength,
      ) as ArrayBuffer;
      scope.postMessage(
        {
          kind: 'export',
          id: request.id,
          bytes: out,
          rewritten: result.rewritten,
          streamed: result.streamed,
          preservation: result.preservation,
          comparison,
          report: result.report,
          ms: performance.now() - startedAt,
        },
        [out],
      );
    }
  } catch (cause) {
    scope.postMessage({ kind: 'failed', id: request.id, failure: describeFailure(cause) });
  }
});

scope.postMessage({ kind: 'ready', environment: probe() });

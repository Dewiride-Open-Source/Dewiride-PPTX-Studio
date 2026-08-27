import { censusPackage } from '@pptx-studio/census';
import { isRelationshipPartName, PartStore } from '@pptx-studio/opc';
import { XmlTokenizer } from '@pptx-studio/xml';
import type {
  MemoryReading,
  StageTiming,
  RunTiming,
  WorkerEnvironment,
  WorkerRequest,
  WorkerResponse,
} from './protocol.js';

/**
 * The parse worker.
 *
 * Every byte of a `.pptx` is read here and nothing but counters goes back. That
 * is the whole point of the boundary: a 200 MB deck inflates to several hundred
 * megabytes of XML, and doing that on the main thread would freeze the tab for
 * as long as it takes - not degrade it, freeze it, because there is no yield
 * point inside a tokenizer.
 *
 * The one thing this file proves rather than asserts is at the top of
 * `probeEnvironment`. The plan's decision to hand-roll an XML tokenizer rests
 * partly on `DOMParser` being unavailable in a Web Worker; that is checked here
 * on every run, in the environment the claim is about, and reported to the page.
 *
 * `self` is deliberately reached through `globalThis` and a narrow interface
 * rather than through a `declare const self`, which would collide with the DOM
 * lib's `Window`-shaped one. The interface names exactly what this file uses:
 * if it starts using something else, it stops compiling instead of quietly
 * becoming `any`.
 */
interface WorkerScope {
  postMessage(message: WorkerResponse): void;
  addEventListener(type: 'message', listener: (event: { data: WorkerRequest }) => void): void;
  readonly crossOriginIsolated: boolean;
  readonly navigator: { readonly hardwareConcurrency: number; readonly userAgent: string };
}

/** The memory API, which the DOM lib does not declare and Chrome only exposes when isolated. */
interface MemoryCapablePerformance {
  measureUserAgentSpecificMemory?: () => Promise<{ bytes: number }>;
}

const scope = globalThis as unknown as WorkerScope;

function probeEnvironment(): WorkerEnvironment {
  const has = (name: string): boolean =>
    (globalThis as unknown as Record<string, unknown>)[name] !== undefined;
  return {
    hasDOMParser: has('DOMParser'),
    hasXMLHttpRequest: has('XMLHttpRequest'),
    hasOffscreenCanvas: has('OffscreenCanvas'),
    hasTextDecoderStream: has('TextDecoderStream'),
    hasMeasureMemory:
      (performance as unknown as MemoryCapablePerformance).measureUserAgentSpecificMemory !==
      undefined,
    crossOriginIsolated: scope.crossOriginIsolated,
    hardwareConcurrency: scope.navigator.hardwareConcurrency,
    userAgent: scope.navigator.userAgent,
  };
}

interface MemoryProbe {
  readonly bytes: number | null;
  readonly note: string | null;
}

async function measureMemory(): Promise<MemoryProbe> {
  const api = performance as unknown as MemoryCapablePerformance;
  if (api.measureUserAgentSpecificMemory === undefined) {
    return { bytes: null, note: 'performance.measureUserAgentSpecificMemory is not defined here' };
  }
  try {
    return { bytes: (await api.measureUserAgentSpecificMemory()).bytes, note: null };
  } catch (error) {
    // A SecurityError means the context is not cross-origin isolated. Saying so
    // is the point: a silent `null` is indistinguishable from "we forgot to
    // ask", and the difference decides whether the headers are wrong or the
    // browser simply does not offer the measurement.
    return { bytes: null, note: error instanceof Error ? error.message : String(error) };
  }
}

function errorCode(error: unknown): string {
  return typeof error === 'object' && error !== null && 'code' in error
    ? String(error.code)
    : 'UNKNOWN';
}

const environment = probeEnvironment();

scope.addEventListener('message', (event) => {
  const request = event.data;
  if (request.type !== 'parse') return;

  void run(request);
});

async function run(request: {
  id: number;
  name: string;
  bytes: ArrayBuffer;
  repeat: number;
  stages: boolean;
}): Promise<void> {
  const { id, name, repeat } = request;
  scope.postMessage({ type: 'accepted', id });
  const bytes = new Uint8Array(request.bytes);
  const before = await measureMemory();

  const runs: RunTiming[] = [];
  let census;

  try {
    for (let index = 0; index < Math.max(1, repeat); index++) {
      const startedAt = performance.now();
      census = censusPackage(bytes, {
        // Progress only on the first pass. A benchmark loop that posts a
        // thousand messages a second is measuring `postMessage`.
        onProgress:
          index === 0
            ? (done, total, part) => {
                if (done % 32 === 0 || done === total) {
                  scope.postMessage({ type: 'progress', id, done, total, part });
                }
              }
            : undefined,
      });
      runs.push({
        index,
        totalMs: performance.now() - startedAt,
        openMs: census.timings.openMs,
        inflateMs: census.timings.inflateMs,
        scanMs: census.timings.scanMs,
        xmlBytesScanned: census.timings.xmlBytesScanned,
      });
    }
  } catch (error) {
    scope.postMessage({
      type: 'failed',
      id,
      code: errorCode(error),
      message: error instanceof Error ? error.message : String(error),
    });
    return;
  }

  const stages = request.stages ? measureStages(bytes) : null;
  const after = await measureMemory();
  const memory: MemoryReading = {
    beforeBytes: before.bytes,
    afterBytes: after.bytes,
    available: before.bytes !== null || after.bytes !== null,
    note: after.note ?? before.note,
  };

  if (census === undefined) return;
  scope.postMessage({
    type: 'done',
    id,
    name,
    bytes: bytes.byteLength,
    census,
    runs,
    stages,
    memory,
    environment,
  });
}

/**
 * Inflate, decode and tokenize, timed separately.
 *
 * Deliberately re-does work the census has already done. The census reports one
 * number for "scan", which is the tokenizer and the counting together, and
 * those two are worth telling apart: if counting were the expensive half, the
 * fix would be in this repository's census and not in its tokenizer.
 *
 * Uses the real `PartStore` rather than a shortcut so the inflate number is the
 * one a user actually pays.
 */
function measureStages(bytes: Uint8Array): StageTiming {
  const store = PartStore.open(bytes);
  const names = store.partNames.filter(
    (name) => !isRelationshipPartName(name) && (store.contentTypeOf(name) ?? '').endsWith('+xml'),
  );

  let at = performance.now();
  const raw = names.map((name) => store.read(name));
  const inflateMs = performance.now() - at;
  const xmlBytes = raw.reduce((sum, part) => sum + part.byteLength, 0);

  const decoder = new TextDecoder('utf-8');
  at = performance.now();
  const sources = raw.map((part) => decoder.decode(part));
  const decodeMs = performance.now() - at;

  at = performance.now();
  let tokens = 0;
  for (const source of sources) {
    const tokenizer = new XmlTokenizer(source);
    for (let token = tokenizer.next(); token !== undefined; token = tokenizer.next()) tokens += 1;
  }
  const tokenizeMs = performance.now() - at;

  return { xmlParts: names.length, xmlBytes, tokens, inflateMs, decodeMs, tokenizeMs };
}

scope.postMessage({ type: 'ready', environment });

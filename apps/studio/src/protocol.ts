import type { PackageCensus } from '@pptx-studio/census';
import type { EditKind, ExportOutcome } from './export.js';

/**
 * The Worker boundary, written down.
 *
 * This file is the contract and it exists as its own module because both sides
 * have to agree and neither may own it. Three rules are baked into the shapes
 * below, and each of them is a decision rather than a convention.
 *
 * **The deck goes in as a transfer, never as a copy.** `bytes` is an
 * `ArrayBuffer` and every `postMessage` that carries one lists it in the
 * transfer array. Structured-cloning 200 MB would allocate a second 200 MB and
 * block the thread that did it; transferring is a pointer move and detaches the
 * original. That is why `ParseRequest.bytes` is an `ArrayBuffer` and not a
 * `Uint8Array` - the transfer list takes buffers, and asking for the buffer up
 * front makes it impossible to forget.
 *
 * **Nothing comes back but plain JSON.** `PackageCensus` is arrays, objects,
 * numbers and strings. No `Map`, no class instance, and above all no `XNode`:
 * a parsed tree holds a parent pointer on every node and the entire decoded
 * source on the document, so returning one would clone the deck a third time.
 * The tree stays in the worker and dies there.
 *
 * **Progress is a message, not a promise.** A 200 MB deck is seconds of work.
 * Without progress the tab is indistinguishable from a hang, and the worker is
 * the only thing that knows how far along it is.
 */

/** Parse and take a census. `bytes` is transferred and is detached afterwards. */
export interface ParseRequest {
  readonly type: 'parse';
  readonly id: number;
  readonly name: string;
  readonly bytes: ArrayBuffer;
  /** Run the census this many times and report every run. 1 unless benchmarking. */
  readonly repeat: number;
  /** Also time inflate, decode and tokenize on their own. */
  readonly stages: boolean;
}

/**
 * Read a package and write it back out. `bytes` is transferred and detached.
 *
 * A separate request rather than a flag on `ParseRequest`, because the two are
 * separate acts: a page can inspect a deck ten times and never export it, and
 * exporting needs the bytes a second time - the first set was detached by the
 * transfer that carried it here. The page re-reads the `File` rather than the
 * worker holding 200 MB alive on the chance somebody presses Save.
 */
export interface ExportRequest {
  readonly type: 'export';
  readonly id: number;
  readonly name: string;
  readonly bytes: ArrayBuffer;
  readonly edit: EditKind;
}

export type WorkerRequest = ParseRequest | ExportRequest;

/**
 * What the worker can actually see.
 *
 * Reported rather than assumed. The plan's second architectural bet rests on
 * `DOMParser` being unavailable in a Web Worker - that is the stated reason a
 * hand-rolled tokenizer exists at all - and a claim that load-bearing should be
 * measured in the environment it is a claim about, on every run, rather than
 * cited from a specification.
 */
export interface WorkerEnvironment {
  readonly hasDOMParser: boolean;
  readonly hasXMLHttpRequest: boolean;
  readonly hasOffscreenCanvas: boolean;
  readonly hasTextDecoderStream: boolean;
  /** Is `performance.measureUserAgentSpecificMemory` even defined here? */
  readonly hasMeasureMemory: boolean;
  readonly crossOriginIsolated: boolean;
  readonly hardwareConcurrency: number;
  readonly userAgent: string;
}

/** One census run, with the memory around it when the browser will say. */
export interface RunTiming {
  readonly index: number;
  /** Wall clock inside the worker, from the first byte to the finished report. */
  readonly totalMs: number;
  readonly openMs: number;
  readonly inflateMs: number;
  readonly scanMs: number;
  /** XML handed to the tokenizer, so throughput is derivable per run. */
  readonly xmlBytesScanned: number;
}

/**
 * Bytes of JS heap and off-heap allocation the browser attributes to the worker.
 *
 * `performance.measureUserAgentSpecificMemory()` is only available in a
 * cross-origin-isolated context, so both fields are `null` on a page served
 * without COOP and COEP. It is the only measurement available that includes the
 * ArrayBuffers, which for a 200 MB deck are all of the memory that matters -
 * `performance.memory` counts the JS heap and would report a 200 MB archive as
 * costing nothing.
 *
 * Measured once on each side of the whole run rather than around each census.
 * The call is queued until the browser finds a convenient moment and can take
 * seconds to settle, so putting it inside the loop would measure the browser's
 * scheduler instead of our parser.
 */
export interface MemoryReading {
  readonly beforeBytes: number | null;
  readonly afterBytes: number | null;
  readonly available: boolean;
  /** Why there is no reading, when there is none. */
  readonly note: string | null;
}

/**
 * Where the time actually goes, measured stage by stage.
 *
 * A single "parse took N seconds" is not a number anyone can act on, because
 * the four things it adds together have wildly different costs and only one of
 * them is ours to improve. Inflating is fflate's, decoding UTF-8 is the
 * platform's, tokenizing is this project's, and counting is the census's.
 *
 * Measured after the census runs, so the code is warm. A cold first pass would
 * charge the tokenizer for JIT compilation and make it look like the bottleneck
 * whether or not it is - which is exactly the mistake that sends someone off to
 * optimise the wrong loop.
 */
export interface StageTiming {
  readonly xmlParts: number;
  readonly xmlBytes: number;
  readonly tokens: number;
  /** Every XML part out of the archive, as bytes. */
  readonly inflateMs: number;
  /** Those bytes to strings. */
  readonly decodeMs: number;
  /** Those strings to tokens, and nothing else - no counting, no tree. */
  readonly tokenizeMs: number;
}

/**
 * Posted the instant the worker picks the request up, before any work.
 *
 * The point of it is to make one claim falsifiable. The deck is *transferred*
 * rather than structure-cloned, so handing over 200 MB should cost the same as
 * handing over 2 KB; if this number tracked file size, transfer would not be
 * happening and every parse would be paying for a second copy of the archive.
 *
 * It has to be its own message. Timing to the first progress message instead -
 * which is what this harness did at first - measures the transfer plus opening
 * the archive plus walking six hundred relationship parts, and reports the lot
 * as though it were the cost of `postMessage`.
 */
export interface AcceptedMessage {
  readonly type: 'accepted';
  readonly id: number;
}

export interface ProgressMessage {
  readonly type: 'progress';
  readonly id: number;
  readonly done: number;
  readonly total: number;
  readonly part: string;
}

export interface DoneMessage {
  readonly type: 'done';
  readonly id: number;
  readonly name: string;
  readonly bytes: number;
  readonly census: PackageCensus;
  readonly runs: readonly RunTiming[];
  readonly stages: StageTiming | null;
  readonly memory: MemoryReading;
  readonly environment: WorkerEnvironment;
}

export interface FailedMessage {
  readonly type: 'failed';
  readonly id: number;
  /** The `OpcError` / `XmlError` code when there is one, else `UNKNOWN`. */
  readonly code: string;
  readonly message: string;
}

export interface ReadyMessage {
  readonly type: 'ready';
  readonly environment: WorkerEnvironment;
}

/**
 * The exported package, and everything that was checked on the way out.
 *
 * `bytes` goes back the same way it came: transferred, not cloned. A structured
 * clone here would allocate a second copy of the whole archive on the main
 * thread at the exact moment the page is about to hand it to a `Blob`, which
 * allocates a third.
 */
export interface ExportedMessage {
  readonly type: 'exported';
  readonly id: number;
  readonly name: string;
  readonly bytes: ArrayBuffer;
  readonly outcome: ExportOutcome;
}

export type WorkerResponse =
  AcceptedMessage | ProgressMessage | DoneMessage | ExportedMessage | FailedMessage | ReadyMessage;

import type { DoneMessage, WorkerEnvironment, WorkerResponse } from './protocol.js';

/**
 * The page's half of the Worker boundary.
 *
 * One line in this file is the portable one, and it is called out because it is
 * the line that changes when `apps/studio` grows a framework:
 *
 * ```ts
 * new Worker(new URL('./worker.js', import.meta.url), { type: 'module' })
 * ```
 *
 * That expression is the canonical form. Vite, webpack 5 and Turbopack all
 * recognise it and rewrite it to their own emitted chunk; a plain bundler
 * leaves it alone and it resolves against the module's own URL at runtime,
 * which is why it works here with no plugin at all. Building the URL any other
 * way - concatenating a base, reading it from a global - is what makes a worker
 * that runs in development and 404s in production.
 *
 * `type: 'module'` is not optional. The worker imports `@pptx-studio/census`,
 * and a classic worker has `importScripts` and no `import` at all.
 */

export interface CensusResult extends DoneMessage {
  /**
   * Milliseconds from `postMessage` to the worker's first response.
   *
   * Measured because it is the number that would justify a different design.
   * The deck is transferred rather than cloned, so this should stay flat as the
   * file grows; if it tracked file size, transfer would not be happening and
   * every parse would be paying for a second copy of the deck.
   */
  readonly dispatchMs: number;
  /** Round trip as the page sees it, including the worker's own work. */
  readonly wallMs: number;
}

export interface RunOptions {
  readonly repeat?: number;
  /** Also time inflate, decode and tokenize separately. */
  readonly stages?: boolean;
  readonly onProgress?: (done: number, total: number, part: string) => void;
}

export class CensusWorker {
  readonly #worker: Worker;
  #nextId = 1;
  readonly ready: Promise<WorkerEnvironment>;

  constructor() {
    this.#worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
    this.ready = new Promise<WorkerEnvironment>((resolve, reject) => {
      const onMessage = (event: MessageEvent<WorkerResponse>): void => {
        if (event.data.type !== 'ready') return;
        this.#worker.removeEventListener('message', onMessage);
        resolve(event.data.environment);
      };
      this.#worker.addEventListener('message', onMessage);
      this.#worker.addEventListener('error', (event) => {
        reject(new Error('the parse worker failed to start: ' + event.message));
      });
    });
  }

  /**
   * Take a census of `buffer`, which is **transferred** and detached.
   *
   * The caller loses its reference to the bytes. That is the point: the
   * alternative is a structured clone, which for a 200 MB deck means allocating
   * another 200 MB and blocking the main thread while it copies.
   */
  run(name: string, buffer: ArrayBuffer, options: RunOptions = {}): Promise<CensusResult> {
    const id = this.#nextId++;
    const startedAt = performance.now();
    let dispatchMs = 0;

    return new Promise<CensusResult>((resolve, reject) => {
      const onMessage = (event: MessageEvent<WorkerResponse>): void => {
        const message = event.data;
        if ('id' in message && message.id !== id) return;

        if (message.type === 'accepted') {
          dispatchMs = performance.now() - startedAt;
          return;
        }
        if (message.type === 'progress') {
          options.onProgress?.(message.done, message.total, message.part);
          return;
        }
        if (message.type === 'failed') {
          this.#worker.removeEventListener('message', onMessage);
          reject(new Error(message.code + ': ' + message.message));
          return;
        }
        if (message.type === 'done') {
          this.#worker.removeEventListener('message', onMessage);
          resolve({ ...message, dispatchMs, wallMs: performance.now() - startedAt });
        }
      };

      this.#worker.addEventListener('message', onMessage);
      this.#worker.postMessage(
        {
          type: 'parse',
          id,
          name,
          bytes: buffer,
          repeat: options.repeat ?? 1,
          stages: options.stages ?? false,
        },
        [buffer],
      );
    });
  }

  terminate(): void {
    this.#worker.terminate();
  }
}

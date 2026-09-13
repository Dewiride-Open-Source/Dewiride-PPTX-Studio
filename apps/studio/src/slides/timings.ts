/** What the page measured while it drew: recorded and shown, never asserted (ADR 0054). */

export interface OpenTimings {
  /** Bytes in hand, from the read the source closure did. */
  readonly readMs: number;
  /** `openDeck`: the package parsed and the model loaded. */
  readonly parseMs: number;
  /** The first slide on the stage and painted, from the bytes. */
  readonly firstStageMs: number;
  /** Every thumbnail mounted, from the bytes, wall clock. */
  readonly stripMs: number;
  /** The main-thread time the thumbnails took, without the frames yielded between chunks. */
  readonly stripCpuMs: number;
  /** Chromium's used JS heap once the strip is done, where the page can ask. */
  readonly heapBytes: number | null;
}

export interface ShowTimings {
  /** The mount, zero when the stage already showed that slide at that zoom on this display. */
  readonly mountMs: number;
  readonly width: number;
  readonly height: number;
}

/** `performance.memory` where it exists, which is Chromium and nowhere else. */
export function heapBytes(): number | null {
  const memory = (performance as unknown as { memory?: { usedJSHeapSize?: number } }).memory;
  return typeof memory?.usedJSHeapSize === 'number' ? memory.usedJSHeapSize : null;
}

/** The next frame's timestamp, so a caller can wait for a paint. */
export function nextFrame(): Promise<number> {
  return new Promise((resolve) => requestAnimationFrame(resolve));
}

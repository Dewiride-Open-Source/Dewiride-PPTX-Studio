import { OpcError } from './errors.js';

/**
 * Decompression budgets.
 *
 * Every number in a ZIP header is attacker-controlled, including the ones that
 * say how big things are. So none of these limits can be enforced by trusting
 * the archive: they are enforced by refusing to allocate past them, and by a
 * running counter that survives across entries.
 *
 * The shape of the threat is worth stating, because it drives the design. The
 * classic 42.zip is recursive and irrelevant here - we never unpack a nested
 * archive. The one that matters is David Fifield's non-recursive
 * quoted-overlap bomb (WOOT '19), where many central-directory entries point
 * into overlapping byte ranges so the sum of the declared sizes grows
 * quadratically against the file size: 10 MB in, 281 TB out, one round of
 * decompression, standard DEFLATE. Three independent defences answer it:
 *
 * 1. `maxEntries` - the construction needs a great many entries.
 * 2. `maxTotalInflatedBytes` as a **running counter** across the whole
 *    archive, charged before each allocation, never reset per entry.
 * 3. Structural overlap detection in the reader, which rejects the shape
 *    outright rather than merely surviving it.
 */
export interface ZipLimits {
  /** Hard ceiling on the archive as read from disk or the network. */
  readonly maxArchiveBytes: number;
  /** Running total of inflated bytes across every entry read from one archive. */
  readonly maxTotalInflatedBytes: number;
  /** Ceiling for any single entry. No legitimate part of a deck approaches this. */
  readonly maxEntryInflatedBytes: number;
  /** Ceiling on central-directory entry count. */
  readonly maxEntries: number;
  /** Ceiling on inflated/compressed for one entry, above `ratioFloorBytes`. */
  readonly maxCompressionRatio: number;
  /**
   * Below this inflated size the ratio test is not applied.
   *
   * Without a floor the test is useless: a 40-byte DEFLATE stream expanding to
   * a 12 KB `[Content_Types].xml` is a 300:1 ratio and completely ordinary. A
   * ratio only means something once the absolute size is large enough to hurt.
   */
  readonly ratioFloorBytes: number;
  /** Ceiling on a ZIP entry name, in UTF-16 code units after decoding. */
  readonly maxNameLength: number;
}

const MIB = 1024 * 1024;

/**
 * The defaults, sized against the plan's stated envelope: a 200 MB, 300-slide
 * deck must load, and nothing much larger is a real presentation.
 */
export const DEFAULT_ZIP_LIMITS: ZipLimits = {
  maxArchiveBytes: 200 * MIB,
  maxTotalInflatedBytes: 1024 * MIB,
  maxEntryInflatedBytes: 256 * MIB,
  maxEntries: 20_000,
  maxCompressionRatio: 200,
  ratioFloorBytes: MIB,
  maxNameLength: 1024,
};

/**
 * The running counter.
 *
 * Charged **before** the allocation it authorises, with the size the archive
 * declares. If the archive lies low, the output buffer caps the damage and the
 * CRC catches the lie; if it lies high, this rejects the read before a single
 * byte is allocated. Either way the counter is monotonic for the lifetime of
 * one archive - that monotonicity is the whole defence, and is why it is an
 * object with state rather than a per-call argument.
 */
export class InflationBudget {
  readonly limit: number;
  #used = 0;

  constructor(limit: number) {
    this.limit = limit;
  }

  get used(): number {
    return this.#used;
  }

  get remaining(): number {
    return this.limit - this.#used;
  }

  /** Would charging `bytes` exceed the limit? Charges nothing. */
  wouldExceed(bytes: number): boolean {
    return this.#used + bytes > this.limit;
  }

  /** Charge `bytes`, or throw `ERR_BUDGET_EXCEEDED` and charge nothing. */
  charge(bytes: number, entry: string): void {
    if (this.wouldExceed(bytes)) {
      throw new OpcError(
        'ERR_BUDGET_EXCEEDED',
        'Reading ' +
          entry +
          ' would inflate ' +
          String(this.#used + bytes) +
          ' bytes in total, past this package\u2019s ' +
          String(this.limit) +
          '-byte decompression budget. The archive is either far larger than a ' +
          'presentation should be, or a decompression bomb.',
        { entry, limit: this.limit, actual: this.#used + bytes },
      );
    }
    this.#used += bytes;
  }
}

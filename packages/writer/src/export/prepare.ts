import type { PartStore } from '@pptx-studio/opc';
import { WriterError } from '../errors.js';

/**
 * Work that has to happen at save time, contributed by the sub-phase that owns it - 8.7's fonts,
 * 3.4's autofit, 10.8's media - so the writer knows none of it. Hooks run before collection, which
 * they change, and before validation, which must check the markup they write. ADR 0011.
 */

export interface PrepareContext {
  /** The package being exported. Hooks mutate this. */
  readonly store: PartStore;
  /** The package as it was opened, when there is one. Read only; never mutate it. */
  readonly baseline: PartStore | null;
  /**
   * Record something worth reporting.
   *
   * Surfaced on `ExportResult.prepared`. A hook that did nothing should say
   * nothing - the common case for every hook on most exports is silence, and a
   * note per hook per save would bury the one that mattered.
   */
  note(message: string): void;
}

export interface PrepareHook {
  /** Short and stable. It appears in `ExportResult` and in the error if this hook throws. */
  readonly name: string;
  run(context: PrepareContext): void;
}

export interface PrepareRecord {
  readonly hook: string;
  readonly notes: readonly string[];
}

/**
 * Run the hooks in the order given.
 *
 * The order is the caller's, not ours, and it is not inferred from anything.
 * Some pairs genuinely depend on each other - autofit must settle before a font
 * pass decides which typefaces are used - and a writer that sorted hooks by
 * some notion of priority would be guessing at a dependency the caller knows
 * for certain.
 *
 * A hook that throws stops the export. That is the point: the alternative is
 * handing over a package with five of the six font artifacts in it, which
 * PowerPoint reports as a problem with the content and no way to find out
 * which.
 */
export function runPrepare(
  hooks: readonly PrepareHook[],
  store: PartStore,
  baseline: PartStore | null,
): PrepareRecord[] {
  const records: PrepareRecord[] = [];
  for (const hook of hooks) {
    const notes: string[] = [];
    const context: PrepareContext = {
      store,
      baseline,
      note: (message: string) => notes.push(message),
    };
    try {
      hook.run(context);
    } catch (error) {
      throw new WriterError(
        'ERR_PREPARE_FAILED',
        'the prepare hook ' +
          hook.name +
          ' threw, so nothing was written: ' +
          (error instanceof Error ? error.message : String(error)),
        { hook: hook.name },
        { cause: error },
      );
    }
    if (notes.length > 0) records.push({ hook: hook.name, notes });
  }
  return records;
}

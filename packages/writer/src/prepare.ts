import type { PartStore } from '@pptx-studio/opc';
import { WriterError } from './errors.js';

/**
 * Work that has to happen at save time, contributed by whoever owns it.
 *
 * Several later sub-phases need to touch the package on the way out, and each
 * of them is somebody else's subject:
 *
 * - **8.7, embedded fonts.** Six artifacts that must all be present or
 *   PowerPoint reports a problem: the `fntdata` parts, the `fntdata` content-type
 *   `Default`, the font relationships on `presentation.xml`, the
 *   `p:embeddedFont` entries, their PANOSE and charset attributes, and
 *   `@embedTrueTypeFonts`. Plus a collection pass, because PowerPoint requires
 *   every listed typeface to actually be used.
 * - **3.4, autofit.** A `txBody` whose text changed needs its computed
 *   `@fontScale` and `@lnSpcReduction` written back - and one whose text did
 *   not must be left exactly alone, or a view-only deck stops round-tripping.
 * - **10.8, media.** The `p14:media` and `a:videoFile` relationships are a
 *   deliberate pair in the file and a duplicate on the way out.
 *
 * None of that is the writer's subject. Written inline it would become a list
 * of special cases in `export.ts` that grows by one every phase, that the
 * writer's own tests would have to know about, and that could not be tested
 * without standing up a font stack or a text engine. As hooks, each lands in
 * the package that owns it, ships with the sub-phase that needs it, and this
 * package keeps knowing nothing about fonts.
 *
 * ## Where they run, and why there
 *
 * Before collection and before validation, both deliberately.
 *
 * Before collection, because a hook is exactly the thing that changes what is
 * referenced - 8.7 both adds font parts and drops the ones nothing uses - and a
 * sweep that ran first would be answering a question about the wrong package.
 *
 * Before validation, because a hook writes markup, and markup this session
 * wrote is precisely what the twenty-nine rules exist to check. A hook that ran
 * after validation would be the one part of an export nothing checked, which is
 * the opposite of how it should be: it is newly synthesized markup, the
 * riskiest kind there is.
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

import {
  CONTENT_TYPES_PART,
  normalizePartName,
  PartStore,
  type ReadZipOptions,
} from '@pptx-studio/opc';
import { assertValid, type Report, type RuleId } from '@pptx-studio/validate';
import { planCollection, type CollectionPlan, type CollectionPolicy } from '../gc/collect.js';
import { assertPreserved, type PreservationCheck } from './preserve.js';
import { runPrepare, type PrepareHook, type PrepareRecord } from './prepare.js';

/**
 * Handing the package back: prepare hooks, collect, write, then check that nothing unedited
 * changed and that the rules pass. The check judges the emitted archive, so bytes are never
 * returned unjudged. ADR 0011.
 */

export interface ExportOptions {
  /** The package to write, with whatever edits have been made to it. */
  readonly store: PartStore;
  /**
   * The package as it was opened.
   *
   * Optional, and the three things it costs to leave out are worth knowing
   * before deciding to. Without it the three preservation rules are skipped
   * rather than passed; the media sweep collects nothing, because there is no
   * way to tell what this session orphaned; and `origin` is unavailable, so
   * every fatal finding counts as one we introduced. `openPackage` produces a
   * correct one in a line.
   */
  readonly baseline?: PartStore;
  /** The archive the baseline was opened from. Needed by `V003` and the preservation check. */
  readonly baselineBytes?: Uint8Array;
  /** Run in the order given, before everything else. See `prepare.ts`. */
  readonly prepare?: readonly PrepareHook[];
  /** Media garbage collection. Defaults to `orphaned-here`. See `collect.ts`. */
  readonly collect?: CollectionPolicy;
  /** Which parts the sweep may touch. Defaults to `/ppt/media/`. */
  readonly sweepable?: (partName: string) => boolean;
  /**
   * Set false to write without the firewall.
   *
   * There is no legitimate production use. It exists so that a test can build a
   * deliberately broken package and look at it, and so that `cli bisect` in 1.5
   * can emit the intermediate packages whose whole purpose is to be rejected.
   */
  readonly validate?: boolean;
  /** A subset of rules to run. Defaults to every rule. */
  readonly rules?: readonly RuleId[];
  /** Set false to skip comparing the output against the source archive. */
  readonly verifyPreservation?: boolean;
  /** Passed through to `PartStore.write`. */
  readonly normalizeEntryOrder?: boolean;
  /** Passed through to `PartStore.write`. */
  readonly deflateLevel?: number;
  /** Passed to `readZip` wherever this package re-opens an archive. */
  readonly zip?: ReadZipOptions;
}

export interface ExportResult {
  /** The package. Only ever returned when every check above passed. */
  readonly bytes: Uint8Array;
  /** `null` when validation was turned off. */
  readonly report: Report | null;
  /** Hooks that had something to say. Silent hooks are not listed. */
  readonly prepared: readonly PrepareRecord[];
  readonly collection: CollectionPlan;
  /** Parts serialised afresh. Empty on a no-op export - that is the point. */
  readonly rewritten: readonly string[];
  /** Parts streamed out of the source archive without being decompressed. */
  readonly streamed: number;
  readonly preservation: PreservationCheck;
}

/**
 * A package and the untouched twin every check compares it against.
 *
 * Two independent `PartStore.open` calls over the same bytes. Independent
 * matters: the baseline has to keep answering questions about the file as it
 * arrived no matter what happens to the other one, and sharing anything mutable
 * between them would make it a record of the present rather than the past.
 *
 * Neither call inflates anything but `[Content_Types].xml`, so the second is
 * two central-directory parses and a small XML read. They do carry separate
 * inflation budgets, which means an adversarial archive gets two budgets rather
 * than one; that is the correct trade at the size these budgets are set to, and
 * it is written down here so it is a decision rather than an oversight.
 */
export interface OpenPackage {
  readonly store: PartStore;
  readonly baseline: PartStore;
  readonly baselineBytes: Uint8Array;
}

/** Open a package for editing, with the baseline an export will need. */
export function openPackage(bytes: Uint8Array, options: ReadZipOptions = {}): OpenPackage {
  return {
    store: PartStore.open(bytes, options),
    baseline: PartStore.open(bytes, options),
    baselineBytes: bytes,
  };
}

export function exportPackage(options: ExportOptions): ExportResult {
  const store = options.store;
  const baseline = options.baseline ?? null;

  const prepared = runPrepare(options.prepare ?? [], store, baseline);

  const collection = planCollection({
    store,
    baseline,
    ...(options.collect === undefined ? {} : { policy: options.collect }),
    ...(options.sweepable === undefined ? {} : { sweepable: options.sweepable }),
  });

  // Settle the relationship parts before anything reads bytes.
  //
  // `relationships(...)` hands out a live object, so an edge added or removed
  // this session lives in the parsed form and not in the part's bytes until
  // this call. `PartStore.write` would do it a moment later anyway - but the
  // firewall below reads relationship markup as raw XML, deliberately, and
  // would otherwise be checking the `.rels` as it arrived rather than as it is
  // about to be written. The rules would then report an edge we removed as
  // still present, miss one we added, and not see a `.rels` created this
  // session at all. Which is precisely what happened, in the test that exports
  // a deck with one image unlinked.
  store.materializeRelationships();

  const rewritten = store.rewrittenParts();
  const streamed = store.partNames.length - rewritten.length;

  const bytes = store.write({
    ...(options.normalizeEntryOrder === undefined
      ? {}
      : { normalizeEntryOrder: options.normalizeEntryOrder }),
    ...(options.deflateLevel === undefined ? {} : { deflateLevel: options.deflateLevel }),
  });

  const preservation =
    options.verifyPreservation === false
      ? { checked: 0, rewritten: rewritten.length, skipped: 'turned off by the caller.' }
      : assertPreserved(bytes, options.baselineBytes, rewrittenKeys(store, rewritten), options.zip);

  const report =
    options.validate === false
      ? null
      : assertValid({
          store,
          bytes,
          ...(baseline === null ? {} : { baseline }),
          ...(options.baselineBytes === undefined ? {} : { baselineBytes: options.baselineBytes }),
          ...(options.rules === undefined ? {} : { rules: options.rules }),
          ...(options.zip === undefined ? {} : { zip: options.zip }),
        });

  return { bytes, report, prepared, collection, rewritten, streamed, preservation };
}

/**
 * The parts the preservation check must not compare, in its own key space.
 *
 * `[Content_Types].xml` is in the set whenever the map changed, and only then:
 * it is passed through untouched when nothing altered it, which is what lets a
 * deck that spells it `[content_types].xml` round-trip its original bytes under
 * the canonical entry name.
 */
function rewrittenKeys(store: PartStore, rewritten: readonly string[]): ReadonlySet<string> {
  const keys = new Set(rewritten.map((name) => normalizePartName(name)));
  if (store.contentTypes.dirty) keys.add(CONTENT_TYPES_PART);
  return keys;
}

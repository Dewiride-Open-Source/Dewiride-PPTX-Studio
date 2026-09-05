import { normalizePartName, type PartStore } from '@pptx-studio/opc';
import { WriterError } from '../errors.js';
import { orphanedParts, reachableParts, type Reachability } from './reachability.js';

/**
 * Mark and sweep, for media.
 *
 * ## What it collects, and why so little
 *
 * The sweep set is `/ppt/media/` and nothing else. Not layouts, not masters,
 * not themes, not `ppt/embeddings/*.bin`, whatever the relationship graph says
 * about them. PowerPoint keeps slide layouts that no slide uses - that is how
 * the layout picker has anything to offer - so a package where an unused layout
 * has been tidied away is a package where "change layout" has quietly lost
 * options the user had a moment ago. There is no upside to weigh against that:
 * the layout part is a few kilobytes and the media is the megabytes.
 *
 * ## Where the walk is rooted is the same rule, made twice
 *
 * The obvious implementation asks "which images do the slides use?" and roots
 * the walk at the slide list. It is wrong, and wrong in the way that only shows
 * up weeks later: a layout no slide currently uses still has its background
 * picture, and switching a slide onto that layout has to still work. Rooting at
 * the package root instead gets this right with no special case - layouts hang
 * off the master, the master off `presentation.xml` - so every image a layout
 * or master references is live whether or not a slide has ever used it.
 *
 * ## It cannot break a relationship
 *
 * Worth stating as a property rather than a hope, because it is what makes this
 * safe to run without supervision. A part is collected only when **no**
 * relationship anywhere in the package resolves to it. So there is no edge left
 * to dangle - and the case that kills naive implementations, two slides sharing
 * one image with one of them deleted, cannot arise: the surviving slide's
 * relationship still resolves and still marks the image live.
 *
 * The corollary is a division of labour. This collects nothing until something
 * else has removed the relationship, and removing it is part of deleting a
 * picture, which is a document operation and belongs to Phase 5's command
 * layer rather than to the writer.
 *
 * ## By default it collects only what we orphaned
 *
 * The question `@pptx-studio/validate` asks about findings, asked about parts:
 * not "is this unreferenced" but "did *we* leave it unreferenced". A deck can
 * arrive with an orphaned image - decks that have been through several editors
 * often do - and collecting it would mean that merely opening and saving a file
 * changes it. That breaks the property the whole of Phase 1 exists to
 * establish, and it breaks it in the direction where nobody notices for months.
 *
 * So the default policy differences the reachability of the package we are
 * about to write against the package as it was opened, and collects only what
 * moved from reachable to unreachable. With no baseline there is no "we", and
 * it collects nothing.
 *
 * Measured, for scale: across the fifty-two committed decks there are 1474
 * parts, 24 of them under `/ppt/media/`, and every part of every deck is
 * reachable from the package root. On this corpus the two policies are the same
 * policy. The difference is a promise about the decks we have not seen.
 */

export type CollectionPolicy =
  /** Only parts this session left unreferenced. The default. */
  | 'orphaned-here'
  /** Every unreferenced part in the sweep set, however it got that way. */
  | 'every-orphan'
  /** Collect nothing. */
  | 'none';

/** `/ppt/media/...` - the default sweep set. */
export function isMediaPart(partName: string): boolean {
  return /^\/ppt\/media\//i.test(partName);
}

export interface CollectOptions {
  /** The package about to be written. */
  readonly store: PartStore;
  /** The package as it was opened. Required by `orphaned-here`, ignored otherwise. */
  readonly baseline?: PartStore | null;
  /** Defaults to `orphaned-here`. */
  readonly policy?: CollectionPolicy;
  /**
   * Which parts the sweep may touch at all. Defaults to `isMediaPart`.
   *
   * Sub-phase 8.7 will pass one of its own, to collect a `ppt/fonts/*.fntdata`
   * whose typeface no run uses any more - PowerPoint requires every listed
   * embedded font to be in use. Anything wider than that wants the header above
   * read first.
   */
  readonly sweepable?: (partName: string) => boolean;
}

/** An orphan the sweep declined to collect, and why. */
export interface KeptPart {
  readonly part: string;
  readonly why: string;
}

export interface CollectionPlan {
  /** Parts removed, in the order they were found. */
  readonly collect: readonly string[];
  /** Orphans left in place, each with a reason. Empty is the normal case. */
  readonly kept: readonly KeptPart[];
  /**
   * False when the relationship graph could not be walked completely.
   *
   * A caller that deletes on the strength of an incomplete walk is deleting on
   * the strength of a parse error.
   */
  readonly safe: boolean;
  /** Relationship parts that would not parse. Empty when `safe`. */
  readonly blockedBy: readonly string[];
}

/**
 * Plan and apply in one pass.
 *
 * Not split into a pure plan and a separate apply, and the reason is the fixed
 * point below: removing a part can orphan the parts *it* referenced, and the
 * only way to observe that is to remove and look again. A pure planner would
 * have to simulate the store to find the second round, which is a second
 * implementation of the store.
 *
 * Never throws. A package too broken to reason about yields a plan that
 * collects nothing and says why - refusing would make a deck with one
 * unparseable `.rels` permanently unsaveable, and that `.rels` may well have
 * arrived that way. `collectGarbage` is the variant that refuses, for callers
 * who asked for a sweep and would misread silence as "nothing to do".
 */
export function planCollection(options: CollectOptions): CollectionPlan {
  const policy = options.policy ?? 'orphaned-here';
  if (policy === 'none') return { collect: [], kept: [], safe: true, blockedBy: [] };

  const sweepable = options.sweepable ?? isMediaPart;
  const store = options.store;

  let before: Reachability | null = null;
  if (policy === 'orphaned-here') {
    const baseline = options.baseline ?? null;
    if (baseline === null) {
      const now = reachableParts(store);
      return {
        collect: [],
        kept: candidateOrphans(store, now, sweepable).map((part) => ({
          part,
          why:
            'it is unreferenced, but with no baseline to compare against there is no way to tell ' +
            'whether this session orphaned it or it arrived that way.',
        })),
        safe: now.complete,
        blockedBy: now.blockedBy,
      };
    }
    before = reachableParts(baseline);
    if (!before.complete)
      return { collect: [], kept: [], safe: false, blockedBy: before.blockedBy };
  }

  const collect: string[] = [];
  const kept: KeptPart[] = [];
  const decided = new Set<string>();

  // Runs to a fixed point rather than once, because a collected part can own a
  // `.rels` whose targets then become unreferenced. It terminates: every round
  // that continues removes at least one part from a finite set. On the default
  // sweep set the second round always finds nothing - media parts have no
  // relationships - so the loop costs one extra walk and buys correctness for
  // the sweep sets later sub-phases will pass.
  for (;;) {
    const now = reachableParts(store);
    if (!now.complete) {
      // On the first round nothing has been removed and the empty plan is the
      // honest answer. On a later round `collect` holds work already applied,
      // which is reported as done rather than silently dropped.
      return { collect, kept, safe: false, blockedBy: now.blockedBy };
    }

    const found = candidateOrphans(store, now, sweepable).filter(
      (part) => !decided.has(normalizePartName(part)),
    );
    if (found.length === 0) break;

    const removing: string[] = [];
    for (const part of found) {
      decided.add(normalizePartName(part));
      if (before !== null && !before.reachable.has(normalizePartName(part))) {
        kept.push({
          part,
          why: 'it was already unreferenced when this package was opened, so it is not ours to remove.',
        });
        continue;
      }
      removing.push(part);
    }
    if (removing.length === 0) break;

    for (const part of removing) {
      store.removePart(part);
      collect.push(part);
    }
  }

  return { collect, kept, safe: true, blockedBy: [] };
}

function candidateOrphans(
  store: PartStore,
  walk: Reachability,
  sweepable: (partName: string) => boolean,
): string[] {
  if (!walk.complete) return [];
  return orphanedParts(store, walk).filter(sweepable);
}

/**
 * Collect, and refuse rather than guess.
 *
 * The entry point for a deliberate clean-up: a user asking for one, or a
 * command that knows it has just orphaned something. It throws when the graph
 * could not be walked completely, because a caller who asked for a sweep and
 * got silence would reasonably conclude there had been nothing to sweep.
 *
 * `exportPackage` deliberately does not use this. An export that failed because
 * some unrelated `.rels` will not parse would be refusing to save a file over a
 * defect it did not cause and cannot fix.
 */
export function collectGarbage(options: CollectOptions): CollectionPlan {
  const plan = planCollection(options);
  if (plan.safe) return plan;
  throw new WriterError(
    'ERR_COLLECTION_UNSAFE',
    'refusing to collect: ' +
      String(plan.blockedBy.length) +
      ' relationship part(s) would not parse, so no part in this package can be shown to be ' +
      'unreferenced. First: ' +
      (plan.blockedBy[0] ?? '(unknown)') +
      '. An unreadable edge is an invisible edge, and an invisible edge makes its target look ' +
      'like garbage.',
    { blockedBy: plan.blockedBy },
  );
}

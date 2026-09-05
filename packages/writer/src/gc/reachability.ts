import {
  normalizePartName,
  relsPartNameFor,
  ROOT_RELS_PART,
  type PartStore,
  type Relationships,
} from '@pptx-studio/opc';

/**
 * Which parts anything can still get to, and - the part that matters - whether
 * we are sure.
 *
 * `@pptx-studio/census` has a walk of its own and this is deliberately not it.
 * The two answer the same question with opposite failure modes, and the census
 * says so in its own comment: it swallows a relationship part it cannot parse
 * and carries on, because a census must never refuse. That is exactly right for
 * a report and exactly wrong here. A swallowed parse error means some edges are
 * invisible, invisible edges make their targets look unreferenced, and
 * unreferenced is what this walk is used to justify deleting. The census's
 * worst case is an inaccurate line of output; ours is a user's image gone.
 *
 * So every failure is recorded rather than skipped, and `complete` is false if
 * there was one. `collect.ts` does nothing at all unless it is true.
 *
 * ## Relationship parts are reachable when their source is
 *
 * Nothing ever *relates* to a `.rels` - the spec forbids it and PowerPoint
 * refuses a package that tries, which is one of the assertions in `PartStore`.
 * They are found by naming convention instead. Walking edges alone would
 * therefore report every relationship part in the package as unreachable, which
 * is true of the graph and false of the container. They are marked alongside
 * their source, so that "unreachable" keeps meaning "nothing can find this".
 */

export interface Reachability {
  /** Normalised names of every part something can still get to, including `.rels` parts. */
  readonly reachable: ReadonlySet<string>;
  /**
   * False when at least one relationship part could not be read.
   *
   * A caller that deletes on the strength of an incomplete walk is deleting on
   * the strength of a parse error.
   */
  readonly complete: boolean;
  /** The relationship parts that would not parse, and the root if it was unusable. */
  readonly blockedBy: readonly string[];
}

/** Every part reachable by following relationships from the package root. */
export function reachableParts(store: PartStore): Reachability {
  const reachable = new Set<string>();
  const blockedBy: string[] = [];
  const queue: string[] = [];

  const mark = (partName: string): void => {
    const key = normalizePartName(partName);
    if (reachable.has(key)) return;
    reachable.add(key);
    queue.push(partName);
    // Its relationship part comes with it, if it has one.
    const rels = relsPartNameFor(partName);
    if (store.has(rels)) reachable.add(normalizePartName(rels));
  };

  const follow = (rels: Relationships): void => {
    for (const rel of rels.all) {
      if (rel.targetMode === 'External') continue;
      let target: string;
      try {
        target = rels.resolve(rel);
      } catch {
        // A target that will not resolve names no part, so it can neither mark
        // one nor hide one. `V008` reports it; there is nothing to do here.
        continue;
      }
      if (store.has(target)) mark(target);
    }
  };

  let root: Relationships;
  try {
    root = store.rootRelationships();
  } catch {
    return { reachable, complete: false, blockedBy: [ROOT_RELS_PART] };
  }
  if (root.size === 0) {
    // Not a package we can reason about. Every part in it would look orphaned,
    // which is the one wrong answer with consequences. `PartStore.write` refuses
    // such a package outright a moment later; we simply decline to act first.
    return { reachable, complete: false, blockedBy: [ROOT_RELS_PART] };
  }
  if (store.has(ROOT_RELS_PART)) reachable.add(normalizePartName(ROOT_RELS_PART));
  follow(root);

  // Every reached part is asked for its relationships - including the ones with
  // no `.rels` part in the package.
  //
  // The census restricts this to parts that own a relationship part, and on a
  // package being read that is a free optimisation. Here it is a bug, and a
  // quiet one: a relationship collection created during this session exists
  // only as a parsed object until `PartStore.write` materialises it, so its
  // `.rels` is not among `partNames` yet. Keying the walk on existing `.rels`
  // parts therefore makes exactly the edges this session added invisible - and
  // those are the edges a writer is about. It cost one failing test to find,
  // and the optimisation it replaces was buying a `Map` lookup per part.
  for (let part = queue.pop(); part !== undefined; part = queue.pop()) {
    try {
      follow(store.relationships(part));
    } catch {
      blockedBy.push(relsPartNameFor(part));
    }
  }

  return { reachable, complete: blockedBy.length === 0, blockedBy };
}

/**
 * Parts nothing in the package can get to.
 *
 * Relationship parts are excluded by construction rather than by filter - see
 * the header. `[Content_Types].xml` never appears because it is not a part.
 */
export function orphanedParts(store: PartStore, walk: Reachability): string[] {
  return store.partNames.filter((name) => !walk.reachable.has(normalizePartName(name)));
}

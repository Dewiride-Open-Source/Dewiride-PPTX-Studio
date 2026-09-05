import {
  deflatedEntry,
  passthroughEntry,
  readZip,
  writeZip,
  type ReadZipOptions,
  type ZipArchive,
  type ZipEntry,
  type ZipEntryInput,
} from '@pptx-studio/opc';
import {
  encodeXmlSource,
  parseXml,
  type XAttribute,
  type XElement,
  type XNode,
} from '@pptx-studio/xml';

/**
 * Narrowing a package PowerPoint refuses down to the change that causes it.
 *
 * ## Why this is the debugger for this project
 *
 * PowerPoint emits no diagnostic log. A file it will not open produces one
 * sentence - "PowerPoint found a problem with content in deck.pptx" - naming no
 * part, no element and no reason, and a file it *repairs* produces a sentence
 * that does not even say what was repaired. `@pptx-studio/validate` covers the
 * twenty-nine failures we know how to describe; this covers the rest, which is
 * the interesting ones. The plan's own caveat: passing the rules is necessary
 * and not sufficient, because PowerPoint rejects some schema-legal markup for
 * reasons nobody has written down.
 *
 * So the only way to find out what a refusal is about is to ask PowerPoint
 * again with less of the change present, and keep asking. That is delta
 * debugging: Zeller and Hildebrandt's `ddmin` over a set of changes
 * (*Simplifying and Isolating Failure-Inducing Input*, IEEE TSE 2002), applied
 * level by level down a tree, which is Misherghi and Su's HDD (*HDD:
 * Hierarchical Delta Debugging*, ICSE 2006).
 *
 * ## This is the isolation problem, not the simplification one
 *
 * The distinction is Zeller's and it decides the whole design. Simplification
 * starts from one failing input and cuts it down; it has to guess what a
 * smaller input looks like, and most of its guesses are not even well formed.
 * Isolation starts from a **passing** configuration and a **failing** one and
 * reduces the *difference* between them.
 *
 * We are always in the second case. A deck that PowerPoint repairs came from a
 * deck that it opened, and we have both. So the atoms here are not elements of
 * the document: they are the **changes** between the two packages, and a
 * configuration is a subset of them applied to the original. That buys three
 * things a simplifier cannot have:
 *
 *   - Every configuration is well formed by construction. The splice unit is a
 *     whole node replaced by the whole node the other package has in that
 *     position, so no configuration can invent unbalanced markup.
 *   - The empty configuration is known to pass and the full one to fail, which
 *     is exactly `ddmin`'s precondition - and both are *checked* here rather
 *     than assumed, because when they do not hold the reason is the most useful
 *     thing this command can say.
 *   - The answer is "these two changes, out of ninety-one", which is a sentence
 *     about your edit rather than about PowerPoint's file format.
 *
 * ## Entries, not parts
 *
 * The delta is taken over **ZIP entries**, not over `PartStore` parts, and the
 * difference is not academic: `[Content_Types].xml` is not a part. It is a
 * package-level stream, `PartStore.partNames` deliberately excludes it, and a
 * missing `<Default Extension="fntdata"/>` in it is the canonical cause of
 * "PowerPoint found a problem with content" - the one this project's own plan
 * cites. A bisector that could only vary parts would be blind to the single
 * best-documented repair prompt there is.
 *
 * Working at the entry level also means nothing has to know what an entry
 * holds. Anything that parses as XML is descended into - parts, `.rels`,
 * `docProps` and the content-type stream alike - and anything that does not is
 * one atom, whether it is a PNG, an OLE2 compound file, or markup too damaged
 * to read. There is no content-type table to keep in step.
 */

// -------------------------------------------------------------------- the delta

export interface Span {
  readonly start: number;
  readonly end: number;
}

export type ChangeKind =
  /** The entry exists only in the broken package. */
  | 'entry-added'
  /** The entry exists only in the original. */
  | 'entry-removed'
  /** A whole entry. Binary, or markup whose two sides do not line up. */
  | 'entry'
  /** One element and everything under it. */
  | 'element'
  /** One text, CDATA, comment, PI or declaration node. */
  | 'node'
  /**
   * A run of children replaced by another run of a different length.
   *
   * What an insertion or a deletion looks like. `was` empty means children were
   * added; `text` empty means they were taken away.
   */
  | 'children'
  /** One element's start tag: its name, its attributes and their spacing. */
  | 'tag'
  /** One attribute, including the whitespace in front of it. */
  | 'attribute';

export interface Change {
  readonly id: number;
  /** ZIP entry name, without a leading slash: `ppt/slides/slide1.xml`. */
  readonly entry: string;
  readonly kind: ChangeKind;
  /**
   * Where in the entry, for a person to read.
   *
   * An XPath-ish path for markup - `/p:sld/p:cSld/p:spTree/p:sp[3]/@name` -
   * built from qualified names exactly as the file writes them, because a path
   * that renamed prefixes could not be pasted back into a search.
   */
  readonly where: string;
  /** Distance from the entry, so a report can indent. */
  readonly depth: number;
  /** Where this sits in the **original** entry's text. Null for a whole entry. */
  readonly span: Span | null;
  /** What the broken package has there. Null when the entry is only original. */
  readonly text: string | null;
  /** What the original has there. Null when the entry is only in the broken one. */
  readonly was: string | null;
  /**
   * Finer changes that together reproduce this one exactly.
   *
   * The exactness is what makes descending free: the regions between the
   * children are the regions that compared equal, so applying every child is
   * the same package as applying the parent. HDD relies on it to move down a
   * level without spending a run to re-establish that the configuration fails.
   */
  readonly children: readonly Change[];
}

/** Every change in the forest, parents before children. */
export function flattenChanges(changes: readonly Change[]): Change[] {
  const out: Change[] = [];
  const stack = [...changes].reverse();
  while (stack.length > 0) {
    const change = stack.pop()!;
    out.push(change);
    for (let i = change.children.length - 1; i >= 0; i -= 1) stack.push(change.children[i]!);
  }
  return out;
}

/** The leaves: the finest changes this delta can express. */
export function leafChanges(changes: readonly Change[]): Change[] {
  return flattenChanges(changes).filter((change) => change.children.length === 0);
}

// ------------------------------------------------------------------- the oracle

/**
 * What the oracle says about one candidate package.
 *
 * Three values rather than two, and the third is Zeller's. `unresolved` is a
 * run that answered neither question - PowerPoint timed out, the harness would
 * not start, a subprocess was killed. Treating it as "passes" would let the
 * reducer discard the change that matters and then blame an innocent one;
 * treating it as "fails" would let it keep everything. It is counted and
 * reported instead, and never reduced towards.
 */
export type Verdict = 'fails' | 'passes' | 'unresolved';

/**
 * Does this package still show the problem?
 *
 * Synchronous on purpose. Every oracle we have is either pure computation - the
 * twenty-nine rules, the round-trip comparison - or a subprocess the caller
 * blocks on anyway, and `spawnSync` blocks perfectly well. Making the reducer
 * asynchronous to serve a caller that does not exist would put a Promise in the
 * middle of a browser package for nothing.
 */
export type Oracle = (bytes: Uint8Array, label: string) => Verdict;

export interface RunEvent {
  readonly run: number;
  readonly applied: number;
  readonly total: number;
  readonly label: string;
}

export interface BisectOptions {
  readonly oracle: Oracle;
  /**
   * Ceiling on oracle runs.
   *
   * `ddmin` is quadratic in the worst case and the interesting oracle takes
   * about a second and a half, so a run against a large delta can take a long
   * time. When the ceiling is hit the result carries `exhausted: true` and the
   * smallest failing configuration found so far - never a truncated answer
   * presented as a minimal one.
   */
  readonly maxRuns?: number;
  /** Called before each oracle run, for a progress line. */
  readonly onRun?: (run: RunEvent) => void;
  readonly zip?: ReadZipOptions;
  /**
   * Ceiling on the size of the delta tree.
   *
   * Two packages that share nothing produce a tree the size of the document.
   * Past this, subtrees stop being decomposed and stay atomic, which costs
   * precision and not correctness; `truncated` says when it happened.
   */
  readonly maxChanges?: number;
}

export type BisectOutcome =
  /** A minimal failing set was found. */
  | 'localized'
  /** The two packages hold the same entries, byte for byte. */
  | 'identical'
  /** The broken package passes the oracle. There is nothing to look for. */
  | 'broken-passes'
  /** The original fails too, so the cause is not in the difference. */
  | 'original-fails';

export interface BisectResult {
  readonly outcome: BisectOutcome;
  /** The whole delta, as a forest. One root per differing entry. */
  readonly changes: readonly Change[];
  /** The 1-minimal failing subset: drop any one of these and it passes. */
  readonly minimal: readonly Change[];
  /** The smallest package still showing the problem. Ready to open, or to keep. */
  readonly bytes: Uint8Array;
  /** Oracle runs actually performed. */
  readonly runs: number;
  /** Configurations answered from the cache instead of the oracle. */
  readonly cached: number;
  /** Runs that answered neither question. */
  readonly unresolved: number;
  /** True when `maxRuns` stopped the reduction early. */
  readonly exhausted: boolean;
  /** True when `maxChanges` stopped the delta being decomposed further. */
  readonly truncated: boolean;
}

// --------------------------------------------------------------- building files

interface Entry {
  readonly archive: ZipArchive;
  readonly entry: ZipEntry;
}

interface Context {
  readonly order: readonly string[];
  readonly original: ReadonlyMap<string, Entry>;
  readonly broken: ReadonlyMap<string, Entry>;
  /** Decoded text of every original entry we descended into. */
  readonly text: ReadonlyMap<string, string>;
  readonly byEntry: ReadonlyMap<string, readonly Change[]>;
}

function index(archive: ZipArchive): Map<string, Entry> {
  const map = new Map<string, Entry>();
  for (const entry of archive.entries) {
    if (entry.isDirectory) continue;
    map.set(entry.name, { archive, entry });
  }
  return map;
}

function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * The package with exactly this subset of the changes applied to the original.
 *
 * Entries no change touches are passed through still compressed rather than
 * re-deflated - the same trick `PartStore.write` uses, and it matters more
 * here, because this runs once per oracle call rather than once per export.
 */
function buildPackage(context: Context, selected: ReadonlySet<number>): Uint8Array {
  const applied = new Map<string, Change[]>();
  for (const [name, changes] of context.byEntry) {
    const chosen = changes.filter((change) => selected.has(change.id));
    if (chosen.length > 0) applied.set(name, chosen);
  }

  const entries: ZipEntryInput[] = [];
  for (const name of context.order) {
    const source = context.original.get(name);
    const chosen = applied.get(name);

    if (chosen === undefined) {
      // Untouched by this configuration: whatever the original had, or nothing
      // at all if this entry exists only in the broken package.
      if (source !== undefined) entries.push(passthroughEntry(source.archive, source.entry));
      continue;
    }

    const whole = chosen.find((change) => change.span === null);
    if (whole !== undefined) {
      if (whole.kind === 'entry-removed') continue;
      const target = context.broken.get(name)!;
      entries.push(passthroughEntry(target.archive, target.entry));
      continue;
    }

    // Splice the broken text into the original at each selected span, back to
    // front, so an earlier span's offsets are still the offsets of the string
    // being edited. The selected set is always an antichain - HDD expands a
    // change into its children and never keeps both - so no two spans overlap.
    let text = context.text.get(name)!;
    const ordered = [...chosen].sort((a, b) => b.span!.start - a.span!.start);
    for (const change of ordered) {
      text = text.slice(0, change.span!.start) + change.text! + text.slice(change.span!.end);
    }
    entries.push(deflatedEntry(name, encodeXmlSource(text)));
  }

  return writeZip(entries);
}

// ---------------------------------------------------------------- the delta walk

interface Builder {
  readonly next: () => number;
  readonly budget: { left: number };
  readonly truncated: { yes: boolean };
}

function atom(
  builder: Builder,
  entry: string,
  kind: ChangeKind,
  where: string,
  depth: number,
  span: Span | null,
  was: string | null,
  text: string | null,
  children: readonly Change[] = [],
): Change {
  return { id: builder.next(), entry, kind, where, depth, span, text, was, children };
}

function slice(source: string, node: XNode): string {
  return source.slice(node.start, node.end);
}

function nameOf(node: XNode): string {
  switch (node.type) {
    case 'element':
      return node.qname;
    case 'text':
      return 'text()';
    case 'cdata':
      return 'cdata()';
    case 'comment':
      return 'comment()';
    case 'processingInstruction':
      return 'processing-instruction()';
    case 'declaration':
      return 'xml-declaration()';
  }
}

/** `/p:sld/p:cSld/p:spTree/p:sp[3]` - one-based, and indexed only where it must be. */
function pathTo(parent: string, siblings: readonly XNode[], at: number): string {
  const name = nameOf(siblings[at]!);
  let position = 1;
  let total = 0;
  for (let i = 0; i < siblings.length; i += 1) {
    if (nameOf(siblings[i]!) !== name) continue;
    total += 1;
    if (i < at) position += 1;
  }
  return parent + '/' + name + (total > 1 ? '[' + String(position) + ']' : '');
}

/**
 * Attributes paired by position, or null when they do not line up.
 *
 * Null whenever anything outside the attributes themselves differs - the
 * element name, the number or order of attributes, the whitespace before `/>`.
 * Otherwise the attribute spans would not cover the whole difference, and
 * reverting every one of them would not reproduce the original tag, which is
 * the single property the descent is allowed to assume.
 */
function pairAttributes(a: XElement, b: XElement): [XAttribute, XAttribute][] | null {
  if (a.attributes.length === 0 || a.attributes.length !== b.attributes.length) return null;
  // `<qname` on both sides, then the same run-up to the first attribute.
  if (a.nameEnd - a.start !== b.nameEnd - b.start) return null;
  if (a.attributes[0]!.start - a.start !== b.attributes[0]!.start - b.start) return null;
  // ...and the same tail: the spacing before `>` or `/>`, and nothing between
  // the last attribute and it.
  if (a.attributes.at(-1)!.end !== a.trailingSpaceStart) return null;
  if (b.attributes.at(-1)!.end !== b.trailingSpaceStart) return null;
  if (a.openTagEnd - a.trailingSpaceStart !== b.openTagEnd - b.trailingSpaceStart) return null;

  const pairs: [XAttribute, XAttribute][] = [];
  for (let i = 0; i < a.attributes.length; i += 1) {
    const left = a.attributes[i]!;
    const right = b.attributes[i]!;
    if (left.qname !== right.qname) return null;
    pairs.push([left, right]);
  }
  return pairs;
}

/** The start tag, decomposed into attributes when the two tags line up. */
function tagChange(
  builder: Builder,
  entry: string,
  where: string,
  depth: number,
  a: XElement,
  b: XElement,
  ao: string,
  bo: string,
): Change {
  const children: Change[] = [];
  const pairs = pairAttributes(a, b);
  if (pairs !== null) {
    for (const [left, right] of pairs) {
      const was = ao.slice(left.start, left.end);
      const text = bo.slice(right.start, right.end);
      if (was === text) continue;
      children.push(
        atom(
          builder,
          entry,
          'attribute',
          where + '/@' + left.qname,
          depth + 1,
          { start: left.start, end: left.end },
          was,
          text,
        ),
      );
    }
  }

  return atom(
    builder,
    entry,
    'tag',
    where,
    depth,
    { start: a.start, end: a.openTagEnd },
    ao.slice(a.start, a.openTagEnd),
    bo.slice(b.start, b.openTagEnd),
    children,
  );
}

/**
 * Line up two lists of children and describe the difference between them.
 *
 * The two lists are matched from both ends first, and only what is left in the
 * middle is treated as changed. That is what makes an insertion or a deletion
 * nameable: without it, any change to the *number* of children shifts every
 * position after it, nothing pairs, and the answer degrades to "something in
 * this element". The case that forced it is the headline one - a missing
 * `<Default Extension="fntdata"/>` reduces `<Types>` from seventeen children to
 * sixteen, and positional pairing reported the whole content-type stream.
 *
 * When the middles are the same length they pair off one to one, which is the
 * ordinary edit-in-place case. When they are not, the whole middle is one
 * change: a run of children replaced by another run. That is coarse if a part
 * had several separate insertions, and it is exact in the common case of one -
 * which is what a writer produces.
 */
function pairChildren(
  builder: Builder,
  entry: string,
  where: string,
  depth: number,
  a: XElement,
  b: XElement,
  ao: string,
  bo: string,
): Change[] {
  const left = a.children;
  const right = b.children;

  let lo = 0;
  while (lo < left.length && lo < right.length && slice(ao, left[lo]!) === slice(bo, right[lo]!)) {
    lo += 1;
  }
  let hiA = left.length;
  let hiB = right.length;
  while (hiA > lo && hiB > lo && slice(ao, left[hiA - 1]!) === slice(bo, right[hiB - 1]!)) {
    hiA -= 1;
    hiB -= 1;
  }

  if (hiA === lo && hiB === lo) return [];

  const changes: Change[] = [];
  if (hiA - lo === hiB - lo) {
    for (let i = lo; i < hiA; i += 1) {
      builder.budget.left -= 1;
      changes.push(
        pairNodes(builder, entry, pathTo(where, left, i), depth + 1, left[i]!, right[i]!, ao, bo),
      );
    }
    return changes;
  }

  // The middles are different lengths, so this is an insertion, a deletion, or
  // a replacement of one run by another. The span is the original's middle -
  // empty, and therefore zero-width, when children were only added.
  const start = lo < hiA ? left[lo]!.start : (left[hiA]?.start ?? a.closeTagStart);
  const end = lo < hiA ? left[hiA - 1]!.end : start;
  const text = lo < hiB ? bo.slice(right[lo]!.start, right[hiB - 1]!.end) : '';
  const at = lo < hiA ? pathTo(where, left, lo) : where;

  builder.budget.left -= 1;
  changes.push(
    atom(builder, entry, 'children', at, depth + 1, { start, end }, ao.slice(start, end), text),
  );
  return changes;
}

/** One change turning the original's node into the broken package's node. */
function pairNodes(
  builder: Builder,
  entry: string,
  where: string,
  depth: number,
  a: XNode,
  b: XNode,
  ao: string,
  bo: string,
): Change {
  const span = { start: a.start, end: a.end };
  const was = slice(ao, a);
  const text = slice(bo, b);

  const aligned =
    a.type === 'element' &&
    b.type === 'element' &&
    a.qname === b.qname &&
    a.selfClosing === b.selfClosing;

  if (!aligned) {
    return atom(
      builder,
      entry,
      a.type === 'element' ? 'element' : 'node',
      where,
      depth,
      span,
      was,
      text,
    );
  }

  if (builder.budget.left <= 0) {
    builder.truncated.yes = true;
    return atom(builder, entry, 'element', where, depth, span, was, text);
  }

  const element = a;
  const other = b;
  const children: Change[] = [];

  if (ao.slice(element.start, element.openTagEnd) !== bo.slice(other.start, other.openTagEnd)) {
    builder.budget.left -= 1;
    children.push(tagChange(builder, entry, where, depth + 1, element, other, ao, bo));
  }

  children.push(...pairChildren(builder, entry, where, depth, element, other, ao, bo));

  // Same name, same shape, same children, same start tag - and yet the texts
  // differ. Only the end tag is left and it cannot differ, so this is
  // unreachable. Returning an atom rather than a childless parent keeps it
  // harmless if the node model ever grows a region this walk does not cover.
  if (children.length === 0) return atom(builder, entry, 'element', where, depth, span, was, text);

  return atom(builder, entry, 'element', where, depth, span, was, text, children);
}

function deltaForEntry(
  builder: Builder,
  name: string,
  originalBytes: Uint8Array,
  brokenBytes: Uint8Array,
  text: Map<string, string>,
): Change {
  const whole = (): Change => atom(builder, name, 'entry', '/', 0, null, null, null);

  let a;
  let b;
  try {
    a = parseXml(originalBytes);
    b = parseXml(brokenBytes);
  } catch {
    // Not XML, or XML we cannot read. Either way this entry is one atom.
    return whole();
  }

  if (a.children.length !== b.children.length) return whole();

  const children: Change[] = [];
  for (let i = 0; i < a.children.length; i += 1) {
    const left = a.children[i]!;
    const right = b.children[i]!;
    if (slice(a.source, left) === slice(b.source, right)) continue;
    builder.budget.left -= 1;
    children.push(
      pairNodes(builder, name, pathTo('', a.children, i), 1, left, right, a.source, b.source),
    );
  }

  if (children.length === 0) return whole();
  text.set(name, a.source);
  return atom(builder, name, 'entry', '/', 0, null, null, null, children);
}

/** The forest of changes between two packages. One root per differing entry. */
export function collectDelta(
  original: ZipArchive,
  broken: ZipArchive,
  maxChanges = 20_000,
): { changes: Change[]; text: Map<string, string>; truncated: boolean } {
  let counter = 0;
  const builder: Builder = {
    next: () => counter++,
    budget: { left: maxChanges },
    truncated: { yes: false },
  };

  const left = index(original);
  const right = index(broken);
  const text = new Map<string, string>();
  const changes: Change[] = [];

  for (const [name, source] of left) {
    const target = right.get(name);
    if (target === undefined) {
      changes.push(atom(builder, name, 'entry-removed', '/', 0, null, null, null));
      continue;
    }
    const a = source.archive.read(source.entry);
    const b = target.archive.read(target.entry);
    if (equalBytes(a, b)) continue;
    changes.push(deltaForEntry(builder, name, a, b, text));
  }

  for (const name of right.keys()) {
    if (left.has(name)) continue;
    changes.push(atom(builder, name, 'entry-added', '/', 0, null, null, null));
  }

  return { changes, text, truncated: builder.truncated.yes };
}

// --------------------------------------------------------------------- the search

/** Internal. Never escapes `bisectPackages`; see the `maxRuns` handling there. */
class Exhausted extends Error {}

/**
 * `ddmin`, reducing a set of changes to a 1-minimal failing subset.
 *
 * 1-minimal means every change left in the answer is load-bearing: drop any one
 * and the package stops failing. It does **not** mean no smaller failing set
 * exists - two changes may only fail together, and neither alone - and the
 * difference is worth keeping straight, because a report claiming a minimality
 * it does not have sends someone looking in the wrong place.
 */
function ddmin(
  items: readonly Change[],
  test: (subset: readonly Change[]) => Verdict,
  /**
   * Called with every failing configuration on the way down.
   *
   * So that hitting `maxRuns` returns the smallest set reached rather than the
   * one this call started from. Without it a run that stopped one test short of
   * the answer reported the whole level, which is a true statement and a
   * useless one.
   */
  record: (current: readonly Change[]) => void = () => {},
): readonly Change[] {
  let current = items;
  let n = 2;

  while (current.length >= 2) {
    const size = Math.ceil(current.length / n);
    const subsets: (readonly Change[])[] = [];
    for (let i = 0; i < current.length; i += size) subsets.push(current.slice(i, i + size));

    // Reduce to a subset: is one nth of what is left enough on its own?
    const failing = subsets.find((subset) => test(subset) === 'fails');
    if (failing !== undefined) {
      current = failing;
      record(current);
      n = 2;
      continue;
    }

    // Reduce to a complement: can one nth be dropped?
    let dropped = false;
    for (const subset of subsets) {
      const complement = current.filter((change) => !subset.includes(change));
      if (complement.length === 0 || complement.length === current.length) continue;
      if (test(complement) !== 'fails') continue;
      current = complement;
      record(current);
      n = Math.max(n - 1, 2);
      dropped = true;
      break;
    }
    if (dropped) continue;

    if (n >= current.length) break;
    n = Math.min(current.length, n * 2);
  }

  return current;
}

export function bisectPackages(
  originalBytes: Uint8Array,
  brokenBytes: Uint8Array,
  options: BisectOptions,
): BisectResult {
  const original = readZip(originalBytes, options.zip ?? {});
  const broken = readZip(brokenBytes, options.zip ?? {});
  const { changes, text, truncated } = collectDelta(original, broken, options.maxChanges);

  const order: string[] = [];
  for (const entry of original.entries) if (!entry.isDirectory) order.push(entry.name);
  for (const entry of broken.entries) {
    if (entry.isDirectory || order.includes(entry.name)) continue;
    order.push(entry.name);
  }

  const byEntry = new Map<string, Change[]>();
  for (const change of flattenChanges(changes)) {
    const list = byEntry.get(change.entry);
    if (list === undefined) byEntry.set(change.entry, [change]);
    else list.push(change);
  }

  const context: Context = {
    order,
    original: index(original),
    broken: index(broken),
    text,
    byEntry,
  };

  const maxRuns = options.maxRuns ?? 2000;
  let runs = 0;
  let cached = 0;
  let unresolved = 0;
  let exhausted = false;
  const seen = new Map<string, Verdict>();

  const test = (subset: readonly Change[]): Verdict => {
    const ids = new Set(subset.map((change) => change.id));
    const key = [...ids].sort((a, b) => a - b).join(',');
    const remembered = seen.get(key);
    if (remembered !== undefined) {
      cached += 1;
      return remembered;
    }
    if (runs >= maxRuns) throw new Exhausted();
    runs += 1;
    const label = String(subset.length) + '/' + String(changes.length) + ' change(s)';
    options.onRun?.({ run: runs, applied: subset.length, total: changes.length, label });
    const verdict = options.oracle(buildPackage(context, ids), label);
    if (verdict === 'unresolved') unresolved += 1;
    seen.set(key, verdict);
    return verdict;
  };

  const done = (outcome: BisectOutcome, minimal: readonly Change[]): BisectResult => ({
    outcome,
    changes,
    minimal,
    bytes: buildPackage(context, new Set(minimal.map((change) => change.id))),
    runs,
    cached,
    unresolved,
    exhausted,
    truncated,
  });

  if (changes.length === 0) return done('identical', []);

  let best: readonly Change[] = changes;
  try {
    // `ddmin` assumes the empty configuration passes and the full one fails.
    // Both are checked, because when either does not hold, saying so is the
    // most useful thing this can do - and reducing anyway would produce a
    // confident answer about nothing.
    if (test([]) === 'fails') return done('original-fails', []);
    if (test(changes) !== 'fails') return done('broken-passes', changes);

    let level: readonly Change[] = changes;
    for (;;) {
      best = ddmin(level, test, (current) => {
        best = current;
      });
      const expanded = best.flatMap((change) =>
        change.children.length > 0 ? [...change.children] : [change],
      );
      if (expanded.length === best.length && expanded.every((c, i) => c === best[i])) break;
      // Applying every child of a change is the same package as applying the
      // change itself, so the next level is known to fail and `ddmin` may start
      // again without spending a run to establish it.
      level = expanded;
    }
  } catch (error) {
    if (!(error instanceof Exhausted)) throw error;
    exhausted = true;
  }

  return done('localized', best);
}

// -------------------------------------------------------------------- reporting

/** One line naming a change, without its text. */
export function describeChange(change: Change): string {
  const where = change.where === '/' ? '' : ' ' + change.where;
  switch (change.kind) {
    case 'entry-added':
      return 'added    /' + change.entry;
    case 'entry-removed':
      return 'removed  /' + change.entry;
    case 'entry':
      return 'entry    /' + change.entry;
    case 'element':
      return 'element  /' + change.entry + where;
    case 'node':
      return 'node     /' + change.entry + where;
    case 'children':
      // The two directions read very differently to someone debugging, so they
      // are named rather than both being called "children".
      if (change.was === '') return 'added    /' + change.entry + where;
      if (change.text === '') return 'removed  /' + change.entry + where;
      return 'children /' + change.entry + where;
    case 'tag':
      return 'tag      /' + change.entry + where;
    case 'attribute':
      return 'attr     /' + change.entry + where;
  }
}

/** `2 change(s) in 1 entry(s), from a delta of 91, in 37 oracle run(s)`. */
export function summarizeBisect(result: BisectResult): string {
  const entries = new Set(result.minimal.map((change) => change.entry));
  return (
    String(result.minimal.length) +
    ' change(s) in ' +
    String(entries.size) +
    ' entry(s), from a delta of ' +
    String(flattenChanges(result.changes).length) +
    ', in ' +
    String(result.runs) +
    ' oracle run(s)'
  );
}

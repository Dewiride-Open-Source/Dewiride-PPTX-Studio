import {
  isRelationshipPartName,
  isXmlContentType,
  normalizePartName,
  sha256Hex,
  sha256HexOfText,
  PartStore,
  sourcePartNameForRels,
  type Relationship,
  type Relationships,
} from '@pptx-studio/opc';
import {
  canonicalXml,
  excerpt,
  firstDifference,
  isXmlError,
  NS,
  parseXml,
  type XmlParseLimits,
} from '@pptx-studio/xml';
import { exportPackage, openPackage, type ExportOptions, type ExportResult } from './export.js';

/**
 * Is the deck we wrote the deck we read?
 *
 * ## Why this is not a byte comparison
 *
 * The plan is explicit that raw ZIP byte equality is the wrong test and says
 * why in one line: entry order, deflate level, DOS timestamps and attribute
 * order all legitimately differ, so a byte-equality round-trip test is red on
 * day one and disabled on day two. A test nobody trusts is worse than no test,
 * because it also stops anyone writing the one that would have worked.
 *
 * `assertPreserved` in this same package *does* compare bytes, and the two are
 * not in competition. It asks a narrow question - did the entries we said we
 * would not touch come out with the identical stored bytes - and it is the
 * right question for a no-op export. This module asks the wide one: for a
 * package we edited, or one a different producer wrote, do the two archives
 * hold the same document? That question has to survive an entry moving, a rId
 * being renumbered, and a `Default` becoming an `Override`, none of which
 * changes anything a reader can observe.
 *
 * ## Three comparisons, because a package has three kinds of content
 *
 * **XML parts** compare by canonical form - see `canonicalXml` in
 * `@pptx-studio/xml` for what it normalises and, more importantly, what it
 * refuses to. Whitespace and prefixes are meaning, not formatting.
 *
 * **Relationship parts** compare as a graph with the ids treated as opaque
 * labels. Two `.rels` that disagree only about whether the theme is `rId1` or
 * `rId7` describe the same package, and PowerPoint renumbers on every save, so
 * a comparator that could not see past that would be useless the first time it
 * was pointed at a file PowerPoint had written. Ids are matched by what they
 * point at, and the resulting mapping is then applied to the referring markup:
 * an `r:embed="rId3"` on one side and `r:embed="rId7"` on the other are the
 * same reference exactly when rId3 and rId7 resolve to the same part.
 *
 * That relabelling is keyed on the **namespace** of the attribute rather than
 * on a list of names. Every attribute in the relationships namespace is a
 * relationship reference - `r:id`, `r:embed`, `r:link`, `r:pict`, `r:dm`,
 * `r:lo`, `r:qs`, `r:cs` - and a list would be a list that is one entry short
 * the first time we meet a part type nobody thought about.
 *
 * **Everything else** compares by SHA-256. The bytes are in hand on both sides
 * so the digest is not needed to decide equality; it is needed to *report*, and
 * to give `cli bisect` in 1.5 something to name a part by. CRC-32, which this
 * package already has for the archive, is a 32-bit error-detecting code and not
 * evidence that two files are the same.
 *
 * ## What it does not compare, and why that is not a hole
 *
 * `[Content_Types].xml` is never compared as a document. It is a map, and the
 * same map has many spellings: a `Default` for an extension and an `Override`
 * on every part with that extension mean the identical thing, and PowerPoint
 * chooses between them differently from us. What is compared instead is the
 * resolved content type of every part, which is the only thing the map is for.
 */

export type DifferenceKind =
  /** A part is in the original and not in what was written. */
  | 'part-removed'
  /** A part is in what was written and not in the original. */
  | 'part-added'
  /** Both have the part; the content types disagree. */
  | 'content-type'
  /** Both have the part; the canonical XML disagrees. */
  | 'xml'
  /** Both have the part; the bytes disagree. */
  | 'binary'
  /** The relationship graphs of one source part disagree. */
  | 'relationship'
  /** A part could not be read on one side or the other, so nothing can be said. */
  | 'unreadable';

export interface Difference {
  readonly kind: DifferenceKind;
  /** The part it is about. For a relationship difference, the *source* part. */
  readonly part: string;
  readonly detail: string;
}

export type CompareHow = 'xml' | 'binary' | 'relationships';

export interface PartComparison {
  readonly part: string;
  readonly how: CompareHow;
  readonly same: boolean;
  /**
   * The digest of the compared form on the original side: the canonical XML for
   * an XML part, the bytes for anything else, and the relationship graph with
   * the ids left out for a `.rels`.
   *
   * Two parts with the same digest are the same part in the sense this module
   * means. That is the value worth writing into a report.
   */
  readonly digest: string;
  /** The same digest on the written side. Equal to `digest` whenever `same`. */
  readonly writtenDigest: string;
}

/** One relationship id that means the same thing under a different name. */
export interface Relabelling {
  /** The part whose `.rels` this is - `/` for the package root. */
  readonly source: string;
  readonly from: string;
  readonly to: string;
  /** What both ids point at, which is why they are the same label. */
  readonly target: string;
}

export interface RoundTripReport {
  readonly ok: boolean;
  readonly parts: readonly PartComparison[];
  readonly differences: readonly Difference[];
  /**
   * Ids that moved. Empty for anything this package wrote, and the interesting
   * field when the comparison is against a file PowerPoint saved.
   */
  readonly relabelled: readonly Relabelling[];
  readonly counts: {
    readonly xml: number;
    readonly binary: number;
    readonly relationships: number;
    readonly same: number;
  };
}

export interface CompareOptions {
  /** Passed to `parseXml` for both sides. */
  readonly limits?: XmlParseLimits;
}

/** The relationships namespace: every attribute in it is a relationship id. */
const R = NS.r;

/**
 * What a relationship points at, as a string that does not mention its id.
 *
 * An internal target resolves to an absolute part name, so that `../media/a.png`
 * from a slide and `/ppt/media/a.png` written absolutely are one target. An
 * external one is compared verbatim: it is a URI, we do not own it, and
 * normalising somebody's `http://Example.COM/` would be inventing a rule.
 */
function targetKey(rels: Relationships, rel: Relationship): string {
  if (rel.targetMode === 'External') return 'External ' + rel.target;
  try {
    return 'Internal ' + normalizePartName(rels.resolve(rel));
  } catch {
    // A target that will not resolve is still a target, and comparing the raw
    // string keeps two identically broken packages equal - which they are.
    return 'Internal? ' + rel.target;
  }
}

function relationshipKey(rels: Relationships, rel: Relationship): string {
  return rel.type + ' -> ' + targetKey(rels, rel);
}

interface GraphMatch {
  /** Original id to written id, for every relationship that matched. */
  readonly relabel: ReadonlyMap<string, string>;
  readonly relabelled: readonly Relabelling[];
  readonly differences: readonly Difference[];
  /** The graph as a sorted list of keys, ids omitted. The digest is taken over this. */
  readonly signature: string;
  readonly writtenSignature: string;
}

/**
 * Match two relationship collections by what they point at.
 *
 * Duplicates are legal - a part may relate to the same target twice under two
 * ids, and `a03-fills` in the corpus has eight fills on one image - so this
 * matches within a key group by document order rather than assuming one
 * relationship per key. Order inside a group is arbitrary in principle; it is
 * the only tiebreak available, and picking the wrong pairing among identical
 * relationships cannot change the answer, because the two ids resolve to the
 * same place.
 */
function matchGraphs(source: string, before: Relationships, after: Relationships): GraphMatch {
  const groupsBefore = new Map<string, Relationship[]>();
  const groupsAfter = new Map<string, Relationship[]>();
  const collect = (rels: Relationships, into: Map<string, Relationship[]>): string[] => {
    const keys: string[] = [];
    for (const rel of rels.all) {
      const key = relationshipKey(rels, rel);
      keys.push(key);
      const bucket = into.get(key);
      if (bucket === undefined) into.set(key, [rel]);
      else bucket.push(rel);
    }
    return keys.sort();
  };
  const signature = collect(before, groupsBefore).join('\n');
  const writtenSignature = collect(after, groupsAfter).join('\n');

  const relabel = new Map<string, string>();
  const relabelled: Relabelling[] = [];
  const differences: Difference[] = [];

  for (const [key, mine] of groupsBefore) {
    const theirs = groupsAfter.get(key) ?? [];
    const shared = Math.min(mine.length, theirs.length);
    for (let i = 0; i < shared; i++) {
      const from = mine[i]!.id;
      const to = theirs[i]!.id;
      relabel.set(from, to);
      if (from !== to) relabelled.push({ source, from, to, target: key });
    }
    for (let i = shared; i < mine.length; i++) {
      differences.push({
        kind: 'relationship',
        part: source,
        detail:
          'the relationship ' +
          mine[i]!.id +
          ' (' +
          key +
          ') is in the original and not in what was written',
      });
    }
  }

  for (const [key, theirs] of groupsAfter) {
    const mine = groupsBefore.get(key) ?? [];
    for (let i = mine.length; i < theirs.length; i++) {
      differences.push({
        kind: 'relationship',
        part: source,
        detail:
          'the relationship ' +
          theirs[i]!.id +
          ' (' +
          key +
          ') was written and is not in the original',
      });
    }
  }

  return { relabel, relabelled, differences, signature, writtenSignature };
}

/**
 * Every part that could own relationships, in either package, plus the root.
 *
 * Every part, and not only the ones with a `.rels` in `partNames` - which is
 * the shorter, faster and wrong version of this function, and the same trap
 * `reachability.ts` documents. A relationship collection created in this
 * session exists as a parsed object until `PartStore.write` materialises it, so
 * keying on existing relationship parts makes exactly the edges an editing
 * session added invisible. `PartStore.relationships` returns an empty
 * collection for a part that has none, so asking every part costs a lookup and
 * removes the whole class of miss.
 */
function relationshipSources(original: PartStore, written: PartStore): string[] {
  const sources = new Set<string>(['/']);
  for (const store of [original, written]) {
    for (const partName of store.partNames) {
      if (!isRelationshipPartName(partName)) {
        sources.add(partName);
        continue;
      }
      // And the source a `.rels` names, even when that part is not in the
      // package. An orphan relationship part is a defect, but it is one this
      // comparison should be able to see rather than one it drops.
      try {
        sources.add(sourcePartNameForRels(partName));
      } catch {
        // A name that does not decode to a source is not a relationship part in
        // any useful sense; it is compared as bytes with everything else.
      }
    }
  }
  return [...sources].sort();
}

function readRelationships(store: PartStore, source: string): Relationships | Error {
  try {
    return store.relationships(source);
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
}

/** The content types are the same type if they differ only in case or spacing. */
function sameContentType(a: string | undefined, b: string | undefined): boolean {
  return (a ?? '').trim().toLowerCase() === (b ?? '').trim().toLowerCase();
}

/**
 * Compare two packages.
 *
 * Never throws. Every failure - a part that will not parse, a `.rels` that will
 * not read - is a difference with a reason, because the caller of a comparison
 * wants the whole list and not the first item on it.
 */
export function comparePackages(
  original: PartStore,
  written: PartStore,
  options: CompareOptions = {},
): RoundTripReport {
  const differences: Difference[] = [];
  const parts: PartComparison[] = [];
  const relabelled: Relabelling[] = [];

  // The relationship pass runs first, because the mapping it produces is what
  // makes the XML pass able to see past a renumbering.
  const relabelBySource = new Map<string, ReadonlyMap<string, string>>();
  const graphs = new Map<string, GraphMatch>();
  for (const source of relationshipSources(original, written)) {
    const before = readRelationships(original, source);
    const after = readRelationships(written, source);
    if (before instanceof Error || after instanceof Error) {
      const which = before instanceof Error ? 'the original' : 'what was written';
      const error = before instanceof Error ? before : (after as Error);
      differences.push({
        kind: 'unreadable',
        part: source,
        detail: 'the relationships of ' + which + ' would not parse: ' + error.message,
      });
      continue;
    }
    const match = matchGraphs(source, before, after);
    graphs.set(normalizePartName(source), match);
    relabelBySource.set(normalizePartName(source), match.relabel);
    differences.push(...match.differences);
    relabelled.push(...match.relabelled);
  }

  const names = new Map<string, string>();
  for (const store of [original, written]) {
    for (const partName of store.partNames) names.set(normalizePartName(partName), partName);
  }

  let xmlCount = 0;
  let binaryCount = 0;
  let relsCount = 0;

  for (const key of [...names.keys()].sort()) {
    const part = names.get(key)!;
    const inOriginal = original.has(part);
    const inWritten = written.has(part);

    if (!inOriginal || !inWritten) {
      differences.push({
        kind: inOriginal ? 'part-removed' : 'part-added',
        part,
        detail: inOriginal
          ? 'the original has this part and what was written does not'
          : 'this part was written and the original does not have it',
      });
      continue;
    }

    const typeBefore = original.contentTypeOf(part);
    const typeAfter = written.contentTypeOf(part);
    if (!sameContentType(typeBefore, typeAfter)) {
      differences.push({
        kind: 'content-type',
        part,
        detail:
          'the content type was ' +
          (typeBefore ?? '(none)') +
          ' and is now ' +
          (typeAfter ?? '(none)'),
      });
    }

    if (isRelationshipPartName(part)) {
      relsCount++;
      parts.push(relationshipComparison(part, graphs));
      continue;
    }

    if (isXmlContentType(typeBefore) && isXmlContentType(typeAfter)) {
      xmlCount++;
      parts.push(
        compareXmlPart(part, original, written, relabelBySource, differences, options.limits),
      );
      continue;
    }

    binaryCount++;
    parts.push(compareBinaryPart(part, original, written, differences));
  }

  const same = parts.filter((entry) => entry.same).length;
  return {
    ok: differences.length === 0,
    parts,
    differences,
    relabelled,
    counts: { xml: xmlCount, binary: binaryCount, relationships: relsCount, same },
  };
}

/**
 * A relationship part's verdict, taken from the graph pass rather than from its
 * bytes.
 *
 * The digest is over the graph with the ids omitted, so it is the same on both
 * sides of a renumbering. That is the whole claim of this comparison stated as
 * a single value: two `.rels` with one digest describe one package.
 */
function relationshipComparison(part: string, graphs: Map<string, GraphMatch>): PartComparison {
  let source: string;
  try {
    source = sourcePartNameForRels(part);
  } catch {
    return { part, how: 'relationships', same: true, digest: '', writtenDigest: '' };
  }
  const match = graphs.get(normalizePartName(source));
  if (match === undefined) {
    // Only reachable when the graph pass reported the pair unreadable, which is
    // already a difference; saying nothing further here avoids reporting the
    // same failure twice under two kinds.
    return { part, how: 'relationships', same: false, digest: '', writtenDigest: '' };
  }
  const digest = sha256HexOfText(match.signature);
  const writtenDigest = sha256HexOfText(match.writtenSignature);
  return { part, how: 'relationships', same: digest === writtenDigest, digest, writtenDigest };
}

function compareXmlPart(
  part: string,
  original: PartStore,
  written: PartStore,
  relabelBySource: Map<string, ReadonlyMap<string, string>>,
  differences: Difference[],
  limits: XmlParseLimits | undefined,
): PartComparison {
  // The mapping that applies to *this* part's markup is the one derived from
  // this part's own `.rels`. An rId is scoped to one relationship collection -
  // treating them as global is the bug `PartStore` documents in its allocator -
  // so a slide's rId3 and a chart's rId3 are unrelated names.
  const relabel = relabelBySource.get(normalizePartName(part));
  const rewriteAttributeValue =
    relabel === undefined || relabel.size === 0
      ? undefined
      : (value: string, namespaceUri: string): string =>
          namespaceUri === R ? (relabel.get(value) ?? value) : value;

  // Applied to the original only, never to both. The mapping takes the
  // original's ids *into* the written package's naming, so running it on the
  // written side as well would map them a second time and the two would agree
  // again - including in the case that matters, where two ids swapped what they
  // point at and both documents still say `rId1`.
  const canonical = (
    store: PartStore,
    which: string,
    relabelThis: boolean,
  ): string | Difference => {
    try {
      const document = parseXml(store.read(part), limits);
      return canonicalXml(
        document,
        relabelThis && rewriteAttributeValue !== undefined ? { rewriteAttributeValue } : {},
      );
    } catch (error) {
      return {
        kind: 'unreadable',
        part,
        detail:
          which +
          ' would not parse as XML: ' +
          (isXmlError(error) ? error.code + ' - ' + error.message : String(error)),
      };
    }
  };

  const before = canonical(original, 'the original', true);
  const after = canonical(written, 'what was written', false);
  if (typeof before !== 'string' || typeof after !== 'string') {
    if (typeof before !== 'string') differences.push(before);
    if (typeof after !== 'string') differences.push(after);
    return { part, how: 'xml', same: false, digest: '', writtenDigest: '' };
  }

  const digest = sha256HexOfText(before);
  const writtenDigest = sha256HexOfText(after);
  if (digest !== writtenDigest) {
    const at = firstDifference(before, after);
    differences.push({
      kind: 'xml',
      part,
      detail:
        'the canonical XML differs at character ' +
        String(at) +
        '\n    original: ' +
        excerpt(before, at) +
        '\n    written:  ' +
        excerpt(after, at),
    });
  }
  return { part, how: 'xml', same: digest === writtenDigest, digest, writtenDigest };
}

function compareBinaryPart(
  part: string,
  original: PartStore,
  written: PartStore,
  differences: Difference[],
): PartComparison {
  let digest: string;
  let writtenDigest: string;
  try {
    digest = sha256Hex(original.read(part));
    writtenDigest = sha256Hex(written.read(part));
  } catch (error) {
    differences.push({
      kind: 'unreadable',
      part,
      detail: 'the part could not be read: ' + (error instanceof Error ? error.message : ''),
    });
    return { part, how: 'binary', same: false, digest: '', writtenDigest: '' };
  }
  if (digest !== writtenDigest) {
    differences.push({
      kind: 'binary',
      part,
      detail:
        'the bytes differ: sha256 was ' +
        digest.slice(0, 12) +
        ', is now ' +
        writtenDigest.slice(0, 12) +
        ' (' +
        String(original.read(part).length) +
        ' bytes then, ' +
        String(written.read(part).length) +
        ' now)',
    });
  }
  return { part, how: 'binary', same: digest === writtenDigest, digest, writtenDigest };
}

/** A one-line summary, for a CLI and for a test failure message. */
export function summarizeRoundTrip(report: RoundTripReport): string {
  const { xml, binary, relationships, same } = report.counts;
  const total = xml + binary + relationships;
  return (
    String(same) +
    '/' +
    String(total) +
    ' parts identical (' +
    String(xml) +
    ' xml, ' +
    String(relationships) +
    ' rels, ' +
    String(binary) +
    ' binary)' +
    (report.relabelled.length === 0
      ? ''
      : ', ' + String(report.relabelled.length) + ' relationship id(s) renamed') +
    (report.ok ? '' : ', ' + String(report.differences.length) + ' difference(s)')
  );
}

/**
 * Read a package, write it back, and compare the two.
 *
 * The whole of sub-phase 1.4 in one call, and the shape `cli roundtrip` and the
 * corpus gate both want. Every option `exportPackage` takes is accepted, so a
 * caller can round-trip through a different deflate level or a normalised entry
 * order - which is the point of a comparison that is not byte equality, and
 * worth having a test do deliberately.
 *
 * Throws whatever `exportPackage` throws. A package the firewall refuses is not
 * a package with a round-trip difference; it is one that never got written, and
 * flattening the two into one verdict would lose the distinction exactly where
 * it matters.
 */
export function roundTripPackage(
  bytes: Uint8Array,
  options: RoundTripOptions = {},
): RoundTripResult {
  const { limits, ...exportOptions } = options;
  const opened = openPackage(bytes, options.zip ?? {});
  const exported = exportPackage({ ...exportOptions, ...opened });
  const written = PartStore.open(exported.bytes, options.zip ?? {});
  const comparison = comparePackages(
    opened.baseline,
    written,
    limits === undefined ? {} : { limits },
  );
  return { exported, written, comparison, ok: comparison.ok };
}

export interface RoundTripOptions
  extends Omit<ExportOptions, 'store' | 'baseline' | 'baselineBytes'>, CompareOptions {}

export interface RoundTripResult {
  readonly exported: ExportResult;
  /** The package as re-opened from the bytes that were written. */
  readonly written: PartStore;
  readonly comparison: RoundTripReport;
  readonly ok: boolean;
}

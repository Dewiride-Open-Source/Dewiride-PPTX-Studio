import {
  CONTENT_TYPES_PART,
  isContentTypesStreamName,
  isRelationshipPartName,
  normalizePartName,
  type PartStore,
  type PartName,
  type ZipArchive,
} from '@pptx-studio/opc';
import { parseXml, type XDocument } from '@pptx-studio/xml';
import type { Location } from './report/location.js';
import type { Finding, ReadProblem } from './report/report.js';
import { ruleById, type RuleId } from './rules/rules.js';

/**
 * The state a rule reads, and the one method it writes.
 *
 * Every rule is `(ctx: Context) => void`. That is a deliberately small
 * interface, and it is the same shape `checkLayering({manifests, catalog})` and
 * `checkCorpus({manifests, files, ...})` already have in this repository, for
 * the same reason: a rule that takes a context and returns nothing is a rule
 * that can be run against a package assembled in a test, with no temporary
 * directory, no archive on disk and no other twenty-eight rules running beside
 * it obscuring which one fired.
 *
 * ## Everything is lazy, and everything is cached
 *
 * Twelve of the rules want the parsed tree of every XML part. Parsing each part
 * once per rule would be twelve passes over a deck that can be two hundred
 * megabytes. Parsing all of them up front would be one pass too many for the
 * package-level rules, which need no trees at all.
 *
 * So `document()` parses on first ask and remembers. A part that is not XML
 * returns `null` and is not asked again. A part that will not *parse* also
 * returns `null` - and records a `ReadProblem`, because a rule that silently
 * saw nothing is indistinguishable from a rule that found nothing, and those
 * are opposite answers.
 */
export interface Context {
  /** The package about to be handed over, or the one being inspected. */
  readonly store: PartStore;
  /** The archive bytes, when the caller had them. `null` when only a store was given. */
  readonly bytes: Uint8Array | null;
  /**
   * The archive those bytes describe, central directory only - nothing is
   * inflated to build it.
   *
   * `PartStore` deliberately does not expose this. It drops directory entries
   * on the way in, because a directory entry is not a part and carrying a thing
   * with no part name through the whole system would mean every consumer has to
   * remember it is there. That is right for the store and it leaves three rules
   * with nothing to look at: whether the archive has directory entries, whether
   * it carries a ZIP64 record, and what `[Content_Types].xml` actually *says* as
   * opposed to what parsing it produced.
   *
   * `null` when the caller passed a store and no bytes; the rules that need it
   * are then skipped with a reason rather than passing vacuously.
   */
  readonly archive: ZipArchive | null;
  /** The package as it was opened. `null` for an inspection with no history. */
  readonly baseline: PartStore | null;

  /** Part names in package order, `[Content_Types].xml` excluded - it is not a part. */
  parts(): readonly PartName[];
  /** A part's bytes, or `null` if it is not there. Never throws. */
  read(part: string): Uint8Array | null;
  /** Content type, or `undefined` when nothing in the map covers the part. */
  contentType(part: string): string | undefined;
  /** The parsed tree, cached. `null` for binary parts and for parts that will not parse. */
  document(part: string): XDocument | null;
  /**
   * `[Content_Types].xml` as the archive holds it, parsed.
   *
   * Not reachable through `parts()`: the content-type stream is not a part. It
   * has no content type of its own and nothing ever relates to it, which is why
   * `PartStore` keeps it outside the part map entirely.
   *
   * Read from the markup rather than from `store.contentTypes` for the reason
   * `rels.ts` gives at length: `ContentTypes.parse` collapses a duplicate
   * `Default` that agrees with itself and refuses one that does not, so a rule
   * that asks the parsed map whether there are duplicates is asking the one
   * object in the system that cannot answer.
   */
  contentTypesDocument(): XDocument | null;

  /** The same three, against the baseline. `null` throughout when there is none. */
  baselineRead(part: string): Uint8Array | null;
  baselineDocument(part: string): XDocument | null;

  /**
   * True when this session replaced or added the part.
   *
   * Derived from `PartInfo.fromArchive`, which is the store's own record of
   * whether a part's bytes still come from the archive it was opened from. It
   * is the only honest source for "did we touch this" - a byte comparison
   * cannot tell an edit that happened to produce identical bytes from no edit
   * at all, and the preservation rules care about the difference.
   */
  edited(part: string): boolean;

  add(rule: RuleId, where: Location, message: string): void;
  problem(part: string, message: string): void;
}

export interface ContextInput {
  readonly store: PartStore;
  readonly bytes?: Uint8Array | null;
  readonly archive?: ZipArchive | null;
  readonly baseline?: PartStore | null;
}

/** Content types whose parts are XML, beyond everything ending in `+xml`. */
const XML_CONTENT_TYPES = new Set(['application/xml', 'text/xml']);

function isXmlContentType(contentType: string | undefined): boolean {
  if (contentType === undefined) return false;
  const base = contentType.split(';')[0]!.trim().toLowerCase();
  return base.endsWith('+xml') || XML_CONTENT_TYPES.has(base);
}

class RuntimeContext implements Context {
  readonly store: PartStore;
  readonly bytes: Uint8Array | null;
  readonly archive: ZipArchive | null;
  readonly baseline: PartStore | null;

  readonly findings: Finding[] = [];
  readonly problems: ReadProblem[] = [];

  #parts: PartName[] | null = null;
  readonly #documents = new Map<string, XDocument | null>();
  readonly #baselineDocuments = new Map<string, XDocument | null>();

  constructor(input: ContextInput) {
    this.store = input.store;
    this.bytes = input.bytes ?? null;
    this.archive = input.archive ?? null;
    this.baseline = input.baseline ?? null;
  }

  parts(): readonly PartName[] {
    this.#parts ??= this.store.partNames;
    return this.#parts;
  }

  read(part: string): Uint8Array | null {
    if (!this.store.has(part)) return null;
    try {
      return this.store.read(part);
    } catch (error) {
      this.problem(part, describe(error));
      return null;
    }
  }

  contentType(part: string): string | undefined {
    return this.store.contentTypeOf(part);
  }

  document(part: string): XDocument | null {
    const key = normalizePartName(part);
    const cached = this.#documents.get(key);
    if (cached !== undefined) return cached;
    const parsed = this.#parse(part, this.read(part));
    this.#documents.set(key, parsed);
    return parsed;
  }

  baselineDocument(part: string): XDocument | null {
    if (this.baseline === null) return null;
    const key = normalizePartName(part);
    const cached = this.#baselineDocuments.get(key);
    if (cached !== undefined) return cached;
    const parsed = this.#parse(part, this.baselineRead(part), true);
    this.#baselineDocuments.set(key, parsed);
    return parsed;
  }

  #parse(part: string, bytes: Uint8Array | null, quiet = false): XDocument | null {
    if (bytes === null) return null;
    // The content type decides, not the extension. A `.bin` OLE object and a
    // `.png` are both binary and neither would tokenize; more to the point, an
    // `.xml` extension covers a `Default` that some producer typed
    // `application/octet-stream`, and parsing that would be us disagreeing with
    // the package about what it holds.
    if (!isRelationshipPartName(part) && !isXmlContentType(this.contentType(part))) return null;
    try {
      return parseXml(bytes);
    } catch (error) {
      // Not a rule violation. `@pptx-studio/xml` refuses a DOCTYPE, a bad
      // encoding and markup that is not well formed, and each of those is a
      // fact about the file rather than about one of the twenty-nine rules.
      // Recording it as a problem keeps the report honest: the rules that
      // wanted this part did not run on it.
      if (!quiet) this.problem(part, describe(error));
      return null;
    }
  }

  #contentTypesDocument: XDocument | null | undefined;

  contentTypesDocument(): XDocument | null {
    if (this.#contentTypesDocument !== undefined) return this.#contentTypesDocument;
    this.#contentTypesDocument = null;
    const archive = this.archive;
    if (archive !== null) {
      const entry = archive.entries.find((e) => isContentTypesStreamName(e.name));
      if (entry !== undefined) {
        try {
          this.#contentTypesDocument = parseXml(archive.read(entry));
        } catch (error) {
          this.problem(CONTENT_TYPES_PART, describe(error));
        }
      }
    }
    return this.#contentTypesDocument;
  }

  baselineRead(part: string): Uint8Array | null {
    if (this.baseline === null || !this.baseline.has(part)) return null;
    try {
      return this.baseline.read(part);
    } catch {
      return null;
    }
  }

  edited(part: string): boolean {
    return this.store.info(part)?.fromArchive === false;
  }

  add(rule: RuleId, where: Location, message: string): void {
    const definition = ruleById(rule);
    if (definition === undefined) return;
    this.findings.push({
      rule,
      severity: definition.severity,
      category: definition.category,
      where,
      message,
      // Filled in by the driver, which is the only thing that knows whether a
      // baseline pass ran. A rule never decides its own finding's history.
      origin: 'unknown',
    });
  }

  problem(part: string, message: string): void {
    const already = this.problems.some((p) => p.part === part && p.message === message);
    if (!already) this.problems.push({ part, message });
  }
}

function describe(error: unknown): string {
  if (error instanceof Error) {
    const code = (error as { code?: unknown }).code;
    return typeof code === 'string' ? code + ': ' + error.message : error.message;
  }
  return String(error);
}

export function createContext(input: ContextInput): RuntimeContext {
  return new RuntimeContext(input);
}

export type { RuntimeContext };

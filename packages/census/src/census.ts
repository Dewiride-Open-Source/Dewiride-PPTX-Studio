import {
  CONTENT_TYPE,
  FONT_DATA_CONTENT_TYPE,
  FONT_DATA_EXTENSION,
  isRelationshipPartName,
  isXmlContentType,
  normalizePartName,
  PartStore,
  readZip,
  REL_TYPE,
  ROOT_RELS_PART,
  sourcePartNameForRels,
  type Relationship,
  type Relationships,
  type ZipLimits,
} from '@pptx-studio/opc';
import {
  attributeValue,
  childElements,
  descendantElements,
  isXmlError,
  namespaceOf,
  parseXml,
  type XElement,
  type XmlTokenizerLimits,
} from '@pptx-studio/xml';
import { PART_FEATURE_RULES, FEATURE_RULES } from './features.js';
import { EXTENSION_NS, isExtensionNamespace, isStandardNamespace, OOXML_NS } from './namespaces.js';
import { createTotals, scanXmlPart } from './scan.js';
import type {
  ArchiveCensus,
  CensusProblem,
  EmbeddedFontCensus,
  NameCount,
  PackageCensus,
  PartCensus,
  PresentationCensus,
  RelationshipCensus,
} from './types.js';

/**
 * What is inside this `.pptx`.
 *
 * Two rules govern everything below, and both come from where the answer has to
 * end up.
 *
 * **A census never returns a tree.** The result crosses a `postMessage`
 * boundary from a Web Worker to the tab, and an `XNode` tree carries a parent
 * pointer on every node plus the entire decoded source on the document. Posting
 * one is not a message, it is a second copy of the deck, structure-cloned. So
 * parts are scanned at token level and only counters come back. The single
 * exception is `ppt/presentation.xml`, which is parsed into a tree because it
 * is a few kilobytes and because the structure genuinely wanted is nested; the
 * tree dies inside this function.
 *
 * **A census never refuses.** It is the tool you reach for when a deck is
 * broken, so a part that will not tokenize becomes a problem entry and the scan
 * carries on. The one thing that ends it is an archive that will not open at
 * all, and that throws an `OpcError` with a code, from `opc`.
 *
 * This is not the validator. Sub-phase 1.2 owns the 29 must-not-break rules and
 * refuses to hand over bytes when one fires. The problems reported here are
 * observations about a file someone else wrote, and several of them are things
 * PowerPoint tolerates.
 */
export interface CensusOptions {
  /** Passed to the ZIP reader. Defaults hold a 200 MB / 1 GB envelope. */
  readonly limits?: Partial<ZipLimits> | undefined;
  readonly tokenizerLimits?: XmlTokenizerLimits | undefined;
  /**
   * Called before each part is scanned.
   *
   * A 200 MB deck takes long enough that a progress bar is the difference
   * between "working" and "hung", and the worker has nothing else to report.
   */
  readonly onProgress?: ((done: number, total: number, part: string) => void) | undefined;
  /** Injectable clock, so a test can assert on timings without a real one. */
  readonly now?: (() => number) | undefined;
}

const P = OOXML_NS.p;
const P14 = EXTENSION_NS.p14;

/** ST_OnOff: `1`/`0`, `true`/`false`, and the `on`/`off` spelling Office writes. */
function ooxmlBoolean(value: string | undefined, fallback = false): boolean {
  if (value === undefined) return fallback;
  return value === '1' || value === 'true' || value === 'on';
}

function integerOrNull(value: string | undefined): number | null {
  if (value === undefined) return null;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function childByName(parent: XElement, uri: string, local: string): XElement | undefined {
  for (const child of childElements(parent)) {
    if (child.local === local && namespaceOf(child) === uri) return child;
  }
  return undefined;
}

/** Histogram map to sorted array: count descending, then name, so output is stable. */
function toNameCounts(
  map: ReadonlyMap<string, number>,
  limit = Number.POSITIVE_INFINITY,
): NameCount[] {
  return [...map]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
    .slice(0, limit);
}

/** `.../officeDocument/2006/relationships/slideLayout` -> `slideLayout`. */
function shortRelationshipType(type: string): string {
  const cut = type.lastIndexOf('/');
  return cut < 0 ? type : type.slice(cut + 1);
}

export function censusPackage(bytes: Uint8Array, options: CensusOptions = {}): PackageCensus {
  const now = options.now ?? (() => performance.now());
  const started = now();
  const problems: CensusProblem[] = [];
  const note = (
    severity: CensusProblem['severity'],
    code: string,
    message: string,
    part: string | null = null,
  ): void => {
    problems.push({ severity, code, message, part });
  };

  const readOptions = options.limits === undefined ? {} : { limits: options.limits };

  // The central directory is read twice on purpose. `PartStore` owns the
  // archive it opened and does not hand it out, which is the right shape for a
  // store; a census wants per-entry compression facts the store has no reason
  // to carry. Re-reading the directory inflates nothing and costs microseconds
  // even at 200 MB, so the alternative - widening `opc`'s API to serve one
  // consumer - would be the more expensive of the two.
  const archive = readZip(bytes, readOptions);
  const store = PartStore.open(bytes, readOptions);
  const openMs = now() - started;

  const entryByName = new Map(archive.entries.map((entry) => [entry.name, entry]));

  // --- the archive ---------------------------------------------------------
  let directoryEntries = 0;
  let storedEntries = 0;
  let deflatedEntries = 0;
  let dataDescriptorEntries = 0;
  let declaredInflatedBytes = 0;
  let compressedBytes = 0;
  for (const entry of archive.entries) {
    if (entry.isDirectory) directoryEntries += 1;
    if (entry.method === 0) storedEntries += 1;
    else if (entry.method === 8) deflatedEntries += 1;
    if (entry.hasDataDescriptor) dataDescriptorEntries += 1;
    declaredInflatedBytes += entry.uncompressedSize;
    compressedBytes += entry.compressedSize;
  }

  // --- relationships -------------------------------------------------------
  const partNames = store.partNames;
  const relationshipCounts = new Map<string, number>();
  const byType = new Map<string, number>();
  const edges: { source: string; rel: Relationship }[] = [];
  let external = 0;

  for (const name of partNames) {
    if (!isRelationshipPartName(name)) continue;
    const source = sourcePartNameForRels(name);
    let rels;
    try {
      rels = store.relationships(source);
    } catch (error) {
      note(
        'error',
        'RELS_UNREADABLE',
        'relationship part could not be read: ' + describe(error),
        name,
      );
      continue;
    }
    relationshipCounts.set(normalizePartName(source), rels.size);
    for (const rel of rels.all) {
      edges.push({ source, rel });
      bumpMap(byType, shortRelationshipType(rel.type));
      if (rel.targetMode === 'External') external += 1;
    }
  }

  const dangling = store.danglingRelationships().map((entry) => ({
    source: String(entry.source),
    id: entry.relationship.id,
    target: String(entry.target),
  }));
  for (const entry of dangling) {
    note(
      'warning',
      'DANGLING_RELATIONSHIP',
      entry.id + ' points at ' + entry.target + ', which is not in the package',
      entry.source,
    );
  }

  const reachable = walkReachable(store, partNames);
  const unreachableParts = partNames
    .filter((name) => !isRelationshipPartName(name) && !reachable.has(normalizePartName(name)))
    .map(String);

  const relationships: RelationshipCensus = {
    total: edges.length,
    internal: edges.length - external,
    external,
    byType: toNameCounts(byType),
    dangling,
    unreachableParts,
  };

  // --- parts ---------------------------------------------------------------
  const totals = createTotals();
  const parts: PartCensus[] = [];
  const contentTypeCounts = new Map<string, number>();
  const featureCounts = totals.features;
  let inflateMs = 0;
  let scanMs = 0;
  let xmlBytesScanned = 0;
  let partsScanned = 0;
  let largestPart: NameCount | null = null;

  const scannable = partNames.filter((name) => !isRelationshipPartName(name));

  for (const [index, name] of scannable.entries()) {
    options.onProgress?.(index, scannable.length, String(name));

    const resolved = store.contentTypes.resolve(name);
    const contentType = resolved?.contentType ?? '';
    if (resolved === undefined) {
      note(
        'error',
        'NO_CONTENT_TYPE',
        'no Default or Override gives this part a content type',
        name,
      );
    }
    bumpMap(contentTypeCounts, contentType === '' ? '(none)' : contentType);

    for (const rule of PART_FEATURE_RULES) {
      if (rule.match(String(name).toLowerCase(), contentType)) bumpMap(featureCounts, rule.key);
    }

    const entry = entryByName.get(String(name).slice(1));
    const declaredSize = entry?.uncompressedSize ?? 0;
    if (largestPart === null || declaredSize > largestPart.count) {
      largestPart = { name: String(name), count: declaredSize };
    }

    const xml = isXmlContentType(contentType);
    let elements: number | null = null;
    let maxDepth: number | null = null;
    let size = declaredSize;

    if (xml) {
      // `.rels` parts are excluded above rather than here: in production they
      // are read by the OPC layer's own flat-XML reader and never reach the
      // tokenizer, so counting their elements would describe a code path that
      // does not exist.
      try {
        const beforeInflate = now();
        const partBytes = store.read(name);
        inflateMs += now() - beforeInflate;
        size = partBytes.byteLength;

        const beforeScan = now();
        const scan = scanXmlPart(decodeUtf8(partBytes), totals, options.tokenizerLimits);
        scanMs += now() - beforeScan;

        elements = scan.elements;
        maxDepth = scan.maxDepth;
        if (scan.imbalance !== null) {
          note('error', 'PART_NOT_WELL_FORMED', scan.imbalance, name);
        }
        xmlBytesScanned += partBytes.byteLength;
        partsScanned += 1;
      } catch (error) {
        note('error', isXmlError(error) ? error.code : 'PART_UNREADABLE', describe(error), name);
      }
    }

    parts.push({
      name: String(name),
      contentType,
      contentTypeOrigin: resolved?.origin ?? 'unknown',
      kind: xml ? 'xml' : 'binary',
      bytes: size,
      compressedBytes: entry?.compressedSize ?? 0,
      elements,
      maxDepth,
      relationships: relationshipCounts.get(normalizePartName(name)) ?? 0,
    });
  }
  options.onProgress?.(scannable.length, scannable.length, '');

  // --- the presentation ----------------------------------------------------
  const presentation = readPresentation(store, parts, note);

  // --- package-level observations -----------------------------------------
  const hasFontData = parts.some((part) =>
    part.name.toLowerCase().endsWith('.' + FONT_DATA_EXTENSION),
  );
  const declaresFontData = store.contentTypes.defaults.some(
    (entry) => entry.extension.toLowerCase() === FONT_DATA_EXTENSION,
  );
  if (hasFontData && !declaresFontData) {
    note(
      'error',
      'FNTDATA_DEFAULT_MISSING',
      'the package ships .' +
        FONT_DATA_EXTENSION +
        ' parts with no <Default Extension="' +
        FONT_DATA_EXTENSION +
        '" ContentType="' +
        FONT_DATA_CONTENT_TYPE +
        '"/>. This is the canonical "PowerPoint found a problem with some content" bug.',
    );
  }
  if (
    presentation !== null &&
    presentation.embeddedFonts.length > 0 &&
    !presentation.embedTrueTypeFonts
  ) {
    note(
      'warning',
      'EMBED_FLAG_MISSING',
      'p:embeddedFontLst has ' +
        String(presentation.embeddedFonts.length) +
        ' entries but @embedTrueTypeFonts is not set, so PowerPoint ignores all of them',
    );
  }
  if (directoryEntries > 0) {
    note(
      'note',
      'DIRECTORY_ENTRIES',
      String(directoryEntries) +
        ' directory entries. PowerPoint writes none; this package was rebuilt by another tool.',
    );
  }
  if (unreachableParts.length > 0) {
    note(
      'note',
      'UNREACHABLE_PARTS',
      String(unreachableParts.length) + ' parts are not reachable from the package relationships',
    );
  }

  const archiveCensus: ArchiveCensus = {
    bytes: bytes.byteLength,
    entries: archive.entries.length,
    parts: scannable.length,
    directoryEntries,
    storedEntries,
    deflatedEntries,
    zip64: hasZip64Locator(bytes),
    dataDescriptorEntries,
    declaredInflatedBytes,
    compressedBytes,
    largestPart,
  };

  const featureLabels = new Map(
    [...FEATURE_RULES, ...PART_FEATURE_RULES].map((rule) => [rule.key, rule]),
  );

  return {
    censusVersion: 1,
    archive: archiveCensus,
    parts,
    contentTypes: toNameCounts(contentTypeCounts),
    relationships,
    presentation,
    features: [...featureCounts]
      .map(([key, count]) => {
        const rule = featureLabels.get(key);
        return { key, label: rule?.label ?? key, count, phase: rule?.phase ?? '?' };
      })
      .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key)),
    namespaces: [...totals.namespaces]
      .map(([uri, use]) => ({
        uri,
        count: use.count,
        prefixes: [...use.prefixes].sort(),
        standard: isStandardNamespace(uri),
        extension: isExtensionNamespace(uri),
        required: use.required,
        ignorable: use.ignorable,
      }))
      .sort((a, b) => b.count - a.count || a.uri.localeCompare(b.uri)),
    elements: toNameCounts(totals.elements, 50),
    presetGeometry: toNameCounts(totals.presetGeometry),
    typefaces: toNameCounts(totals.typefaces),
    languages: toNameCounts(totals.languages),
    text: {
      paragraphs: totals.paragraphs,
      runs: totals.runs,
      characters: totals.characters,
      fields: totals.fields,
      hardBreaks: totals.hardBreaks,
      bulletChar: totals.bulletChar,
      bulletAutoNum: totals.bulletAutoNum,
      bulletBlip: totals.bulletBlip,
      normAutofit: totals.normAutofit,
      shapeAutofit: totals.shapeAutofit,
    },
    problems,
    timings: {
      openMs,
      inflateMs,
      scanMs,
      totalMs: now() - started,
      xmlBytesScanned,
      partsScanned,
    },
  };
}

function bumpMap(map: Map<string, number>, key: string): void {
  map.set(key, (map.get(key) ?? 0) + 1);
}

function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

const UTF8 = new TextDecoder('utf-8');

/**
 * Bytes to a string.
 *
 * `@pptx-studio/xml` has `decodeXmlSource`, which honours the encoding
 * declaration and a byte-order mark. It is not used here because a census must
 * not fail on a part whose declared encoding we do not implement - it should
 * count what it can and report what it could not.
 */
function decodeUtf8(bytes: Uint8Array): string {
  return UTF8.decode(bytes);
}

/** True when the archive carries a ZIP64 end-of-central-directory locator. */
function hasZip64Locator(bytes: Uint8Array): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The locator sits immediately before the end-of-central-directory record,
  // which is within 64 KB + 22 of the end once a comment is allowed for.
  const from = Math.max(0, bytes.byteLength - 0x10000 - 40);
  for (let at = bytes.byteLength - 20; at >= from; at--) {
    if (view.getUint32(at, true) === 0x07064b50) return true;
  }
  return false;
}

/**
 * Every part reachable by following relationships from the package root.
 *
 * Unreachable is not the same as unused and certainly not the same as
 * deletable - PowerPoint keeps layouts nothing points at, and 1.3's media GC
 * has to know that. This walk exists so a census can *say* what is orphaned,
 * not so anything can act on it.
 */
function walkReachable(store: PartStore, partNames: readonly string[]): Set<string> {
  const reachable = new Set<string>();
  const queue: string[] = [];

  const follow = (rels: Relationships): void => {
    for (const rel of rels.all) {
      if (rel.targetMode === 'External') continue;
      let target: string;
      try {
        target = rels.resolve(rel);
      } catch {
        continue; // reported as a dangling edge elsewhere
      }
      const key = normalizePartName(target);
      if (reachable.has(key) || !store.has(target)) continue;
      reachable.add(key);
      queue.push(target);
    }
  };

  try {
    follow(store.rootRelationships());
  } catch {
    return reachable;
  }

  // Only parts that own a relationship part can extend the walk. Asking the
  // store for the relationships of every part would parse nothing but would
  // still walk the whole name list once per hop.
  const owners = new Set(
    partNames
      .filter(isRelationshipPartName)
      .map((name) => normalizePartName(sourcePartNameForRels(name))),
  );

  for (let part = queue.pop(); part !== undefined; part = queue.pop()) {
    if (!owners.has(normalizePartName(part))) continue;
    try {
      follow(store.relationships(part));
    } catch {
      continue; // reported when the relationship parts are counted
    }
  }
  return reachable;
}

/**
 * `ppt/presentation.xml`, read as a tree.
 *
 * The one place a census parses rather than scans. The part is a few kilobytes
 * and every fact wanted from it is positional - which child of which parent -
 * so a token scan would have to rebuild the nesting it deliberately does not
 * track. The tree does not leave this function.
 */
function readPresentation(
  store: PartStore,
  parts: readonly PartCensus[],
  note: (
    severity: CensusProblem['severity'],
    code: string,
    message: string,
    part?: string | null,
  ) => void,
): PresentationCensus | null {
  const countOfType = (contentType: string): number =>
    parts.filter((part) => part.contentType === contentType).length;

  let rootRels;
  try {
    rootRels = store.rootRelationships();
  } catch (error) {
    note('error', 'ROOT_RELS_UNREADABLE', describe(error), ROOT_RELS_PART);
    return null;
  }
  const officeDocument = rootRels.firstOfType(REL_TYPE.officeDocument);
  if (officeDocument === undefined) {
    note(
      'error',
      'NOT_A_PACKAGE',
      'the package relationships name no officeDocument, so this is not a presentation',
      ROOT_RELS_PART,
    );
    return null;
  }

  let root: XElement;
  try {
    root = parseXml(store.read(rootRels.resolve(officeDocument))).root;
  } catch (error) {
    note(
      'error',
      'PRESENTATION_UNREADABLE',
      describe(error),
      String(rootRels.resolve(officeDocument)),
    );
    return null;
  }

  const sldSz = childByName(root, P, 'sldSz');
  const notesSz = childByName(root, P, 'notesSz');
  const custShowLst = childByName(root, P, 'custShowLst');

  const embeddedFonts: EmbeddedFontCensus[] = [];
  const embeddedFontLst = childByName(root, P, 'embeddedFontLst');
  if (embeddedFontLst !== undefined) {
    for (const entry of childElements(embeddedFontLst)) {
      if (entry.local !== 'embeddedFont' || namespaceOf(entry) !== P) continue;
      const font = childByName(entry, P, 'font');
      const slots = childElements(entry)
        .filter((child) => child.local !== 'font' && namespaceOf(child) === P)
        .map((child) => child.local);
      embeddedFonts.push({
        typeface: font === undefined ? '' : (attributeValue(font, 'typeface') ?? ''),
        panose: font === undefined ? null : (attributeValue(font, 'panose') ?? null),
        charset: font === undefined ? null : integerOrNull(attributeValue(font, 'charset')),
        pitchFamily: font === undefined ? null : integerOrNull(attributeValue(font, 'pitchFamily')),
        slots,
      });
    }
  }

  let sections = 0;
  for (const element of descendantElements(root)) {
    if (element.local === 'section' && namespaceOf(element) === P14) sections += 1;
  }

  // Counted from the shape tree rather than from `p:sldIdLst`, because the two
  // can disagree and the list is what PowerPoint reads. Both are reported.
  const listedSlides = childByName(root, P, 'sldIdLst');
  const listed = listedSlides === undefined ? 0 : childElements(listedSlides).length;
  const slideParts = countOfType(CONTENT_TYPE.slide);
  if (listed !== slideParts) {
    note(
      'warning',
      'SLIDE_COUNT_MISMATCH',
      'p:sldIdLst names ' +
        String(listed) +
        ' slides but the package holds ' +
        String(slideParts) +
        ' slide parts',
    );
  }

  return {
    slideWidthEmu: sldSz === undefined ? null : integerOrNull(attributeValue(sldSz, 'cx')),
    slideHeightEmu: sldSz === undefined ? null : integerOrNull(attributeValue(sldSz, 'cy')),
    slideSizeType: sldSz === undefined ? null : (attributeValue(sldSz, 'type') ?? null),
    notesWidthEmu: notesSz === undefined ? null : integerOrNull(attributeValue(notesSz, 'cx')),
    notesHeightEmu: notesSz === undefined ? null : integerOrNull(attributeValue(notesSz, 'cy')),
    slides: slideParts,
    slideLayouts: countOfType(CONTENT_TYPE.slideLayout),
    slideMasters: countOfType(CONTENT_TYPE.slideMaster),
    notesSlides: countOfType(CONTENT_TYPE.notesSlide),
    notesMasters: countOfType(CONTENT_TYPE.notesMaster),
    handoutMasters: countOfType(CONTENT_TYPE.handoutMaster),
    embedTrueTypeFonts: ooxmlBoolean(attributeValue(root, 'embedTrueTypeFonts')),
    saveSubsetFonts: ooxmlBoolean(attributeValue(root, 'saveSubsetFonts')),
    embeddedFonts,
    sections,
    customShows: custShowLst === undefined ? 0 : childElements(custShowLst).length,
    firstSlideNum: integerOrNull(attributeValue(root, 'firstSlideNum')),
    rtl: ooxmlBoolean(attributeValue(root, 'rtl')),
  };
}

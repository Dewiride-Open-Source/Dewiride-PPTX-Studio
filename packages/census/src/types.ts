/**
 * The shape of a census.
 *
 * One constraint runs through every type here: **a census is plain JSON.** No
 * `Map`, no `Set`, no `undefined`, no class instances - histograms are arrays
 * of `{ name, count }` and absence is `null`.
 *
 * That is not aesthetic. A census is computed in a Web Worker and has to cross
 * the `postMessage` boundary to reach the tab that asked for it, and it is
 * written to disk by `pptx-studio inspect --json` so a later run can diff
 * against it. Structured clone would carry a `Map` across; `JSON.stringify`
 * would quietly drop it, and the two consumers would then disagree about what a
 * census is. Choosing the intersection of the two makes that impossible.
 *
 * The harder rule the boundary imposes is stated where it is enforced, in
 * `census.ts`: nothing here may hold an `XNode`. A parsed tree carries a parent
 * pointer on every node and the whole source string on the document, so posting
 * one is not a message, it is a second copy of the deck.
 */

/** A `{ name, count }` histogram entry. Sorted by count descending, then name. */
export interface NameCount {
  readonly name: string;
  readonly count: number;
}

/** How the archive itself is put together, before anything is inflated. */
export interface ArchiveCensus {
  /** The `.pptx` as read: compressed, headers and all. */
  readonly bytes: number;
  readonly entries: number;
  /** Entries that are OPC parts. The content-type stream is not one. */
  readonly parts: number;
  /**
   * Entries whose name ends in a slash.
   *
   * PowerPoint writes none and our writer never will. A package that has them
   * still opens, so this is a note rather than a problem - but it is the
   * clearest single signal that a deck was rebuilt by a third-party tool.
   */
  readonly directoryEntries: number;
  readonly storedEntries: number;
  readonly deflatedEntries: number;
  /**
   * True when the archive carries a ZIP64 end-of-central-directory record.
   *
   * PowerPoint never writes one and our writer never will, but other producers
   * emit one even for small archives - so this says which producer wrote the
   * file, not how big it is.
   */
  readonly zip64: boolean;
  /** Entries carrying a streaming data descriptor rather than sizes up front. */
  readonly dataDescriptorEntries: number;
  /** Sum of every entry's uncompressed size, as the central directory declares. */
  readonly declaredInflatedBytes: number;
  readonly compressedBytes: number;
  readonly largestPart: NameCount | null;
}

export type PartKind = 'xml' | 'binary';

/** One part, and what a scan of it found. Ordered as the archive stores them. */
export interface PartCensus {
  readonly name: string;
  readonly contentType: string;
  /** Where the content type came from: a `Default` extension or an `Override`. */
  readonly contentTypeOrigin: 'default' | 'override' | 'unknown';
  readonly kind: PartKind;
  readonly bytes: number;
  readonly compressedBytes: number;
  /** Elements in the part. `null` for a part that was not scanned. */
  readonly elements: number | null;
  /** Deepest element nesting. `null` for a part that was not scanned. */
  readonly maxDepth: number | null;
  /** Outbound relationships declared in this part's own relationship part. */
  readonly relationships: number;
}

export interface DanglingRelationship {
  readonly source: string;
  readonly id: string;
  readonly target: string;
}

export interface RelationshipCensus {
  readonly total: number;
  readonly internal: number;
  readonly external: number;
  /** Relationship type histogram, shortened to the last path segment. */
  readonly byType: readonly NameCount[];
  /** Edges whose target part is not in the package. PowerPoint tolerates these. */
  readonly dangling: readonly DanglingRelationship[];
  /** Parts nothing relates to. Harmless, but they are what a media GC would take. */
  readonly unreachableParts: readonly string[];
}

/** An embedded font, as `p:embeddedFontLst` declares it. */
export interface EmbeddedFontCensus {
  readonly typeface: string;
  readonly panose: string | null;
  readonly charset: number | null;
  readonly pitchFamily: number | null;
  /** Which of regular/bold/italic/boldItalic carry a relationship. */
  readonly slots: readonly string[];
}

export interface PresentationCensus {
  readonly slideWidthEmu: number | null;
  readonly slideHeightEmu: number | null;
  /** `p:sldSz/@type`, for instance `screen16x9`. */
  readonly slideSizeType: string | null;
  readonly notesWidthEmu: number | null;
  readonly notesHeightEmu: number | null;
  readonly slides: number;
  readonly slideLayouts: number;
  readonly slideMasters: number;
  readonly notesSlides: number;
  readonly notesMasters: number;
  readonly handoutMasters: number;
  /**
   * `p:presentation/@embedTrueTypeFonts`.
   *
   * Omit it and PowerPoint ignores a perfect `p:embeddedFontLst` entirely, so
   * the flag and the list are reported separately on purpose: a deck with fonts
   * and no flag is a bug someone shipped, not a deck without fonts.
   */
  readonly embedTrueTypeFonts: boolean;
  /** `p:presentation/@saveSubsetFonts`. Governs subsetting, never compression. */
  readonly saveSubsetFonts: boolean;
  readonly embeddedFonts: readonly EmbeddedFontCensus[];
  /** `p14:sectionLst` entries. */
  readonly sections: number;
  readonly customShows: number;
  readonly firstSlideNum: number | null;
  readonly rtl: boolean;
}

export interface TextCensus {
  readonly paragraphs: number;
  readonly runs: number;
  readonly characters: number;
  /** `a:fld` - dates, slide numbers, and the GUIDs that must not be regenerated. */
  readonly fields: number;
  readonly hardBreaks: number;
  readonly bulletChar: number;
  readonly bulletAutoNum: number;
  readonly bulletBlip: number;
  /** `a:normAutofit` - text shrunk to fit, at the scale PowerPoint chose. */
  readonly normAutofit: number;
  /** `a:spAutoFit` - the shape grows instead. */
  readonly shapeAutofit: number;
}

/** A namespace as it was actually used, with every prefix it was spelled with. */
export interface NamespaceCensus {
  readonly uri: string;
  /** Elements and attributes in this namespace, across every part. */
  readonly count: number;
  readonly prefixes: readonly string[];
  /** True when ECMA-376 itself defines the namespace. */
  readonly standard: boolean;
  /** True for a post-ECMA-376 extension namespace we recognise by name. */
  readonly extension: boolean;
  /** True when it appeared in an `mc:Choice` requirement. */
  readonly required: boolean;
  /** True when some `mc:Ignorable` named it. */
  readonly ignorable: boolean;
}

/**
 * A detected feature, and the phase of the plan that makes it render.
 *
 * `phase` is what turns a census into a schedule: dropping a deck on the page
 * and reading back "SmartArt 14, ChartEx 3, ink 0" says exactly which of this
 * project's unfinished phases stand between that file and a faithful render.
 */
export interface FeatureCensus {
  readonly key: string;
  readonly label: string;
  readonly count: number;
  /** The sub-phase or phase that implements it, for instance `4.6` or `10.2`. */
  readonly phase: string;
}

export type ProblemSeverity = 'error' | 'warning' | 'note';

export interface CensusProblem {
  readonly severity: ProblemSeverity;
  readonly code: string;
  readonly message: string;
  readonly part: string | null;
}

export interface CensusTimings {
  /** Opening the archive: end-of-central-directory, entry table, content types. */
  readonly openMs: number;
  /** Inflating every XML part. */
  readonly inflateMs: number;
  /** Tokenizing and counting. */
  readonly scanMs: number;
  readonly totalMs: number;
  /** XML bytes handed to the tokenizer, so throughput is derivable. */
  readonly xmlBytesScanned: number;
  readonly partsScanned: number;
}

export interface PackageCensus {
  /** Schema version of this document, so a stored census can be re-read later. */
  readonly censusVersion: 1;
  readonly archive: ArchiveCensus;
  readonly parts: readonly PartCensus[];
  readonly contentTypes: readonly NameCount[];
  readonly relationships: RelationshipCensus;
  readonly presentation: PresentationCensus | null;
  readonly features: readonly FeatureCensus[];
  readonly namespaces: readonly NamespaceCensus[];
  /** The most common qualified element names, spelled as they were written. */
  readonly elements: readonly NameCount[];
  /** `a:prstGeom/@prst` - which of the 187 presets this deck actually uses. */
  readonly presetGeometry: readonly NameCount[];
  /** Every typeface named by `a:latin`, `a:ea`, `a:cs` or `a:sym`. */
  readonly typefaces: readonly NameCount[];
  /** Every `@lang` seen on a run. */
  readonly languages: readonly NameCount[];
  readonly text: TextCensus;
  readonly problems: readonly CensusProblem[];
  readonly timings: CensusTimings;
}

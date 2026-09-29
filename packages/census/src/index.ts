/**
 * `@pptx-studio/census` - what is inside a `.pptx`, as plain JSON: parts, edges, a feature census
 * and namespaces, scanned at token level in a browser or a Worker. It never refuses: describing a
 * broken file is what it is for, and guarding a file about to be written is validate's. ADR 0008.
 */

export { censusPackage, type CensusOptions } from './census.js';

export { formatCensus, humanBytes, type FormatOptions } from './format.js';

export {
  OOXML_NS,
  EXTENSION_NS,
  CONVENTIONAL_PREFIX,
  isStandardNamespace,
  isExtensionNamespace,
} from './namespaces.js';

export { FEATURE_RULES, PART_FEATURE_RULES, type FeatureRule } from './features.js';

export { createTotals, scanXmlPart, type ScanTotals, type PartScan } from './scan.js';

export type {
  ArchiveCensus,
  CensusProblem,
  CensusTimings,
  DanglingRelationship,
  EmbeddedFontCensus,
  FeatureCensus,
  NameCount,
  NamespaceCensus,
  PackageCensus,
  PartCensus,
  PartKind,
  PresentationCensus,
  ProblemSeverity,
  RelationshipCensus,
  TextCensus,
} from './types.js';

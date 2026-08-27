/**
 * `@pptx-studio/census` - what is inside a `.pptx`.
 *
 * Open a package, walk its parts and its relationship graph, scan every XML
 * part at token level, and return a plain-JSON description: the archive, the
 * parts, the edges, the presentation's own structure, a feature census keyed to
 * the phase of the plan that renders each feature, every namespace with the
 * prefixes it was spelled with, and the timings.
 *
 * It exists because two consumers need the same answer and neither can be the
 * home of it. `pptx-studio inspect` is Node and could never run in a tab; the
 * browser explorer runs in a Web Worker and could never import Node. Layer 1,
 * browser runtime, no Node anywhere - the CLI is the one that adapts.
 *
 * Three properties are load-bearing rather than incidental:
 *
 * - **The result is plain JSON.** No `Map`, no `Set`, no class instance and no
 *   `undefined`, because it has to survive both `postMessage` and
 *   `JSON.stringify` and mean the same thing on the other side of each.
 * - **No tree ever leaves.** Scanning is done over tokens; the working set does
 *   not grow with the deck. `ppt/presentation.xml` is the single part parsed
 *   into a tree, and that tree does not escape the function that built it.
 * - **It does not refuse.** A part that will not tokenize becomes a problem
 *   entry and the scan continues, because a census is what you reach for when a
 *   file is already broken.
 *
 * This is not the validator. The 29 must-not-break rules and the refusal to
 * hand over bytes are sub-phase 1.2's, and they answer a different question:
 * this one describes a file somebody else wrote, that one guards a file we are
 * about to write.
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

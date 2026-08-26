/**
 * `@pptx-studio/opc` - the Open Packaging Conventions container.
 *
 * A `.pptx` is a ZIP archive of XML parts plus a graph of relationships between
 * them. This package owns the container: getting bytes out of an archive that
 * may be hostile, and naming the things inside it.
 *
 * As of sub-phase 0.3 that means the ZIP reader and writer, the decompression
 * budgets, the OPC part-name grammar, content-type resolution, relationships
 * and the part store. Part *content* is still opaque bytes: the XML layer is
 * 0.4.
 *
 * Everything here runs in a browser tab and in a Web Worker. There is no Node.
 */

export {
  CONTENT_TYPES_PART,
  ROOT_RELS_PART,
  OPC_NS,
  REL_TYPE,
  CONTENT_TYPE,
  FONT_DATA_EXTENSION,
  FONT_DATA_CONTENT_TYPE,
  type RelationshipType,
  type ContentType,
} from './constants.js';

export {
  OpcError,
  OPC_ERROR_CODES,
  isOpcError,
  guard,
  type OpcErrorCode,
  type OpcErrorDetail,
} from './errors.js';

export { crc32 } from './crc32.js';

export { DEFAULT_ZIP_LIMITS, InflationBudget, type ZipLimits } from './limits.js';

export {
  readZip,
  ZipArchive,
  COMPRESSION_STORE,
  COMPRESSION_DEFLATE,
  type ZipEntry,
  type ReadZipOptions,
} from './zip-reader.js';

export {
  MAX_PART_NAME_LENGTH,
  validatePartName,
  isValidPartName,
  toPartName,
  partNameFromZipEntry,
  zipEntryNameFor,
  normalizePartName,
  checkPartNameCollisions,
  partExtension,
  partDirectory,
  relsPartNameFor,
  isRelationshipPartName,
  isContentTypesStreamName,
  resolveRelativeTarget,
  type PartName,
  type PartNameRule,
  type PartNameSeverity,
  type PartNameViolation,
  type PartNameCollision,
} from './pack-uri.js';

export {
  readFlatXml,
  escapeAttribute,
  localName,
  XML_DECLARATION,
  DEFAULT_FLAT_XML_LIMITS,
  type FlatElement,
  type FlatXmlLimits,
} from './flat-xml.js';

export {
  ContentTypes,
  type ContentTypeOrigin,
  type DefaultEntry,
  type OverrideEntry,
  type ResolvedContentType,
} from './content-types.js';

export {
  Relationships,
  isValidRelationshipId,
  sourcePartNameForRels,
  relativeTargetFor,
  type Relationship,
  type RelationshipOrigin,
  type TargetMode,
} from './relationships.js';

export {
  writeZip,
  storedEntry,
  deflatedEntry,
  passthroughEntry,
  entryOverhead,
  EMPTY_ARCHIVE_SIZE,
  type ZipEntryInput,
} from './zip-writer.js';

export { PartStore, type PartInfo, type WritePackageOptions } from './part-store.js';

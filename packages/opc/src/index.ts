/**
 * `@pptx-studio/opc` - the Open Packaging Conventions container.
 *
 * A `.pptx` is a ZIP archive of XML parts plus a graph of relationships between
 * them. This package owns the container: getting bytes out of an archive that
 * may be hostile, and naming the things inside it.
 *
 * As of sub-phase 0.2 that means the ZIP reader, its decompression budgets, and
 * the OPC part-name grammar. `PartStore`, `[Content_Types].xml` resolution,
 * relationships and the writer land in 0.3.
 *
 * Everything here runs in a browser tab and in a Web Worker. There is no Node.
 */

export {
  CONTENT_TYPES_PART,
  ROOT_RELS_PART,
  OPC_NS,
  REL_TYPE,
  FONT_DATA_EXTENSION,
  FONT_DATA_CONTENT_TYPE,
  type RelationshipType,
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

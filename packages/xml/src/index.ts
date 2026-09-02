/**
 * `@pptx-studio/xml` - byte-preserving XML for OOXML.
 *
 * A `.pptx` part is XML we must be able to hand back exactly as we received it,
 * except for the one node someone edited. That is a harder guarantee than any
 * general-purpose XML library offers, and it is why this package exists:
 *
 *   - `DOMParser`/`XMLSerializer` may rewrite namespace prefixes, and
 *     `mc:Ignorable` and `mc:Choice/@Requires` hold **prefixes, not URIs**, so
 *     a rename silently turns ignorable extension markup into a hard error.
 *     `DOMParser` is also not reliably available in a Web Worker, which is
 *     where our parsing runs.
 *   - `@xmldom/xmldom` loses namespaces on `createElementNS` beneath a
 *     prefixed parent.
 *   - `fast-xml-parser` guarantees nothing about whitespace, self-closing form
 *     or quote style.
 *
 * As of sub-phase 0.6 this is the tokenizer, the node model, the serializer and
 * the round-trip gate that holds them together, plus schema-ordered insertion,
 * the Markup Compatibility walker, `extLst` as an opaque list, and edits that
 * each carry their exact inverse.
 *
 * Everything here runs in a browser tab and in a Web Worker. There is no Node.
 */

export { NS, XML_SPACE_PRESERVE, type NamespacePrefix, type NamespaceUri } from './namespaces.js';

export {
  XmlError,
  XML_ERROR_CODES,
  isXmlError,
  positionOf,
  type XmlErrorCode,
  type XmlErrorDetail,
} from './errors.js';

export {
  isXmlWhitespace,
  isNameStartChar,
  isNameChar,
  isXmlChar,
  isXmlCodePoint,
  isUnpairedSurrogate,
  normalizeLineEndings,
  isAllWhitespace,
} from './chars.js';

export { decodeXmlSource, encodeXmlSource, type XmlEncoding, type XmlSource } from './source.js';

export {
  decodeReference,
  decodeCharacterData,
  normalizeAttributeValue,
  isLiteralRun,
  escapeText,
  escapeAttributeValue,
} from './references.js';

export {
  XmlTokenizer,
  tokenize,
  checkSpanCoverage,
  splitQName,
  type XAttribute,
  DEFAULT_TOKENIZER_LIMITS,
  type XmlToken,
  type XmlTokenType,
  type XmlTokenizerLimits,
  type XmlDeclarationToken,
  type XmlProcessingInstructionToken,
  type XmlCommentToken,
  type XmlElementToken,
  type XmlEndTagToken,
  type XmlTextToken,
  type XmlCdataToken,
  type SpanGap,
} from './tokenizer.js';

export {
  parseXml,
  parseXmlString,
  sourceOf,
  startTagOf,
  childElements,
  firstChild,
  attribute,
  attributeValue,
  descendantElements,
  textContent,
  declaredNamespaces,
  resolvePrefix,
  namespaceOf,
  attributeNamespaceOf,
  namespaceScope,
  prefixFor,
  prefixMap,
  undeclaredPrefixes,
  xmlSpace,
  markDirty,
  markAttributeDirty,
  checkDirtyInvariant,
  checkTreeCoverage,
  DEFAULT_PARSE_LIMITS,
  type XDocument,
  type XNode,
  type XElement,
  type XText,
  type XCData,
  type XComment,
  type XProcessingInstruction,
  type XDeclaration,
  type XmlParseLimits,
  type CoverageGap,
} from './xnode.js';

export { serializeXmlString, serializeXml, serializeNode } from './serialize.js';

export { checkRoundTrip, type RoundTripDifference } from './roundtrip.js';

export { canonicalXml, firstDifference, excerpt, type CanonicalXmlOptions } from './canonical.js';

export {
  qualifiedKey,
  elementKey,
  childRanks,
  childRank,
  insertionIndex,
  isSchemaNamespace,
  outOfOrderChildren,
  SCHEMA_NAMESPACES,
  SCHEMA_SOURCES,
} from './schema-order.js';

export {
  MC_ALTERNATE_CONTENT,
  MC_CHOICE,
  MC_FALLBACK,
  isAlternateContent,
  ignorableNamespaces,
  mustUnderstandNamespaces,
  processContentNames,
  selectAlternateContent,
  effectiveChildren,
  effectiveAttributes,
  checkMarkupCompatibility,
  boundNamespaces,
  type McProblem,
} from './mce.js';

export {
  newAttribute,
  newElement,
  newText,
  applyEdit,
  applyEdits,
  insertInOrder,
  type XmlEdit,
  type DirtyRestore,
  type SetAttributeEdit,
  type AddAttributeEdit,
  type RemoveAttributeEdit,
  type InsertChildEdit,
  type RemoveChildEdit,
  type SetValueEdit,
} from './edit.js';

export {
  extensionList,
  extensions,
  findExtension,
  planAddExtension,
  planRemoveExtension,
} from './ext-lst.js';

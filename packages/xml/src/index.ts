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
 * As of sub-phase 0.4 this is the tokenizer and the node model. The serializer
 * and its byte-identical round-trip gate are 0.5; schema-ordered insertion, the
 * Markup Compatibility walker and the invertible edit operations are 0.6.
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
  prefixMap,
  undeclaredPrefixes,
  xmlSpace,
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

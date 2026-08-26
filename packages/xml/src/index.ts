/**
 * `@pptx-studio/xml` - byte-preserving XML for OOXML.
 *
 * The tokenizer, `XNode` and the serializer land in sub-phases 0.4-0.6. What is
 * here now is the namespace vocabulary they are written against, because it is
 * the one part of this package that is pure data and cannot be got wrong later
 * without breaking everything above it.
 */

/**
 * Namespace URIs that the tokenizer and serializer treat specially.
 *
 * A note that shapes the whole package: Markup Compatibility attributes hold
 * *prefixes*, not URIs - `mc:Ignorable="a14 p14"` and `mc:Choice Requires="a14"`
 * name prefixes declared elsewhere in the document. A serializer that is free to
 * rewrite prefixes (which the DOM spec permits, and which `XMLSerializer` does)
 * silently converts ignorable extension markup into a hard error. That is why
 * this package exists instead of a call to `DOMParser`.
 */
export const NS = {
  /** Bound to the `xml:` prefix implicitly; never declared. */
  xml: 'http://www.w3.org/XML/1998/namespace',
  /** The namespace of `xmlns` declarations themselves. */
  xmlns: 'http://www.w3.org/2000/xmlns/',
  /** Markup Compatibility and Extensibility (`mc:`). */
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  /** DrawingML main (`a:`). */
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  /** PresentationML main (`p:`). */
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  /** Relationship references inside part XML (`r:`). */
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
} as const;

export type NamespacePrefix = keyof typeof NS;
export type NamespaceUri = (typeof NS)[NamespacePrefix];

/**
 * `xml:space="preserve"`. Any `a:t` carrying leading or trailing whitespace
 * needs this or PowerPoint drops the whitespace on load.
 */
export const XML_SPACE_PRESERVE = 'preserve';

/**
 * The namespace vocabulary the tokenizer and serializer are written against.
 *
 * A note that shapes the whole package: Markup Compatibility attributes hold
 * *prefixes*, not URIs - `mc:Ignorable="a14 p14"` and `mc:Choice Requires="a14"`
 * name prefixes declared elsewhere in the document. A serializer that is free to
 * rewrite prefixes (which the DOM spec permits, and which `XMLSerializer` does)
 * silently converts ignorable extension markup into a hard error. That is why
 * this package exists instead of a call to `DOMParser`.
 *
 * The corpus contains four live examples, all `Requires` on an
 * `mc:AlternateContent`: three naming `c14` and one naming `p14`. The `p14` one
 * is the instructive case, because `p14` is bound to two different URIs in
 * different parts of the same corpus - so the prefix cannot be resolved without
 * knowing where in which document it was written.
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
 * `xml:space="preserve"`.
 *
 * Exported for reading, not for writing. PresentationML does not need it: our
 * corpus contains none, 120 `<a:t>` elements carry leading or trailing
 * whitespace without it, and PowerPoint round-trips such a part with the
 * whitespace intact.
 */
export const XML_SPACE_PRESERVE = 'preserve';

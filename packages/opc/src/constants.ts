/**
 * The fixed vocabulary of the package format: names and namespaces that are
 * defined by the spec and never vary.
 *
 * `PartStore`, content-type resolution and the writer land in 0.3.
 */

/**
 * The content-type stream. Always at the package root, always this exact name,
 * including the square brackets. It is not a "part" in the OPC sense - it has no
 * content type of its own and never appears in a relationship.
 */
export const CONTENT_TYPES_PART = '[Content_Types].xml';

/** The package-level relationship part: the only way in to a `.pptx`. */
export const ROOT_RELS_PART = '_rels/.rels';

/** Namespaces used by the two package-level parts above. */
export const OPC_NS = {
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  relationships: 'http://schemas.openxmlformats.org/package/2006/relationships',
} as const;

/**
 * Relationship type URIs resolved during load.
 *
 * `slideLayout` is listed here for a reason worth stating early: **a slide part
 * does not name its own layout**. The binding exists only as a relationship in
 * `ppt/slides/_rels/slideN.xml.rels`, which is why "change layout" is a
 * relationship rewrite plus a placeholder rebind rather than an edit to the
 * slide XML.
 */
export const REL_TYPE = {
  officeDocument:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  font: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
} as const;

export type RelationshipType = (typeof REL_TYPE)[keyof typeof REL_TYPE];

/**
 * The `Default` extension for embedded fonts, and its content type.
 *
 * Omitting this `Default` entry while shipping `ppt/fonts/*.fntdata` is the
 * canonical "PowerPoint found a problem with some content" bug. The payload
 * itself is EOT (Embedded OpenType) - not raw TrueType, and not Word's ODTTF
 * obfuscation. See docs/adr and sub-phase 8.2.
 */
export const FONT_DATA_EXTENSION = 'fntdata';
export const FONT_DATA_CONTENT_TYPE = 'application/x-fontdata';

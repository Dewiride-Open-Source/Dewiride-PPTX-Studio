/**
 * The fixed vocabulary of the package format: names and namespaces that are
 * defined by the spec and never vary.
 *
 * `PartStore`, content-type resolution and the writer land in 0.3.
 */

/**
 * The content-type stream, as a **ZIP entry name**.
 *
 * Always at the package root, always this exact name, brackets included. There
 * is no part-name form of it because it is not a part: it has no content type
 * of its own and nothing ever relates to it. That is why this constant is an
 * entry name where `ROOT_RELS_PART` below is a part name - the two live in
 * different namespaces and the asymmetry is the point rather than an oversight.
 */
export const CONTENT_TYPES_PART = '[Content_Types].xml';

/**
 * The package-level relationship part, as a **part name**: the only way in to
 * a `.pptx`.
 *
 * Its ZIP entry name is `zipEntryNameFor(ROOT_RELS_PART)`, without the leading
 * slash. Removing this part from a package makes PowerPoint report the file as
 * corrupted and unreadable, which we verified by doing it.
 */
export const ROOT_RELS_PART = '/_rels/.rels';

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
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships/';
const MS07 = 'http://schemas.microsoft.com/office/2007/relationships/';
const MS11 = 'http://schemas.microsoft.com/office/2011/relationships/';

export const REL_TYPE = {
  officeDocument: OD + 'officeDocument',
  slide: OD + 'slide',
  slideLayout: OD + 'slideLayout',
  slideMaster: OD + 'slideMaster',
  notesSlide: OD + 'notesSlide',
  notesMaster: OD + 'notesMaster',
  handoutMaster: OD + 'handoutMaster',
  theme: OD + 'theme',
  image: OD + 'image',
  font: OD + 'font',
  tags: OD + 'tags',
  presProps: OD + 'presProps',
  viewProps: OD + 'viewProps',
  tableStyles: OD + 'tableStyles',
  chart: OD + 'chart',
  /** The embedded workbook behind a chart. */
  package: OD + 'package',
  oleObject: OD + 'oleObject',
  diagramData: OD + 'diagramData',
  diagramLayout: OD + 'diagramLayout',
  diagramColors: OD + 'diagramColors',
  /** Note the name: the type is `diagramQuickStyle`, the part is `quickStyle`. */
  diagramQuickStyle: OD + 'diagramQuickStyle',
  diagramDrawing: MS07 + 'diagramDrawing',
  chartStyle: MS11 + 'chartStyle',
  chartColorStyle: MS11 + 'chartColorStyle',
  hyperlink: OD + 'hyperlink',
  video: OD + 'video',
  audio: OD + 'audio',
  media: OD + 'media',
  extendedProperties: OD + 'extended-properties',
  customProperties: OD + 'custom-properties',
  /**
   * The two below are in the **package** namespace, not `officeDocument`.
   *
   * That is not a typo and it is easy to get wrong: `extended-properties`
   * above is `officeDocument/...` while `core-properties` here is
   * `package/.../metadata/...`. Both appear side by side in every `_rels/.rels`
   * we have measured.
   */
  coreProperties: PKG + 'metadata/core-properties',
  thumbnail: PKG + 'metadata/thumbnail',
} as const;

export type RelationshipType = (typeof REL_TYPE)[keyof typeof REL_TYPE];

/**
 * Content-type strings, transcribed from packages Office wrote.
 *
 * Every one of these is compared verbatim against what a real file contains,
 * so they are copied rather than composed - a `+xml` in the wrong place is a
 * part that silently resolves to nothing.
 */
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
const DML = 'application/vnd.openxmlformats-officedocument.drawingml.';
const OFP = 'application/vnd.openxmlformats-officedocument.';
const PKGT = 'application/vnd.openxmlformats-package.';

export const CONTENT_TYPE = {
  relationships: PKGT + 'relationships+xml',
  coreProperties: PKGT + 'core-properties+xml',
  xml: 'application/xml',

  /** `.pptx`. A `.potx` is `template.main+xml` and the two are not interchangeable. */
  presentation: PML + 'presentation.main+xml',
  template: PML + 'template.main+xml',
  slideshow: PML + 'slideshow.main+xml',

  slide: PML + 'slide+xml',
  slideLayout: PML + 'slideLayout+xml',
  slideMaster: PML + 'slideMaster+xml',
  notesSlide: PML + 'notesSlide+xml',
  notesMaster: PML + 'notesMaster+xml',
  handoutMaster: PML + 'handoutMaster+xml',
  presProps: PML + 'presProps+xml',
  viewProps: PML + 'viewProps+xml',
  tableStyles: PML + 'tableStyles+xml',
  tags: PML + 'tags+xml',
  commentAuthors: PML + 'commentAuthors+xml',
  comments: PML + 'comments+xml',

  theme: OFP + 'theme+xml',
  extendedProperties: OFP + 'extended-properties+xml',
  customProperties: OFP + 'custom-properties+xml',
  spreadsheet: OFP + 'spreadsheetml.sheet',

  chart: DML + 'chart+xml',
  diagramData: DML + 'diagramData+xml',
  diagramLayout: DML + 'diagramLayout+xml',
  diagramStyle: DML + 'diagramStyle+xml',
  diagramColors: DML + 'diagramColors+xml',

  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
} as const;

export type ContentType = (typeof CONTENT_TYPE)[keyof typeof CONTENT_TYPE];

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

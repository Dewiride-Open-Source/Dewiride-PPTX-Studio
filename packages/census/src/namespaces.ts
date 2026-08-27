/**
 * The namespaces a census recognises by name.
 *
 * Two things about this table are deliberate.
 *
 * **It is keyed by URI, never by prefix.** `p14` is bound to two different URIs
 * in different parts of the same corpus, so a scanner that matched `p14:media`
 * on the spelling would be reading a coin flip. Every lookup here goes through
 * the prefix scope the element was actually written in.
 *
 * **Being absent from this table is not an error.** Anything unrecognised still
 * lands in the namespace histogram with its URI, its prefixes and a count. That
 * is the more useful half of the output: a census whose job is to tell you what
 * a deck contains cannot afford to be silent about the parts it has no name
 * for, and the histogram is how a namespace we have never seen becomes visible
 * instead of becoming nothing.
 *
 * URIs verified against `dotnet/Open-XML-SDK`'s `data/namespaces.json` (MIT,
 * read-for-facts — see LEGAL.md) rather than transcribed from memory.
 */

/** ECMA-376 namespaces: the standard vocabulary. */
export const OOXML_NS = {
  /** DrawingML main. Shapes, colours, text, geometry. */
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  /** PresentationML main. */
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  /** Relationship references *inside* part XML - `r:id`, `r:embed`, `r:link`. */
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  /** Markup Compatibility and Extensibility. */
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  /** DrawingML charts. */
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  /** SmartArt data, layout, colours and quick styles. */
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  /** The `p:pic` wrapper used inside a `a:graphicFrame`. */
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  /** SpreadsheetML - reachable through an embedded chart workbook. */
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  /** WordprocessingML - reachable through an embedded document. */
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  /** Office Math. */
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  /** `docProps/app.xml`. */
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  /** `docProps/core.xml`. */
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  /** The `a:lockedCanvas` wrapper. */
  lockedCanvas: 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas',
} as const;

/**
 * Microsoft extension namespaces, and the two VML ones.
 *
 * These are what live inside `mc:AlternateContent/mc:Choice/@Requires` and
 * inside `extLst`. Every entry here names a feature that arrived after
 * ECMA-376 was published, which is exactly why the census counts them: the
 * histogram of extension namespaces in a deck *is* the list of things a reader
 * built strictly to the standard would silently drop.
 */
export const EXTENSION_NS = {
  /** DrawingML 2010: `a14:imgProps`, `a14:hiddenFill`, `a14:cameraTool`. */
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  /** DrawingML 2014: `a16:creationId`, `a16:rowId`. */
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  /** PowerPoint 2010: media trim, sections, `p14:creationId`. */
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  /** PowerPoint 2012: `p15:sldGuideLst`, `p15:notesGuideLst`, threaded comments. */
  p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
  /** PowerPoint 2015. */
  p16: 'http://schemas.microsoft.com/office/powerpoint/2015/main',
  /** The older PowerPoint extension namespace, still present in real decks. */
  p2007: 'http://schemas.microsoft.com/office/powerpoint/2007/7/12/main',
  /** Chart 2007. */
  c14: 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
  /** Chart 2014. */
  c16: 'http://schemas.microsoft.com/office/drawing/2014/chart',
  /** ChartEx: waterfall, funnel, treemap, sunburst, histogram, box-whisker, map. */
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  /** Chart styles - `style1.xml`. */
  chartStyle: 'http://schemas.microsoft.com/office/drawing/2012/chartStyle',
  /** The SmartArt *drawing* fallback: `drawingN.xml`. */
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
  /** SmartArt 2016 extensions. */
  dgm1612: 'http://schemas.microsoft.com/office/drawing/2016/12/diagram',
  /** Scalable vector blips - `asvg:svgBlip`. */
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
  /** 3-D models - `am3d:model3d`. */
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  /** The decorative marker screen readers are meant to honour. */
  adec: 'http://schemas.microsoft.com/office/drawing/2017/decorative',
  /** Ink annotations wrapper. */
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  /** W3C InkML: the trace data itself. */
  inkml: 'http://www.w3.org/2003/InkML',
  /** VML, still the carrier for OLE object previews. */
  v: 'urn:schemas-microsoft-com:vml',
  /** The VML companion namespace. */
  o: 'urn:schemas-microsoft-com:office:office',
} as const;

/** Prefix conventionally used for a URI, for display only. Never for matching. */
export const CONVENTIONAL_PREFIX: ReadonlyMap<string, string> = new Map(
  [...Object.entries(OOXML_NS), ...Object.entries(EXTENSION_NS)].map(([key, uri]) => [uri, key]),
);

/** True for a namespace that ECMA-376 itself defines. */
export function isStandardNamespace(uri: string): boolean {
  return (Object.values(OOXML_NS) as string[]).includes(uri);
}

/** True for a namespace that postdates ECMA-376 - the extension surface. */
export function isExtensionNamespace(uri: string): boolean {
  return (Object.values(EXTENSION_NS) as string[]).includes(uri);
}

import { EXTENSION_NS, OOXML_NS } from './namespaces.js';

/**
 * What a census looks for, and which phase of the plan makes it render.
 *
 * The `phase` column is the point of this table. A viewer that renders half of
 * PowerPoint has to be able to say *which* half, per file, before it starts;
 * dropping a deck on the page and reading back "SmartArt 14 (phase 4.6),
 * ChartEx 3 (phase 10.2), ink 0" is a schedule, not a disclaimer. It is also
 * how the corpus manifest's `features[]` array gets filled in without anyone
 * curating it by hand.
 *
 * Two entries carry a phase of `preserve`. Animations and transitions are
 * never played and never authored here - they are carried across an export
 * byte-for-byte and nothing more. Saying so in the same table as everything
 * else keeps the promise auditable rather than buried in a scope document.
 */
export interface FeatureRule {
  readonly key: string;
  readonly label: string;
  /** Sub-phase that implements it, or `preserve` for carry-across-only. */
  readonly phase: string;
  readonly uri: string;
  readonly local: string;
}

const A = OOXML_NS.a;
const P = OOXML_NS.p;
const C = OOXML_NS.c;
const DGM = OOXML_NS.dgm;
const MC = OOXML_NS.mc;

export const FEATURE_RULES: readonly FeatureRule[] = [
  // --- the shape tree ------------------------------------------------------
  { key: 'shape', label: 'Shapes', phase: '2.10', uri: P, local: 'sp' },
  { key: 'group', label: 'Groups', phase: '2.10', uri: P, local: 'grpSp' },
  { key: 'picture', label: 'Pictures', phase: '2.10', uri: P, local: 'pic' },
  { key: 'connector', label: 'Connectors', phase: '2.10', uri: P, local: 'cxnSp' },
  { key: 'graphicFrame', label: 'Graphic frames', phase: '10.11', uri: P, local: 'graphicFrame' },
  { key: 'placeholder', label: 'Placeholders', phase: '7.1', uri: P, local: 'ph' },

  // --- geometry and paint --------------------------------------------------
  { key: 'presetGeom', label: 'Preset geometry', phase: '2.1', uri: A, local: 'prstGeom' },
  { key: 'customGeom', label: 'Custom geometry', phase: '2.4', uri: A, local: 'custGeom' },
  { key: 'gradientFill', label: 'Gradient fills', phase: '2.7', uri: A, local: 'gradFill' },
  { key: 'patternFill', label: 'Pattern fills', phase: '2.7', uri: A, local: 'pattFill' },
  { key: 'blipFill', label: 'Picture fills', phase: '2.10', uri: A, local: 'blipFill' },
  {
    // Not `noFill`. `a:grpFill` inherits the enclosing group's fill, and
    // treating the two as the same is how a themed shape renders empty.
    key: 'groupFill',
    label: 'Group fills',
    phase: '2.10',
    uri: A,
    local: 'grpFill',
  },
  { key: 'shadow', label: 'Outer shadows', phase: '2.8', uri: A, local: 'outerShdw' },
  { key: 'innerShadow', label: 'Inner shadows', phase: '2.8', uri: A, local: 'innerShdw' },
  { key: 'glow', label: 'Glows', phase: '2.8', uri: A, local: 'glow' },
  { key: 'softEdge', label: 'Soft edges', phase: '2.8', uri: A, local: 'softEdge' },
  { key: 'reflection', label: 'Reflections', phase: '2.8', uri: A, local: 'reflection' },
  { key: 'scene3d', label: '3-D scenes', phase: 'preserve', uri: A, local: 'scene3d' },

  // --- text ----------------------------------------------------------------
  { key: 'field', label: 'Text fields', phase: '3.5', uri: A, local: 'fld' },
  { key: 'hyperlink', label: 'Hyperlinks', phase: '3.5', uri: A, local: 'hlinkClick' },

  // --- hard content --------------------------------------------------------
  { key: 'table', label: 'Tables', phase: '4.1', uri: A, local: 'tbl' },
  { key: 'chart', label: 'Charts', phase: '9.1', uri: C, local: 'chartSpace' },
  {
    key: 'chartEx',
    label: 'ChartEx (waterfall, treemap, …)',
    phase: '10.1',
    uri: EXTENSION_NS.cx,
    local: 'chartSpace',
  },
  { key: 'smartArt', label: 'SmartArt', phase: '4.5', uri: DGM, local: 'relIds' },
  {
    key: 'smartArtDrawing',
    label: 'SmartArt drawing fallback',
    phase: '4.6',
    uri: EXTENSION_NS.dsp,
    local: 'drawing',
  },
  { key: 'oleObject', label: 'OLE objects', phase: '10.9', uri: P, local: 'oleObj' },
  { key: 'video', label: 'Video', phase: '10.8', uri: A, local: 'videoFile' },
  { key: 'audio', label: 'Audio', phase: '10.8', uri: A, local: 'audioFile' },
  {
    key: 'media',
    label: 'Media (2010 extension)',
    phase: '10.8',
    uri: EXTENSION_NS.p14,
    local: 'media',
  },
  { key: 'svgBlip', label: 'SVG blips', phase: '10.7', uri: EXTENSION_NS.asvg, local: 'svgBlip' },
  { key: 'model3d', label: '3-D models', phase: '10.10', uri: EXTENSION_NS.am3d, local: 'model3d' },
  { key: 'ink', label: 'Ink traces', phase: '10.10', uri: EXTENSION_NS.inkml, local: 'trace' },
  {
    key: 'contentPart',
    label: 'Content parts (ink hosts)',
    phase: '10.10',
    uri: P,
    local: 'contentPart',
  },
  { key: 'vml', label: 'VML drawings', phase: '10.9', uri: EXTENSION_NS.v, local: 'shape' },
  { key: 'math', label: 'Office Math', phase: 'preserve', uri: OOXML_NS.m, local: 'oMath' },

  // --- structure and extensibility ----------------------------------------
  {
    key: 'alternateContent',
    label: 'Markup Compatibility switches',
    phase: '0.6',
    uri: MC,
    local: 'AlternateContent',
  },
  { key: 'embeddedFont', label: 'Embedded fonts', phase: '8.2', uri: P, local: 'embeddedFont' },
  { key: 'section', label: 'Sections', phase: '7.4', uri: EXTENSION_NS.p14, local: 'section' },
  { key: 'customShow', label: 'Custom shows', phase: '5.6', uri: P, local: 'custShow' },
  {
    key: 'decorative',
    label: 'Decorative markers',
    phase: '12.3',
    uri: EXTENSION_NS.adec,
    local: 'decorative',
  },

  // --- preserved, never authored ------------------------------------------
  { key: 'animation', label: 'Animation timelines', phase: 'preserve', uri: P, local: 'timing' },
  { key: 'transition', label: 'Slide transitions', phase: 'preserve', uri: P, local: 'transition' },
] as const;

/**
 * Features detected from the package rather than from any element.
 *
 * A macro-enabled deck is a `vbaProject.bin` part and nothing else; there is no
 * markup anywhere that says so.
 */
export const PART_FEATURE_RULES: readonly {
  readonly key: string;
  readonly label: string;
  readonly phase: string;
  readonly match: (partName: string, contentType: string) => boolean;
}[] = [
  {
    key: 'macros',
    label: 'VBA macros',
    phase: 'preserve',
    match: (name) => name.toLowerCase().endsWith('/vbaproject.bin'),
  },
  {
    key: 'notesSlide',
    label: 'Notes slides',
    phase: '3.1',
    match: (_name, type) => type.endsWith('notesSlide+xml'),
  },
  {
    key: 'embeddedPackage',
    label: 'Embedded workbooks and documents',
    phase: '10.9',
    match: (name) => name.toLowerCase().startsWith('/ppt/embeddings/'),
  },
  {
    key: 'comment',
    label: 'Comments',
    phase: 'preserve',
    match: (_name, type) => type.endsWith('presentationml.comments+xml'),
  },
  {
    key: 'thumbnail',
    label: 'Package thumbnail',
    phase: 'preserve',
    match: (name) => name.toLowerCase() === '/docprops/thumbnail.jpeg',
  },
];

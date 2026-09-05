/**
 * Synthetic benchmark decks.
 *
 * Sub-phase 0.8's verification is "parse a 200 MB / 300-slide deck in a browser
 * tab and record the timing", and there is no such deck to hand. There is also
 * a rule in CLAUDE.md that says not to go looking for one on this machine. So
 * we make it, from a recipe, deterministically - the same recipe produces the
 * same bytes, which is what lets `corpus/bench/manifest.json` pin a SHA-256
 * against a file far too large to commit.
 *
 * Three things this generator is careful about, because getting any of them
 * wrong would produce a benchmark that measures the wrong work:
 *
 * - **The media is incompressible.** A deck padded with zeroes deflates to
 *   nothing, and the ZIP reader then never does what a real file makes it do.
 *   Every image is noise, stored rather than deflated.
 * - **The XML is varied.** A deck of 300 identical slides is a test of the
 *   branch predictor. Shape counts, geometry, fills, text length and which
 *   optional features appear all vary with the slide index.
 * - **The markup is real.** Content types, relationships and element order are
 *   what PowerPoint expects, so the deck can be opened in PowerPoint and the
 *   answer means something. Sub-phase 0.7 established that an unescaped `<` in
 *   an `a:t` makes PowerPoint refuse a whole package with no diagnostic, which
 *   is the only warning anyone gets that hand-written markup is not free.
 *
 * What it is not: a fidelity fixture. Nothing here is a design, and the corpus
 * of 50 real decks is sub-phase 1.1's job.
 */

import { readFileSync } from 'node:fs';
import { writeZip as writeMemoryZip } from '../ground-truth/lib/zip.ts';
import { mulberry32, noisePng, sizeForBytes } from './png.ts';
import type { ZipStream } from './zip-stream.ts';

// ---------------------------------------------------------------- namespaces

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const NS_DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const REL = NS_R + '/';
const REL_MS07 = 'http://schemas.microsoft.com/office/2007/relationships/';

const CT_PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
const CT_DML = 'application/vnd.openxmlformats-officedocument.drawingml.';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_DECLS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

const encoder = new TextEncoder();

export const SLIDE_WIDTH = 12192000;
export const SLIDE_HEIGHT = 6858000;
const NOTES_WIDTH = 6858000;
const NOTES_HEIGHT = 9144000;

/**
 * XML text escaping.
 *
 * Not optional and not a nicety. A literal `<` inside an `a:t` is what made
 * PowerPoint refuse an entire package in sub-phase 0.7, with the message "the
 * file or directory is corrupted and unreadable" and no part name.
 */
export function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ------------------------------------------------------------------- recipes

/**
 * When a feature appears.
 *
 * A number is "every Nth slide", which is the right idiom for a benchmark deck
 * spreading a feature evenly through three hundred of them. An array is the
 * exact 1-based slide numbers.
 *
 * The array form exists because the interval form cannot express the shape
 * every feature probe in the corpus has: *one* chart, on slide 1. `charts: 1`
 * means a chart on every slide, and there is no number that means "only the
 * first". Padding a deck out to twenty-five slides so that `charts: 25` finds
 * one is a fixture that tests the padding.
 */
export type Cadence = number | readonly number[];

/** Does the feature land on this 0-based slide index? */
function occurs(index: number, cadence: Cadence): boolean {
  return typeof cadence === 'number'
    ? cadence > 0 && (index + 1) % cadence === 0
    : cadence.includes(index + 1);
}

export interface DeckRecipe {
  /** Manifest identity and the key in `RECIPES`. Never appears in the file. */
  readonly id: string;
  /** `dc:title`. Separate from `id` because this one is read by people. */
  readonly title: string;
  readonly slides: number;
  /** Approximate bytes per generated image. 0 for a deck with no media at all. */
  readonly imageBytes: number;
  /** Free shapes on each slide, beyond the two placeholders. */
  readonly shapesPerSlide: number;
  /** Bullets in each slide's body placeholder. */
  readonly bulletsPerSlide: number;
  readonly notes: boolean;

  readonly tables: Cadence;
  readonly charts: Cadence;
  readonly diagrams: Cadence;
  readonly alternateContent: Cadence;
  readonly animations: Cadence;
  readonly hyperlinks: Cadence;
  readonly comments: Cadence;

  /** Clamped to the slide count: a section needs a slide to be a section of. */
  readonly sections: number;
  readonly customShows: number;
  /** Embed the CC0 probe font committed by sub-phase 0.7. */
  readonly embedFont: boolean;

  /** EMU. 12192000 x 6858000 is 16:9; 9144000 x 6858000 is 4:3. */
  readonly slideWidth: number;
  readonly slideHeight: number;
}

type RecipeDefaults = Omit<DeckRecipe, 'id' | 'title' | 'slides' | 'imageBytes' | 'shapesPerSlide'>;

const BASE: RecipeDefaults = {
  bulletsPerSlide: 5,
  notes: true,
  tables: 10,
  charts: 25,
  diagrams: 40,
  alternateContent: 7,
  animations: 5,
  hyperlinks: 12,
  comments: 15,
  sections: 6,
  customShows: 2,
  embedFont: true,
  slideWidth: SLIDE_WIDTH,
  slideHeight: SLIDE_HEIGHT,
};

export const RECIPES: Readonly<Record<string, DeckRecipe>> = {
  /**
   * The plan's own envelope: 300 slides, about 200 MB, dominated by media.
   *
   * This is the shape a real large deck has. It is the deck the sub-phase is
   * verified against, and it is mostly a test of the ZIP layer - the XML in it
   * is a rounding error beside the pictures.
   */
  'media-200mb': {
    ...BASE,
    id: 'media-200mb',
    title: 'PPTX Studio benchmark deck: media-200mb',
    slides: 300,
    imageBytes: 640 * 1024,
    shapesPerSlide: 12,
  },

  /**
   * The same 300 slides with no media and forty times the markup.
   *
   * The media deck cannot exercise the tokenizer: 195 MB of it never reaches
   * the XML layer at all. This one is small on disk and enormous inflated,
   * which is the case that finds the decompression budget and the case whose
   * throughput number actually predicts anything about Phase 2 onwards.
   */
  'xml-heavy': {
    ...BASE,
    id: 'xml-heavy',
    title: 'PPTX Studio benchmark deck: xml-heavy',
    slides: 300,
    imageBytes: 0,
    shapesPerSlide: 400,
    bulletsPerSlide: 24,
  },

  /**
   * Small, fast, and every feature at least twice.
   *
   * The one that gets opened in real PowerPoint by hand. The intervals are
   * tighter than the big recipes because "every 25th slide" finds no charts in
   * a deck of 20, and a deck that silently omits the feature you wanted to
   * check is worse than one that fails.
   */
  small: {
    ...BASE,
    id: 'small',
    title: 'PPTX Studio benchmark deck: small',
    slides: 20,
    imageBytes: 96 * 1024,
    shapesPerSlide: 8,
    sections: 3,
    tables: 5,
    charts: 6,
    diagrams: 8,
    alternateContent: 4,
    animations: 3,
    hyperlinks: 7,
  },
};

// ------------------------------------------------------------------- helpers

const LAYOUTS = [
  { type: 'title', name: 'Title Slide' },
  { type: 'obj', name: 'Title and Content' },
  { type: 'twoObj', name: 'Two Content' },
  { type: 'blank', name: 'Blank' },
] as const;

const WORDS = [
  'inheritance',
  'placeholder',
  'resolver',
  'geometry',
  'gradient',
  'preset',
  'relationship',
  'namespace',
  'tokenizer',
  'package',
  'serializer',
  'transform',
  'autofit',
  'theme',
  'layout',
  'master',
  'cascade',
  'provenance',
  'fidelity',
  'census',
  'ordinal',
  'traversal',
  'boundary',
];

function words(random: () => number, count: number): string {
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    out.push(WORDS[Math.floor(random() * WORDS.length)] ?? 'shape');
  }
  return out.join(' ');
}

/** A deterministic `ST_Guid`. PowerPoint rejects anything that is not this shape. */
function guid(seed: number): string {
  const random = mulberry32(seed);
  const hex = (n: number): string =>
    Array.from({ length: n }, () =>
      Math.floor(random() * 16)
        .toString(16)
        .toUpperCase(),
    ).join('');
  return '{' + hex(8) + '-' + hex(4) + '-' + hex(4) + '-' + hex(4) + '-' + hex(12) + '}';
}

interface Rel {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

function relsXml(entries: readonly Rel[]): string {
  return (
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    entries
      .map(
        (entry) =>
          `<Relationship Id="${entry.id}" Type="${entry.type}" Target="${entry.target}"` +
          (entry.external === true ? ' TargetMode="External"' : '') +
          '/>',
      )
      .join('') +
    '</Relationships>'
  );
}

function spTreeHead(): string {
  return (
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  );
}

function paragraph(text: string, level = 0): string {
  return (
    (level === 0 ? '<a:p>' : `<a:p><a:pPr lvl="${String(level)}"/>`) +
    `<a:r><a:rPr lang="en-GB" dirty="0"/><a:t>${escapeXml(text)}</a:t></a:r>` +
    '</a:p>'
  );
}

// --------------------------------------------------------------- fixed parts

function themeXml(name: string): string {
  const scheme = [
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>',
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>',
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>',
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>',
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>',
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>',
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>',
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>',
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>',
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>',
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>',
  ].join('');

  // Exactly three entries in each list: `fillRef/@idx` and `bgRef/@idx` are
  // 1-based indices into them, and a short list is an out-of-range reference on
  // the first slide that uses it.
  const fill = (i: number): string =>
    i === 0
      ? '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
      : '<a:gradFill rotWithShape="1"><a:gsLst>' +
        `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="${String(60000 - i * 10000)}"/></a:schemeClr></a:gs>` +
        `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="${String(100000 - i * 20000)}"/></a:schemeClr></a:gs>` +
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>';
  const line = (w: number): string =>
    `<a:ln w="${String(w)}" cap="flat" cmpd="sng" algn="ctr">` +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';

  return (
    DECLARATION +
    `<a:theme xmlns:a="${NS_A}" name="${name}">` +
    '<a:themeElements>' +
    `<a:clrScheme name="${name}">${scheme}</a:clrScheme>` +
    `<a:fontScheme name="${name}">` +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    `<a:fmtScheme name="${name}">` +
    `<a:fillStyleLst>${fill(0)}${fill(1)}${fill(2)}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">' +
    '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>' +
    '</a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>' +
    '</a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    '</a:theme>'
  );
}

function placeholder(
  id: number,
  name: string,
  type: string,
  idx: number | null,
  x: number,
  y: number,
  cx: number,
  cy: number,
  body: string,
): string {
  const ph = `<p:ph type="${type}"` + (idx === null ? '' : ` idx="${String(idx)}"`) + '/>';
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/>` +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr>${ph}</p:nvPr>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    '</p:spPr>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody>` +
    '</p:sp>'
  );
}

function textStyle(kind: 'title' | 'body' | 'other'): string {
  const level = (n: number, size: number): string =>
    `<a:lvl${String(n)}pPr marL="${String(n === 1 ? 0 : 342900 * (n - 1))}" indent="${String(n === 1 ? 0 : -342900)}">` +
    '<a:defRPr sz="' +
    String(size) +
    '" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
    `<a:latin typeface="+m${kind === 'title' ? 'j' : 'n'}-lt"/></a:defRPr></a:lvl${String(n)}pPr>`;
  const sizes = kind === 'title' ? [4400] : [2800, 2400, 2000, 1800, 1800];
  return (
    `<p:${kind}Style>` +
    sizes.map((size, index) => level(index + 1, size)).join('') +
    `</p:${kind}Style>`
  );
}

function slideMasterXml(layoutCount: number): string {
  const layoutIds = Array.from(
    { length: layoutCount },
    (_, i) => `<p:sldLayoutId id="${String(2147483649 + i)}" r:id="rId${String(i + 1)}"/>`,
  ).join('');
  return (
    DECLARATION +
    `<p:sldMaster ${NS_DECLS}>` +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    spTreeHead() +
    placeholder(
      2,
      'Title Placeholder 1',
      'title',
      null,
      838200,
      365125,
      10515600,
      1325563,
      '<a:p><a:endParaRPr lang="en-GB"/></a:p>',
    ) +
    placeholder(
      3,
      'Text Placeholder 2',
      'body',
      1,
      838200,
      1825625,
      10515600,
      4351338,
      '<a:p><a:endParaRPr lang="en-GB"/></a:p>',
    ) +
    '</p:spTree></p:cSld>' +
    // All twelve clrMap attributes. Eleven is a repair prompt.
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
    ' hlink="hlink" folHlink="folHlink"/>' +
    `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>` +
    `<p:txStyles>${textStyle('title')}${textStyle('body')}${textStyle('other')}</p:txStyles>` +
    '</p:sldMaster>'
  );
}

function slideLayoutXml(index: number): string {
  const layout = LAYOUTS[index] ?? LAYOUTS[3];
  const empty = '<a:p><a:endParaRPr lang="en-GB"/></a:p>';
  let shapes = '';
  if (layout.type === 'title') {
    shapes =
      placeholder(2, 'Title 1', 'ctrTitle', null, 1524000, 1122363, 9144000, 2387600, empty) +
      placeholder(3, 'Subtitle 2', 'subTitle', 1, 1524000, 3602038, 9144000, 1655762, empty);
  } else if (layout.type === 'obj') {
    shapes =
      placeholder(2, 'Title 1', 'title', null, 838200, 365125, 10515600, 1325563, empty) +
      placeholder(3, 'Content Placeholder 2', 'body', 1, 838200, 1825625, 10515600, 4351338, empty);
  } else if (layout.type === 'twoObj') {
    shapes =
      placeholder(2, 'Title 1', 'title', null, 838200, 365125, 10515600, 1325563, empty) +
      placeholder(3, 'Content Placeholder 2', 'body', 1, 838200, 1825625, 5181600, 4351338, empty) +
      placeholder(4, 'Content Placeholder 3', 'body', 2, 6172200, 1825625, 5181600, 4351338, empty);
  }
  return (
    DECLARATION +
    `<p:sldLayout ${NS_DECLS} type="${layout.type}" preserve="1">` +
    `<p:cSld name="${layout.name}">` +
    spTreeHead() +
    shapes +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sldLayout>'
  );
}

function notesMasterXml(): string {
  return (
    DECLARATION +
    `<p:notesMaster ${NS_DECLS}>` +
    '<p:cSld>' +
    spTreeHead() +
    placeholder(
      2,
      'Notes Placeholder 1',
      'body',
      1,
      685800,
      4400550,
      5486400,
      4114800,
      '<a:p><a:endParaRPr lang="en-GB"/></a:p>',
    ) +
    '</p:spTree></p:cSld>' +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
    ' hlink="hlink" folHlink="folHlink"/>' +
    '<p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle>' +
    '</p:notesMaster>'
  );
}

// ------------------------------------------------------------- slide content

function tableXml(id: number, rows: number, columns: number, random: () => number): string {
  const width = Math.floor(9144000 / columns);
  const grid = Array.from({ length: columns }, () => `<a:gridCol w="${String(width)}"/>`).join('');
  const body = Array.from(
    { length: rows },
    () =>
      `<a:tr h="370840">` +
      Array.from(
        { length: columns },
        () =>
          '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' +
          paragraph(words(random, 2)) +
          '</a:txBody><a:tcPr/></a:tc>',
      ).join('') +
      '</a:tr>',
  ).join('');
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="${String(id)}" name="Table ${String(id)}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
    '</p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="1524000" y="1825625"/><a:ext cx="9144000" cy="2965120"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
    '<a:tbl><a:tblPr firstRow="1" bandRow="1"/>' +
    `<a:tblGrid>${grid}</a:tblGrid>${body}</a:tbl>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

function chartFrameXml(id: number, relId: string): string {
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="${String(id)}" name="Chart ${String(id)}"/>` +
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="1524000" y="1825625"/><a:ext cx="9144000" cy="4114800"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${relId}"/>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

function diagramFrameXml(id: number, ids: readonly [string, string, string, string]): string {
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="${String(id)}" name="Diagram ${String(id)}"/>` +
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    '<p:xfrm><a:off x="1524000" y="1825625"/><a:ext cx="9144000" cy="4114800"/></p:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
    `<dgm:relIds xmlns:dgm="${NS_DGM}" xmlns:r="${NS_R}"` +
    ` r:dm="${ids[0]}" r:lo="${ids[1]}" r:qs="${ids[2]}" r:cs="${ids[3]}"/>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

function pictureXml(id: number, relId: string, x: number, y: number, counters: Counters): string {
  counters.presetGeoms += 1;
  return (
    '<p:pic><p:nvPicPr>' +
    `<p:cNvPr id="${String(id)}" name="Picture ${String(id)}" descr="Synthetic noise, not a photograph"/>` +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>' +
    '</p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    '<p:spPr>' +
    `<a:xfrm><a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="2743200" cy="2743200"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</p:spPr></p:pic>'
  );
}

const PRESETS = [
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'diamond',
  'pentagon',
  'hexagon',
  'star5',
  'rightArrow',
  'chevron',
  'cloud',
  'heart',
  'donut',
  'plaque',
  'wedgeRoundRectCallout',
];

function freeShapeXml(
  id: number,
  index: number,
  random: () => number,
  hyperlink: string | null,
  counters: Counters,
  slideWidth: number,
  slideHeight: number,
): string {
  counters.presetGeoms += 1;
  if (index % 3 === 2) counters.patternFills += 1;
  const preset = PRESETS[index % PRESETS.length] ?? 'rect';
  // The scatter spans the slide less a margin and the shape’s own extent.
  // At 16:9 these are the 9500000 x 5500000 the benchmark decks were built
  // with, so the pinned hashes are unaffected.
  const spanX = Math.max(1, slideWidth - 2692000);
  const spanY = Math.max(1, slideHeight - 1358000);
  const x = 200000 + ((index * 937) % spanX);
  const y = 400000 + ((index * 613) % spanY);
  const accent = 'accent' + String((index % 6) + 1);

  // Three fill kinds in rotation, so the census sees more than solid colour and
  // Phase 2 has something to render other than flat rectangles.
  const fill =
    index % 3 === 0
      ? `<a:solidFill><a:schemeClr val="${accent}"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>`
      : index % 3 === 1
        ? '<a:gradFill><a:gsLst>' +
          `<a:gs pos="0"><a:schemeClr val="${accent}"><a:tint val="66000"/></a:schemeClr></a:gs>` +
          `<a:gs pos="100000"><a:schemeClr val="${accent}"><a:shade val="70000"/></a:schemeClr></a:gs>` +
          '</a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>'
        : `<a:pattFill prst="ltUpDiag"><a:fgClr><a:schemeClr val="${accent}"/></a:fgClr>` +
          '<a:bgClr><a:schemeClr val="bg1"/></a:bgClr></a:pattFill>';

  const link = hyperlink === null ? '' : `<a:hlinkClick xmlns:r="${NS_R}" r:id="${hyperlink}"/>`;

  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="Shape ${String(id)}">${link}</p:cNvPr>` +
    '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr>' +
    `<a:xfrm rot="${String((index * 300000) % 21600000)}">` +
    `<a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="685800" cy="457200"/></a:xfrm>` +
    `<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>` +
    fill +
    `<a:ln w="12700"><a:solidFill><a:schemeClr val="${accent}"><a:shade val="50000"/></a:schemeClr></a:solidFill>` +
    '<a:prstDash val="dash"/></a:ln>' +
    '</p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square"><a:normAutofit fontScale="92500"/></a:bodyPr><a:lstStyle/>' +
    paragraph(words(random, 2)) +
    '</p:txBody></p:sp>'
  );
}

/**
 * A text shape whose run carries a hyperlink.
 *
 * The host for a link on a slide that has no free shapes. `a:hlinkClick` hangs
 * off `a:rPr` here rather than off `p:cNvPr` as it does on a free shape, which
 * is the other of the two places DrawingML puts it and worth having in a deck
 * for that reason alone.
 */
function linkShapeXml(id: number, relId: string, counters: Counters): string {
  counters.presetGeoms += 1;
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="Link ${String(id)}"/>` +
    '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="838200" y="5600000"/><a:ext cx="4000000" cy="365125"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>' +
    `<a:rPr lang="en-GB" dirty="0"><a:hlinkClick xmlns:r="${NS_R}" r:id="${relId}"/></a:rPr>` +
    '<a:t>PPTX Studio on GitHub</a:t></a:r></a:p></p:txBody></p:sp>'
  );
}

/**
 * An `mc:AlternateContent` around a shape.
 *
 * `Requires` names a *prefix*, declared on the `mc:AlternateContent` itself.
 * A reader that resolved it anywhere else - or that let a serializer rename the
 * prefix - would turn ignorable extension markup into a hard error, which is
 * the whole reason this project does not use `DOMParser`.
 */
function alternateContentXml(id: number, counters: Counters): string {
  counters.presetGeoms += 2;
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}" xmlns:a14="${NS_A14}">` +
    '<mc:Choice Requires="a14">' +
    `<p:sp><p:nvSpPr><p:cNvPr id="${String(id)}" name="Extension shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    '<p:spPr><a:xfrm><a:off x="9000000" y="600000"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>' +
    `<a:extLst><a:ext uri="${guid(id * 31)}"><a14:hiddenFill xmlns:a14="${NS_A14}">` +
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a14:hiddenFill></a:ext></a:extLst>' +
    '</p:spPr></p:sp>' +
    '</mc:Choice>' +
    '<mc:Fallback>' +
    `<p:sp><p:nvSpPr><p:cNvPr id="${String(id)}" name="Fallback shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    '<p:spPr><a:xfrm><a:off x="9000000" y="600000"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr></p:sp>' +
    '</mc:Fallback>' +
    '</mc:AlternateContent>'
  );
}

/** A `p:timing` tree: the deepest markup in a real deck, and never rendered here. */
function timingXml(targetId: number): string {
  return (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">' +
    '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">' +
    '<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
    '<p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
    '<p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">' +
    '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    '<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>' +
    `<p:tgtEl><p:spTgt spid="${String(targetId)}"/></p:tgtEl>` +
    '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>' +
    '<p:to><p:strVal val="visible"/></p:to></p:set>' +
    '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst>' +
    '</p:cTn></p:par></p:childTnLst></p:cTn>' +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
    '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
  );
}

function chartXml(seed: number): string {
  const random = mulberry32(seed);
  const categories = ['North', 'South', 'East', 'West'];
  const values = categories.map(() => Math.round(random() * 900) / 10);
  const point = (index: number, value: string): string =>
    `<c:pt idx="${String(index)}"><c:v>${value}</c:v></c:pt>`;
  return (
    DECLARATION +
    `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    '<c:chart><c:plotArea><c:layout/>' +
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
    '<c:ser><c:idx val="0"/><c:order val="0"/>' +
    '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>' +
    point(0, 'Revenue') +
    '</c:strCache></c:strRef></c:tx>' +
    '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache>' +
    `<c:ptCount val="${String(categories.length)}"/>` +
    categories.map((name, index) => point(index, name)).join('') +
    '</c:strCache></c:strRef></c:cat>' +
    '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode>' +
    `<c:ptCount val="${String(values.length)}"/>` +
    values.map((value, index) => point(index, String(value))).join('') +
    '</c:numCache></c:numRef></c:val>' +
    '</c:ser><c:gapWidth val="150"/>' +
    '<c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>' +
    '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>' +
    '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>' +
    '</c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    '</c:chartSpace>'
  );
}

function diagramParts(
  seed: number,
  counters: Counters,
): Record<'data' | 'layout' | 'quickStyle' | 'colors' | 'drawing', string> {
  const nodes = ['Discover', 'Design', 'Build', 'Verify'];
  // One `a:prstGeom` per `dsp:sp` in the drawing fallback.
  counters.presetGeoms += nodes.length;
  const points = nodes
    .map(
      (label, index) =>
        `<dgm:pt modelId="${guid(seed + index)}"><dgm:prSet/><dgm:spPr/>` +
        `<dgm:t><a:bodyPr/><a:lstStyle/>${paragraph(label)}</dgm:t></dgm:pt>`,
    )
    .join('');
  return {
    data:
      DECLARATION +
      `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
      `<dgm:ptLst><dgm:pt modelId="${guid(seed + 90)}" type="doc"><dgm:prSet/></dgm:pt>${points}</dgm:ptLst>` +
      '<dgm:cxnLst/><dgm:bg/><dgm:whole/>' +
      '</dgm:dataModel>',
    layout:
      DECLARATION +
      `<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:pptx-studio/bench/process">` +
      '<dgm:title val="Bench process"/><dgm:desc val="A synthetic diagram"/>' +
      '<dgm:catLst><dgm:cat type="process" pri="1000"/></dgm:catLst>' +
      '<dgm:sampData/><dgm:styleData/><dgm:clrData/>' +
      '<dgm:layoutNode name="root"><dgm:alg type="lin"/><dgm:shape xmlns:r="' +
      NS_R +
      '"/><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>' +
      '</dgm:layoutDef>',
    quickStyle:
      DECLARATION +
      `<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:pptx-studio/bench/style">` +
      '<dgm:title val="Bench style"/><dgm:desc val="A synthetic quick style"/>' +
      '<dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>' +
      '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>' +
      '<dgm:styleLbl name="node0"><dgm:scene3d><a:camera prst="orthographicFront"/>' +
      '<a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/>' +
      '<dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
      '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
      '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>' +
      '</dgm:styleDef>',
    colors:
      DECLARATION +
      `<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:pptx-studio/bench/colors">` +
      '<dgm:title val="Bench colours"/><dgm:desc val="A synthetic colour set"/>' +
      '<dgm:catLst><dgm:cat type="accent1" pri="11100"/></dgm:catLst>' +
      '<dgm:styleLbl name="node0"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>' +
      '<dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>' +
      '<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/></dgm:styleLbl>' +
      '</dgm:colorsDef>',
    drawing:
      DECLARATION +
      `<dsp:drawing xmlns:dsp="${NS_DSP}" xmlns:a="${NS_A}">` +
      '<dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>' +
      nodes
        .map(
          (label, index) =>
            `<dsp:sp modelId="${guid(seed + index)}"><dsp:nvSpPr>` +
            `<dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr>` +
            `<a:xfrm><a:off x="${String(index * 2286000)}" y="0"/><a:ext cx="2057400" cy="1143000"/></a:xfrm>` +
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
            '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></dsp:spPr>' +
            `<dsp:txBody><a:bodyPr/><a:lstStyle/>${paragraph(label)}</dsp:txBody>` +
            // The text rectangle is separate from the shape rectangle. Ignoring
            // it puts text in the wrong place inside every chevron.
            `<dsp:txXfrm><a:off x="${String(index * 2286000 + 45720)}" y="45720"/>` +
            '<a:ext cx="1965960" cy="1051560"/></dsp:txXfrm></dsp:sp>',
        )
        .join('') +
      '</dsp:spTree></dsp:drawing>',
  };
}

/** A minimal workbook, so a chart's `package` relationship lands on a real file. */
function workbookParts(): { name: string; content: string }[] {
  return [
    {
      name: '[Content_Types].xml',
      content:
        DECLARATION +
        `<Types xmlns="${NS_CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '</Types>',
    },
    {
      name: '_rels/.rels',
      content: relsXml([{ id: 'rId1', type: REL + 'officeDocument', target: 'xl/workbook.xml' }]),
    },
    {
      name: 'xl/workbook.xml',
      content:
        DECLARATION +
        `<workbook xmlns="${NS_X}" xmlns:r="${NS_R}">` +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      content: relsXml([{ id: 'rId1', type: REL + 'worksheet', target: 'worksheets/sheet1.xml' }]),
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      content:
        DECLARATION +
        `<worksheet xmlns="${NS_X}"><sheetData>` +
        '<row r="1"><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row>' +
        '</sheetData></worksheet>',
    },
  ];
}

/**
 * A whole ZIP in memory, for the workbook that goes inside the deck.
 *
 * `ZipStream` writes to a file descriptor, which is exactly what a 200 MB deck
 * needs and exactly what a 2 KB nested archive cannot use. `tools/ground-truth`
 * already has an in-memory writer built for small archives, so this reuses it
 * rather than growing a third.
 */
function writeInnerZip(parts: readonly { name: string; content: string }[]): Uint8Array {
  return writeMemoryZip(
    parts.map((part) => ({ name: part.name, bytes: encoder.encode(part.content) })),
  );
}

// ------------------------------------------------------------------- writing

export interface DeckSummary {
  readonly recipe: DeckRecipe;
  readonly bytes: number;
  readonly entries: number;
  readonly imagePixels: number;
  /** Shapes across every slide, as the generator allocated ids for them. */
  readonly slideShapes: number;
  readonly expected: DeckExpectation;
}

/**
 * What the generator put in, tallied where it was written.
 *
 * Every field is incremented at the point the markup is emitted, never derived
 * afterwards from the recipe. That distinction is not stylistic: `section` and
 * `customShow` used to be read back off the recipe while `writeDeck` clamped
 * both to the slide count, so a recipe asking for six sections in a one-slide
 * deck declared six and wrote one.
 */
export interface Counters {
  /** Shapes on slides only. The census counts masters, layouts and notes too. */
  slideShapes: number;
  tables: number;
  charts: number;
  diagrams: number;
  alternates: number;
  animations: number;
  hyperlinks: number;
  pictures: number;
  fields: number;
  graphicFrames: number;
  embeddedPackages: number;
  comments: number;
  notesSlides: number;
  /**
   * `a:prstGeom` and `a:pattFill` come from several emitters and are switched
   * off by several different recipe fields, so whether a deck has any is not
   * something a predicate over the recipe gets right for long. Counting where
   * the markup is written does.
   */
  presetGeoms: number;
  patternFills: number;
  sections: number;
  customShows: number;
  embeddedFonts: number;
}

export function createCounters(): Counters {
  return {
    slideShapes: 0,
    tables: 0,
    charts: 0,
    diagrams: 0,
    alternates: 0,
    animations: 0,
    hyperlinks: 0,
    pictures: 0,
    fields: 0,
    graphicFrames: 0,
    embeddedPackages: 0,
    comments: 0,
    notesSlides: 0,
    presetGeoms: 0,
    patternFills: 0,
    sections: 0,
    customShows: 0,
    embeddedFonts: 0,
  };
}

/**
 * The generator's statement of what is in the deck it just wrote.
 *
 * Two halves, because two different things are knowable. `exact` is for the
 * features a recipe decides the number of - it asked for four tables, it got
 * four tables. `present` is for the ones that come with the scaffolding: the
 * theme's `effectStyleLst` carries an `a:outerShdw` whether anyone wanted one,
 * and the count of `a:prstGeom` across a package depends on how many shapes a
 * SmartArt fallback happens to draw. Pinning a number to those would be
 * inventing precision; asserting nothing about them would leave a hole exactly
 * where a miscount could hide.
 *
 * Between them the two halves must name **every** key the census can report,
 * including the ones this generator cannot produce, which are declared `0`.
 * `bench.test.ts` asserts that completeness against the census's own table, so
 * a feature rule added there fails this contract until the generator has an
 * opinion about it.
 */
export interface DeckExpectation {
  /** Keys whose count the recipe determines exactly. */
  readonly exact: Readonly<Record<string, number>>;
  /** Keys guaranteed to appear at least once, without pinning a count. */
  readonly present: readonly string[];
}

/**
 * The features this recipe puts in, counted by the generator.
 *
 * Written down separately from the deck so the census has something to be wrong
 * against. Two independent statements of what is in a file catch a miscount in
 * either one; a census checked only against itself catches nothing.
 *
 * It takes the counters and deliberately not the recipe. Every number here has
 * to come from where the markup was written; reading one back off the recipe is
 * what made a one-slide deck declare six sections and write one.
 */
export function expectedFeatures(counters: Counters): DeckExpectation {
  const present: string[] = [];

  // Unconditional scaffolding. Every deck has a master, every master has two
  // placeholders, and `themeXml` writes an `a:gradFill` into `fillStyleLst` and
  // an `a:outerShdw` into `effectStyleLst` whether or not anything uses them.
  // These four cannot be switched off by any recipe, so they need no counter.
  present.push('shape', 'placeholder', 'gradientFill', 'shadow');

  // These two can be switched off, and by more than one field each, so they are
  // counted where they are written rather than predicted from the recipe.
  if (counters.presetGeoms > 0) present.push('presetGeom');
  if (counters.patternFills > 0) present.push('patternFill');

  return {
    exact: {
      // --- what the recipe counts ------------------------------------------
      table: counters.tables,
      chart: counters.charts,
      smartArt: counters.diagrams,
      smartArtDrawing: counters.diagrams,
      alternateContent: counters.alternates,
      animation: counters.animations,
      hyperlink: counters.hyperlinks,
      picture: counters.pictures,
      field: counters.fields,
      graphicFrame: counters.graphicFrames,
      embeddedPackage: counters.embeddedPackages,
      comment: counters.comments,
      notesSlide: counters.notesSlides,
      section: counters.sections,
      customShow: counters.customShows,
      embeddedFont: counters.embeddedFonts,

      // --- what this generator does not write yet ---------------------------
      // Declared rather than omitted: a zero is a statement that the feature
      // was considered and left out, and it is what makes the census check
      // bidirectional. Sub-phase 1.1's corpus decks fill these in.
      group: 0,
      connector: 0,
      customGeom: 0,
      blipFill: 0,
      groupFill: 0,
      innerShadow: 0,
      glow: 0,
      softEdge: 0,
      reflection: 0,
      scene3d: 0,
      chartEx: 0,
      oleObject: 0,
      video: 0,
      audio: 0,
      media: 0,
      svgBlip: 0,
      model3d: 0,
      ink: 0,
      contentPart: 0,
      vml: 0,
      math: 0,
      decorative: 0,
      transition: 0,
      macros: 0,
      thumbnail: 0,
    },
    present,
  };
}

/** Write a whole deck into an open archive. Streams: nothing is held but one part. */
export function writeDeck(
  zip: ZipStream,
  recipe: DeckRecipe,
  fontPath: string | null,
): DeckSummary {
  const layoutCount = LAYOUTS.length;
  const imageSize = recipe.imageBytes > 0 ? sizeForBytes(recipe.imageBytes) : 0;
  const counters = createCounters();

  // Which slides get what has to be known before `[Content_Types].xml` is
  // written, because that part is first in the archive and names every override
  // in it. Deciding twice - once here and once while writing - is how a content
  // type and a part drift apart.
  const slideIndices = Array.from({ length: recipe.slides }, (_, i) => i);
  const chartSlides = slideIndices.filter((i) => occurs(i, recipe.charts));
  const diagramSlides = slideIndices.filter((i) => occurs(i, recipe.diagrams));
  const commentSlides = slideIndices.filter((i) => occurs(i, recipe.comments));

  const overrides: string[] = [];
  const override = (partName: string, contentType: string): void => {
    overrides.push(`<Override PartName="${partName}" ContentType="${contentType}"/>`);
  };

  override('/ppt/presentation.xml', CT_PML + 'presentation.main+xml');
  override('/ppt/presProps.xml', CT_PML + 'presProps+xml');
  override('/ppt/viewProps.xml', CT_PML + 'viewProps+xml');
  override('/ppt/tableStyles.xml', CT_PML + 'tableStyles+xml');
  override('/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
  override('/ppt/slideMasters/slideMaster1.xml', CT_PML + 'slideMaster+xml');
  for (let i = 0; i < layoutCount; i++) {
    override(`/ppt/slideLayouts/slideLayout${String(i + 1)}.xml`, CT_PML + 'slideLayout+xml');
  }
  if (recipe.notes) {
    override('/ppt/theme/theme2.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
    override('/ppt/notesMasters/notesMaster1.xml', CT_PML + 'notesMaster+xml');
  }
  for (const i of slideIndices) {
    override(`/ppt/slides/slide${String(i + 1)}.xml`, CT_PML + 'slide+xml');
    if (recipe.notes) {
      override(`/ppt/notesSlides/notesSlide${String(i + 1)}.xml`, CT_PML + 'notesSlide+xml');
    }
  }
  for (const [n] of chartSlides.entries()) {
    override(`/ppt/charts/chart${String(n + 1)}.xml`, CT_DML + 'chart+xml');
  }
  for (const [n] of diagramSlides.entries()) {
    const k = String(n + 1);
    override(`/ppt/diagrams/data${k}.xml`, CT_DML + 'diagramData+xml');
    override(`/ppt/diagrams/layout${k}.xml`, CT_DML + 'diagramLayout+xml');
    override(`/ppt/diagrams/quickStyle${k}.xml`, CT_DML + 'diagramStyle+xml');
    override(`/ppt/diagrams/colors${k}.xml`, CT_DML + 'diagramColors+xml');
    override(`/ppt/diagrams/drawing${k}.xml`, CT_DML + 'diagramDrawing+xml');
  }
  if (commentSlides.length > 0) override('/ppt/commentAuthors.xml', CT_PML + 'commentAuthors+xml');
  for (const [n] of commentSlides.entries()) {
    override(`/ppt/comments/comment${String(n + 1)}.xml`, CT_PML + 'comments+xml');
  }
  override('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
  override(
    '/docProps/app.xml',
    'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  );

  const defaults =
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    (imageSize > 0 ? '<Default Extension="png" ContentType="image/png"/>' : '') +
    (chartSlides.length > 0
      ? '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
      : '') +
    // Omitting this while shipping .fntdata parts is the canonical
    // "PowerPoint found a problem with some content" bug.
    (recipe.embedFont && fontPath !== null
      ? '<Default Extension="fntdata" ContentType="application/x-fontdata"/>'
      : '');

  zip.add(
    '[Content_Types].xml',
    DECLARATION + `<Types xmlns="${NS_CT}">` + defaults + overrides.join('') + '</Types>',
  );

  zip.add(
    '_rels/.rels',
    relsXml([
      { id: 'rId1', type: REL + 'officeDocument', target: 'ppt/presentation.xml' },
      {
        id: 'rId2',
        type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
        target: 'docProps/core.xml',
      },
      { id: 'rId3', type: REL + 'extended-properties', target: 'docProps/app.xml' },
    ]),
  );

  // --- presentation.xml and its relationships ------------------------------
  const presRels: Rel[] = [];
  let rid = 0;
  const nextRid = (): string => 'rId' + String(++rid);

  const masterRid = nextRid();
  presRels.push({
    id: masterRid,
    type: REL + 'slideMaster',
    target: 'slideMasters/slideMaster1.xml',
  });
  const notesMasterRid = recipe.notes ? nextRid() : null;
  if (notesMasterRid !== null) {
    presRels.push({
      id: notesMasterRid,
      type: REL + 'notesMaster',
      target: 'notesMasters/notesMaster1.xml',
    });
  }
  const slideRids = slideIndices.map((i) => {
    const id = nextRid();
    presRels.push({ id, type: REL + 'slide', target: `slides/slide${String(i + 1)}.xml` });
    return id;
  });
  presRels.push({ id: nextRid(), type: REL + 'presProps', target: 'presProps.xml' });
  presRels.push({ id: nextRid(), type: REL + 'viewProps', target: 'viewProps.xml' });
  presRels.push({ id: nextRid(), type: REL + 'theme', target: 'theme/theme1.xml' });
  presRels.push({ id: nextRid(), type: REL + 'tableStyles', target: 'tableStyles.xml' });
  if (commentSlides.length > 0) {
    // The author list hangs off the presentation, not off a slide. Left out of
    // the presentation's relationships it is an orphan part - which is exactly
    // what the census reported on the first run of this generator.
    presRels.push({ id: nextRid(), type: REL + 'commentAuthors', target: 'commentAuthors.xml' });
  }

  const fontBytes = recipe.embedFont && fontPath !== null ? readFileSync(fontPath) : null;
  let fontRid: string | null = null;
  if (fontBytes !== null) {
    fontRid = nextRid();
    presRels.push({ id: fontRid, type: REL + 'font', target: 'fonts/font1.fntdata' });
    counters.embeddedFonts = 1;
  }

  const sectionCount = Math.min(recipe.sections, recipe.slides);
  counters.sections = sectionCount;
  const perSection = sectionCount > 0 ? Math.ceil(recipe.slides / sectionCount) : 0;
  const sectionsXml =
    sectionCount === 0
      ? ''
      : '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">' +
        `<p14:sectionLst xmlns:p14="${NS_P14}">` +
        Array.from({ length: sectionCount }, (_, s) => {
          const ids = slideIndices
            .slice(s * perSection, (s + 1) * perSection)
            .map((i) => `<p14:sldId id="${String(256 + i)}"/>`)
            .join('');
          return (
            `<p14:section name="Section ${String(s + 1)}" id="${guid(1000 + s)}">` +
            `<p14:sldIdLst>${ids}</p14:sldIdLst></p14:section>`
          );
        }).join('') +
        '</p14:sectionLst></p:ext></p:extLst>';

  const customShowCount = Math.min(recipe.customShows, recipe.slides);
  counters.customShows = customShowCount;
  const customShowsXml =
    customShowCount === 0
      ? ''
      : '<p:custShowLst>' +
        Array.from({ length: customShowCount }, (_, s) => {
          const picked = slideRids
            .filter((_, i) => i % (s + 2) === 0)
            .slice(0, 8)
            .map((id) => `<p:sld r:id="${id}"/>`)
            .join('');
          return (
            `<p:custShow name="Show ${String(s + 1)}" id="${String(s)}">` +
            `<p:sldLst>${picked}</p:sldLst></p:custShow>`
          );
        }).join('') +
        '</p:custShowLst>';

  const embeddedFontLst =
    fontRid === null
      ? ''
      : '<p:embeddedFontLst><p:embeddedFont>' +
        '<p:font typeface="ProbeAlpha" panose="02000000000000000000" pitchFamily="2" charset="0"/>' +
        `<p:regular r:id="${fontRid}"/>` +
        '</p:embeddedFont></p:embeddedFontLst>';

  // CT_Presentation is an xsd:sequence. `embeddedFontLst` goes after `notesSz`
  // and before `custShowLst`; `extLst` is always last.
  zip.add(
    'ppt/presentation.xml',
    DECLARATION +
      `<p:presentation ${NS_DECLS}` +
      // Both attributes or neither. `@embedTrueTypeFonts` is what PowerPoint
      // actually reads: omit it and a perfect `p:embeddedFontLst` is ignored
      // entirely, which is a bug with no symptom until a clean machine opens
      // the file. `@saveSubsetFonts` governs subsetting and never compression -
      // sub-phase 0.7 measured that.
      (fontRid === null ? '' : ' embedTrueTypeFonts="1" saveSubsetFonts="1"') +
      '>' +
      `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>` +
      (notesMasterRid === null
        ? ''
        : `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>`) +
      '<p:sldIdLst>' +
      slideRids.map((id, i) => `<p:sldId id="${String(256 + i)}" r:id="${id}"/>`).join('') +
      '</p:sldIdLst>' +
      `<p:sldSz cx="${String(recipe.slideWidth)}" cy="${String(recipe.slideHeight)}"` +
      (recipe.slideWidth === 12192000 && recipe.slideHeight === 6858000
        ? ' type="screen16x9"'
        : '') +
      '/>' +
      `<p:notesSz cx="${String(NOTES_WIDTH)}" cy="${String(NOTES_HEIGHT)}"/>` +
      embeddedFontLst +
      customShowsXml +
      sectionsXml +
      '</p:presentation>',
  );
  zip.add('ppt/_rels/presentation.xml.rels', relsXml(presRels));

  zip.add('ppt/presProps.xml', DECLARATION + `<p:presentationPr ${NS_DECLS}/>`);
  zip.add('ppt/viewProps.xml', DECLARATION + `<p:viewPr ${NS_DECLS}/>`);
  zip.add(
    'ppt/tableStyles.xml',
    DECLARATION + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
  );
  zip.add('ppt/theme/theme1.xml', themeXml('PPTX Studio Bench'));

  // --- master, layouts, notes master ---------------------------------------
  zip.add('ppt/slideMasters/slideMaster1.xml', slideMasterXml(layoutCount));
  zip.add(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    relsXml([
      ...Array.from({ length: layoutCount }, (_, i) => ({
        id: 'rId' + String(i + 1),
        type: REL + 'slideLayout',
        target: `../slideLayouts/slideLayout${String(i + 1)}.xml`,
      })),
      { id: 'rId' + String(layoutCount + 1), type: REL + 'theme', target: '../theme/theme1.xml' },
    ]),
  );
  for (let i = 0; i < layoutCount; i++) {
    zip.add(`ppt/slideLayouts/slideLayout${String(i + 1)}.xml`, slideLayoutXml(i));
    zip.add(
      `ppt/slideLayouts/_rels/slideLayout${String(i + 1)}.xml.rels`,
      relsXml([
        { id: 'rId1', type: REL + 'slideMaster', target: '../slideMasters/slideMaster1.xml' },
      ]),
    );
  }
  if (recipe.notes) {
    zip.add('ppt/theme/theme2.xml', themeXml('PPTX Studio Bench Notes'));
    zip.add('ppt/notesMasters/notesMaster1.xml', notesMasterXml());
    zip.add(
      'ppt/notesMasters/_rels/notesMaster1.xml.rels',
      relsXml([{ id: 'rId1', type: REL + 'theme', target: '../theme/theme2.xml' }]),
    );
  }

  if (fontBytes !== null) {
    zip.add('ppt/fonts/font1.fntdata', new Uint8Array(fontBytes), { store: true });
  }

  if (commentSlides.length > 0) {
    zip.add(
      'ppt/commentAuthors.xml',
      DECLARATION +
        `<p:cmAuthorLst ${NS_DECLS}>` +
        '<p:cmAuthor id="1" name="Bench" initials="B" lastIdx="1" clrIdx="0"/>' +
        '</p:cmAuthorLst>',
    );
  }

  // --- slides ---------------------------------------------------------------
  let chartNumber = 0;
  let diagramNumber = 0;
  let commentNumber = 0;

  for (const index of slideIndices) {
    const number = index + 1;
    const random = mulberry32(0x51de + index * 7919);
    const layoutIndex = index === 0 ? 0 : 1 + (index % (layoutCount - 1));
    const rels: Rel[] = [
      {
        id: 'rId1',
        type: REL + 'slideLayout',
        target: `../slideLayouts/slideLayout${String(layoutIndex + 1)}.xml`,
      },
    ];
    let slideRid = 1;
    const nextSlideRid = (): string => 'rId' + String(++slideRid);

    let body = '';
    let nextId = 2;

    body += placeholder(
      nextId++,
      'Title ' + String(number),
      layoutIndex === 0 ? 'ctrTitle' : 'title',
      null,
      838200,
      365125,
      10515600,
      1325563,
      paragraph('Slide ' + String(number) + ' — ' + words(random, 3)),
    );

    if (fontRid !== null) {
      body += placeholder(
        nextId++,
        'Embedded font sample ' + String(number),
        'body',
        2,
        838200,
        6100000,
        4000000,
        365125,
        '<a:p><a:r><a:rPr lang="en-GB" sz="1400">' +
          '<a:latin typeface="ProbeAlpha"/></a:rPr>' +
          '<a:t>ABCDEFGH</a:t></a:r></a:p>',
      );
    }

    const bullets = Array.from({ length: recipe.bulletsPerSlide }, (_, b) =>
      paragraph(words(random, 6 + (b % 5)), b % 3),
    ).join('');
    body += placeholder(
      nextId++,
      'Content ' + String(number),
      layoutIndex === 0 ? 'subTitle' : 'body',
      1,
      838200,
      1825625,
      5181600,
      4351338,
      bullets,
    );

    if (imageSize > 0) {
      const imageRid = nextSlideRid();
      rels.push({
        id: imageRid,
        type: REL + 'image',
        target: `../media/image${String(number)}.png`,
      });
      body += pictureXml(nextId++, imageRid, 6172200, 1825625, counters);
      counters.pictures += 1;
    }

    let hyperlinkRid: string | null = null;
    if (occurs(index, recipe.hyperlinks)) {
      hyperlinkRid = nextSlideRid();
      rels.push({
        id: hyperlinkRid,
        type: REL + 'hyperlink',
        target: 'https://github.com/Dewiride-Open-Source/Dewiride-PPTX-Studio',
        external: true,
      });
      counters.hyperlinks += 1;
    }

    // The only thing that carries the link is the first free shape, so a recipe
    // with no free shapes used to write the relationship and then attach it to
    // nothing: an orphan rel, and a census reporting no hyperlink against a
    // recipe that asked for one. `danglingRelationships()` cannot catch it
    // either, because an External relationship has no part to be missing.
    // Rather than drop the feature, give it somewhere to live.

    body += placeholder(
      nextId++,
      'Slide Number Placeholder ' + String(number),
      'sldNum',
      12,
      8724900,
      6356350,
      2743200,
      365125,
      '<a:p><a:pPr algn="r"/>' +
        `<a:fld id="${guid(0x5117 + index)}" type="slidenum">` +
        '<a:rPr lang="en-GB" smtClean="0"/>' +
        `<a:t>${String(number)}</a:t></a:fld>` +
        '<a:endParaRPr lang="en-GB"/></a:p>',
    );
    counters.fields += 1;

    const firstFreeShapeId = nextId;
    for (let s = 0; s < recipe.shapesPerSlide; s++) {
      body += freeShapeXml(
        nextId++,
        index * 17 + s,
        random,
        s === 0 ? hyperlinkRid : null,
        counters,
        recipe.slideWidth,
        recipe.slideHeight,
      );
    }
    if (hyperlinkRid !== null && recipe.shapesPerSlide === 0) {
      body += linkShapeXml(nextId++, hyperlinkRid, counters);
    }

    if (occurs(index, recipe.alternateContent)) {
      body += alternateContentXml(nextId++, counters);
      counters.alternates += 1;
    }
    if (occurs(index, recipe.tables)) {
      body += tableXml(nextId++, 4, 4, random);
      counters.tables += 1;
      counters.graphicFrames += 1;
    }
    if (occurs(index, recipe.charts)) {
      chartNumber += 1;
      const chartRid = nextSlideRid();
      rels.push({
        id: chartRid,
        type: REL + 'chart',
        target: `../charts/chart${String(chartNumber)}.xml`,
      });
      body += chartFrameXml(nextId++, chartRid);
      counters.charts += 1;
      counters.graphicFrames += 1;

      zip.add(`ppt/charts/chart${String(chartNumber)}.xml`, chartXml(0xc4a7 + chartNumber));
      zip.add(
        `ppt/charts/_rels/chart${String(chartNumber)}.xml.rels`,
        relsXml([
          {
            id: 'rId1',
            type: REL + 'package',
            target: `../embeddings/workbook${String(chartNumber)}.xlsx`,
          },
        ]),
      );
      // A real archive inside the archive. Nothing we write ever opens it -
      // charts render from the caches inside `chart1.xml` - but a `package`
      // relationship pointing at a file that is not a package is exactly the
      // sort of thing that only fails when a user double-clicks the chart.
      zip.add(
        `ppt/embeddings/workbook${String(chartNumber)}.xlsx`,
        writeInnerZip(workbookParts()),
        { store: true },
      );
      counters.embeddedPackages += 1;
    }
    if (occurs(index, recipe.diagrams)) {
      diagramNumber += 1;
      const k = String(diagramNumber);
      const ids: [string, string, string, string] = [
        nextSlideRid(),
        nextSlideRid(),
        nextSlideRid(),
        nextSlideRid(),
      ];
      rels.push(
        { id: ids[0], type: REL + 'diagramData', target: `../diagrams/data${k}.xml` },
        { id: ids[1], type: REL + 'diagramLayout', target: `../diagrams/layout${k}.xml` },
        { id: ids[2], type: REL + 'diagramQuickStyle', target: `../diagrams/quickStyle${k}.xml` },
        { id: ids[3], type: REL + 'diagramColors', target: `../diagrams/colors${k}.xml` },
      );
      const drawingRid = nextSlideRid();
      rels.push({
        id: drawingRid,
        type: REL_MS07 + 'diagramDrawing',
        target: `../diagrams/drawing${k}.xml`,
      });
      body += diagramFrameXml(nextId++, ids);
      counters.diagrams += 1;
      counters.graphicFrames += 1;

      const parts = diagramParts(0xd1a6 + diagramNumber * 101, counters);
      // `dsp:dataModelExt/@relId` resolves against the **slide's** relationships,
      // not the data part's, even though the element lives inside data1.xml.
      zip.add(
        `ppt/diagrams/data${k}.xml`,
        parts.data.replace(
          '<dgm:cxnLst/>',
          '<dgm:cxnLst/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
            `<dsp:dataModelExt xmlns:dsp="${NS_DSP}" relId="${drawingRid}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
            '</a:ext></dgm:extLst>',
        ),
      );
      zip.add(`ppt/diagrams/layout${k}.xml`, parts.layout);
      zip.add(`ppt/diagrams/quickStyle${k}.xml`, parts.quickStyle);
      zip.add(`ppt/diagrams/colors${k}.xml`, parts.colors);
      zip.add(`ppt/diagrams/drawing${k}.xml`, parts.drawing);
    }
    if (occurs(index, recipe.comments)) {
      commentNumber += 1;
      const commentRid = nextSlideRid();
      rels.push({
        id: commentRid,
        type: REL + 'comments',
        target: `../comments/comment${String(commentNumber)}.xml`,
      });
      zip.add(
        `ppt/comments/comment${String(commentNumber)}.xml`,
        DECLARATION +
          `<p:cmLst ${NS_DECLS}>` +
          `<p:cm authorId="1" dt="2026-08-27T09:00:00" idx="1"><p:pos x="100" y="100"/>` +
          `<p:text>${escapeXml('Synthetic comment on slide ' + String(number))}</p:text></p:cm>` +
          '</p:cmLst>',
      );
      counters.comments += 1;
    }

    if (recipe.notes) {
      rels.push({
        id: nextSlideRid(),
        type: REL + 'notesSlide',
        target: `../notesSlides/notesSlide${String(number)}.xml`,
      });
    }

    // Ids start at 2: `p:cNvPr id="1"` is the group shape every `p:spTree`
    // opens with. Read here as well as written, which is what keeps the last
    // block above from quietly ceasing to allocate one - `p:cNvPr` has to be
    // unique within a part, and that must not depend on which optional feature
    // happens to come last.
    counters.slideShapes += nextId - 2;

    const timing = occurs(index, recipe.animations) ? timingXml(firstFreeShapeId) : '';
    if (timing !== '') counters.animations += 1;

    zip.add(
      `ppt/slides/slide${String(number)}.xml`,
      DECLARATION +
        `<p:sld ${NS_DECLS}>` +
        '<p:cSld>' +
        spTreeHead() +
        body +
        '</p:spTree></p:cSld>' +
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
        timing +
        '</p:sld>',
    );
    zip.add(`ppt/slides/_rels/slide${String(number)}.xml.rels`, relsXml(rels));

    if (imageSize > 0) {
      // Stored, not deflated. The bytes are noise; DEFLATE would spend real
      // time to make them very slightly larger.
      zip.add(`ppt/media/image${String(number)}.png`, noisePng(imageSize, 0x9e37 + index), {
        store: true,
      });
    }

    if (recipe.notes) {
      zip.add(
        `ppt/notesSlides/notesSlide${String(number)}.xml`,
        DECLARATION +
          `<p:notes ${NS_DECLS}>` +
          '<p:cSld>' +
          spTreeHead() +
          placeholder(
            2,
            'Notes Placeholder ' + String(number),
            'body',
            1,
            685800,
            4400550,
            5486400,
            4114800,
            paragraph('Speaker notes for slide ' + String(number) + ': ' + words(random, 12)),
          ) +
          '</p:spTree></p:cSld>' +
          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
          '</p:notes>',
      );
      zip.add(
        `ppt/notesSlides/_rels/notesSlide${String(number)}.xml.rels`,
        relsXml([
          { id: 'rId1', type: REL + 'notesMaster', target: '../notesMasters/notesMaster1.xml' },
          { id: 'rId2', type: REL + 'slide', target: `../slides/slide${String(number)}.xml` },
        ]),
      );
      counters.notesSlides += 1;
    }
  }

  // --- docProps -------------------------------------------------------------
  zip.add(
    'docProps/core.xml',
    DECLARATION +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
      ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"' +
      ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${escapeXml(recipe.title)}</dc:title>` +
      '<dc:creator>tools/bench/make-deck.ts</dc:creator>' +
      '<cp:lastModifiedBy>tools/bench/make-deck.ts</cp:lastModifiedBy>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-08-27T00:00:00Z</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-08-27T00:00:00Z</dcterms:modified>' +
      '</cp:coreProperties>',
  );
  zip.add(
    'docProps/app.xml',
    DECLARATION +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"' +
      ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Application>PPTX Studio bench</Application>' +
      `<Slides>${String(recipe.slides)}</Slides>` +
      '<ScaleCrop>false</ScaleCrop>' +
      '<Company>Dewiride Technologies</Company>' +
      '<AppVersion>0.0000</AppVersion>' +
      '</Properties>',
  );

  const result = zip.close();
  return {
    recipe,
    bytes: result.bytes,
    entries: result.entries,
    imagePixels: imageSize,
    slideShapes: counters.slideShapes,
    expected: expectedFeatures(counters),
  };
}

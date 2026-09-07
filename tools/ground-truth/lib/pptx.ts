/**
 * Build a minimal PresentationML package by hand.
 *
 * Experiments B and C both need a deck that says exactly what we tell it to
 * say. Authoring one through PowerPoint's object model is no good for either:
 * COM cannot express `lumMod` on a fill, and it certainly cannot embed a font
 * we built ourselves. So we write the XML.
 *
 * That makes this the first `.pptx` the project has ever *written*, which is
 * worth stating plainly: if PowerPoint opens these without a repair prompt,
 * that is early evidence for the writer in Phase 1.3, and if it does not, we
 * have found out in sub-phase 0.7 rather than in Phase 1.
 *
 * The theme here is ours: the Office 2013 colour values, which are the ones
 * every real deck's `accent1` is near, but authored fresh rather than copied
 * out of a Microsoft template. Nothing Microsoft-authored is committed.
 */

import { writeZip, type ZipEntry } from './zip.ts';

/** Content types for the media a probe deck may carry. */
const MEDIA_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
};

export const EMU_PER_INCH = 914400;
export const SLIDE_WIDTH = 12192000; // 13.333in - 16:9
export const SLIDE_HEIGHT = 6858000; // 7.5in

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const NS_DECLS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

/** The twelve `a:clrScheme` children, in the only order the schema allows. */
export const CLR_SCHEME = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
} as const;

function theme(fonts: FontScheme): string {
  // dk1/lt1 are written as sysClr with @lastClr, exactly as PowerPoint writes
  // them, because that is the shape a renderer has to cope with: the resolver
  // prefers @lastClr and only falls back to the system colour name.
  const scheme = [
    `<a:dk1><a:sysClr val="windowText" lastClr="${CLR_SCHEME.dk1}"/></a:dk1>`,
    `<a:lt1><a:sysClr val="window" lastClr="${CLR_SCHEME.lt1}"/></a:lt1>`,
    `<a:dk2><a:srgbClr val="${CLR_SCHEME.dk2}"/></a:dk2>`,
    `<a:lt2><a:srgbClr val="${CLR_SCHEME.lt2}"/></a:lt2>`,
    ...(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const).map(
      (k) => `<a:${k}><a:srgbClr val="${CLR_SCHEME[k]}"/></a:${k}>`,
    ),
    `<a:hlink><a:srgbClr val="${CLR_SCHEME.hlink}"/></a:hlink>`,
    `<a:folHlink><a:srgbClr val="${CLR_SCHEME.folHlink}"/></a:folHlink>`,
  ].join('');

  const font = (kind: 'major' | 'minor', latin: string, ea: string, cs: string): string =>
    `<a:${kind}Font><a:latin typeface="${latin}"/><a:ea typeface="${ea}"/><a:cs typeface="${cs}"/></a:${kind}Font>`;

  // Exactly three entries in each list. PowerPoint is not forgiving about this:
  // `bgRef/@idx` and `fillRef/@idx` are 1-based indices into these lists and a
  // short list is an out-of-range reference on the very first slide.
  const fill = (i: number): string =>
    i === 0
      ? '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
      : `<a:gradFill rotWithShape="1"><a:gsLst>` +
        `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="${String(60000 - i * 10000)}"/><a:satMod val="120000"/></a:schemeClr></a:gs>` +
        `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="${String(100000 - i * 20000)}"/></a:schemeClr></a:gs>` +
        `</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;

  const line = (w: number): string =>
    `<a:ln w="${String(w)}" cap="flat" cmpd="sng" algn="ctr">` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;

  return (
    DECLARATION +
    `<a:theme xmlns:a="${NS_A}" name="PPTX Studio Ground Truth">` +
    '<a:themeElements>' +
    `<a:clrScheme name="Ground Truth">${scheme}</a:clrScheme>` +
    `<a:fontScheme name="Ground Truth">` +
    font('major', fonts.majorLatin, fonts.majorEa, fonts.majorCs) +
    font('minor', fonts.minorLatin, fonts.minorEa, fonts.minorCs) +
    '</a:fontScheme>' +
    '<a:fmtScheme name="Ground Truth">' +
    `<a:fillStyleLst>${fill(0)}${fill(1)}${fill(2)}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>' +
    '</a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    '</a:theme>'
  );
}

/** An empty `p:spTree` prologue. `nvGrpSpPr` then `grpSpPr`, always, in that order. */
function spTreeHead(): string {
  return (
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  );
}

/** The twelve `p:clrMap` attributes, in schema order. */
const CLR_MAP_KEYS = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

export type ClrMapAttrs = Readonly<Record<(typeof CLR_MAP_KEYS)[number], string>>;

/** What every master this project writes has said until now. */
export const IDENTITY_CLR_MAP: ClrMapAttrs = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

function slideMaster(map: ClrMapAttrs): string {
  // All twelve clrMap attributes. Eleven is a repair prompt.
  const clrMap = '<p:clrMap ' + CLR_MAP_KEYS.map((k) => k + '="' + map[k] + '"').join(' ') + '/>';
  return (
    DECLARATION +
    `<p:sldMaster ${NS_DECLS}>` +
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    spTreeHead() +
    '</p:spTree></p:cSld>' +
    clrMap +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '</p:sldMaster>'
  );
}

function slideLayout(): string {
  return (
    DECLARATION +
    `<p:sldLayout ${NS_DECLS} type="blank" preserve="1">` +
    '<p:cSld name="Blank">' +
    spTreeHead() +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sldLayout>'
  );
}

/** `p:bg` comes before `p:spTree`: `CT_CommonSlideData` is an `xsd:sequence`. */
function slide(body: string, background: string | undefined): string {
  // `a:effectLst` is what PowerPoint writes in every `p:bgPr` it authors itself.
  const bg =
    background === undefined ? '' : `<p:bg><p:bgPr>${background}<a:effectLst/></p:bgPr></p:bg>`;
  return (
    DECLARATION +
    `<p:sld ${NS_DECLS}>` +
    '<p:cSld>' +
    bg +
    spTreeHead() +
    body +
    '</p:spTree></p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sld>'
  );
}

function rels(entries: readonly { id: string; type: string; target: string }[]): string {
  return (
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    entries
      .map((e) => `<Relationship Id="${e.id}" Type="${e.type}" Target="${e.target}"/>`)
      .join('') +
    '</Relationships>'
  );
}

/** The theme's `a:fontScheme`. An empty name is what a stock Office theme writes. */
export interface FontScheme {
  readonly majorLatin: string;
  readonly majorEa: string;
  readonly majorCs: string;
  readonly minorLatin: string;
  readonly minorEa: string;
  readonly minorCs: string;
}

/** What every deck this project wrote before T7 used. */
export const DEFAULT_FONT_SCHEME: FontScheme = {
  majorLatin: 'Calibri Light',
  majorEa: '',
  majorCs: '',
  minorLatin: 'Calibri',
  minorEa: '',
  minorCs: '',
};

/** One embedded font: the four optional slots, each pointing at a `.fntdata` part. */
export interface EmbeddedFont {
  readonly typeface: string;
  /** 20 hex characters from `OS/2.panose`. */
  readonly panose: string;
  readonly pitchFamily: number;
  /** A **signed** byte. Shift-JIS is `-128`, not `128`. */
  readonly charset: number;
  /** `regular` | `bold` | `italic` | `boldItalic` -> the EOT bytes for that slot. */
  readonly slots: Readonly<
    Partial<Record<'regular' | 'bold' | 'italic' | 'boldItalic', Uint8Array>>
  >;
}

/** One file under `ppt/media/`, related from every slide. */
export interface MediaPart {
  /** File name with its extension, which is what the content type keys off. */
  readonly name: string;
  readonly bytes: Uint8Array;
}

export interface BuildOptions {
  /** One entry per slide: the children of `p:spTree` after `grpSpPr`. */
  readonly slides: readonly string[];
  /**
   * Images, related from every slide at `rId2` upwards in this order, so a probe
   * can name its `r:embed` without reading anything back.
   */
  readonly media?: readonly MediaPart[];
  readonly fonts?: readonly EmbeddedFont[];
  /**
   * The master colour map. Defaults to the identity, which is what every deck
   * this project has written so far used - and is exactly the configuration in
   * which a resolver that runs `dk1` through the map looks correct.
   */
  readonly clrMap?: ClrMapAttrs;
  /** Defaults to `DEFAULT_FONT_SCHEME`. */
  readonly fontScheme?: FontScheme;
  /**
   * One entry per slide: the fill element inside that slide's `p:bgPr`, or
   * `undefined` to inherit the master's background.
   */
  readonly backgrounds?: readonly (string | undefined)[];
}

export function buildPptx(options: BuildOptions): Uint8Array {
  const fonts = options.fonts ?? [];
  const media = options.media ?? [];

  // ---- presentation.xml.rels: master, then slides, then theme, then fonts ---
  const presRels: { id: string; type: string; target: string }[] = [];
  let rid = 0;
  const nextId = (): string => `rId${String(++rid)}`;

  const masterRid = nextId();
  presRels.push({
    id: masterRid,
    type: `${REL}/slideMaster`,
    target: 'slideMasters/slideMaster1.xml',
  });

  const slideRids = options.slides.map((_, i) => {
    const id = nextId();
    presRels.push({ id, type: `${REL}/slide`, target: `slides/slide${String(i + 1)}.xml` });
    return id;
  });

  const themeRid = nextId();
  presRels.push({ id: themeRid, type: `${REL}/theme`, target: 'theme/theme1.xml' });

  // ---- font parts -----------------------------------------------------------
  const fontParts: ZipEntry[] = [];
  let fontIndex = 0;
  const embeddedFontXml = fonts
    .map((font) => {
      const slots = (['regular', 'bold', 'italic', 'boldItalic'] as const)
        .map((slot) => {
          const bytes = font.slots[slot];
          if (bytes === undefined) return '';
          fontIndex += 1;
          const part = `ppt/fonts/font${String(fontIndex)}.fntdata`;
          fontParts.push({ name: part, bytes });
          const id = nextId();
          presRels.push({
            id,
            type: `${REL}/font`,
            target: `fonts/font${String(fontIndex)}.fntdata`,
          });
          return `<p:${slot} r:id="${id}"/>`;
        })
        .join('');
      return (
        '<p:embeddedFont>' +
        `<p:font typeface="${font.typeface}" panose="${font.panose}"` +
        ` pitchFamily="${String(font.pitchFamily)}" charset="${String(font.charset)}"/>` +
        slots +
        '</p:embeddedFont>'
      );
    })
    .join('');

  // ---- presentation.xml -----------------------------------------------------
  // CT_Presentation's sequence: sldMasterIdLst, notesMasterIdLst,
  // handoutMasterIdLst, sldIdLst, sldSz, notesSz, smartTags, embeddedFontLst,
  // custShowLst, photoAlbum, custDataLst, kinsoku, defaultTextStyle, modifyVerifier,
  // extLst. `embeddedFontLst` goes after `notesSz`, not at the end.
  const presentation =
    DECLARATION +
    `<p:presentation ${NS_DECLS}` +
    (fonts.length > 0 ? ' embedTrueTypeFonts="1"' : '') +
    ' saveSubsetFonts="0">' +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>` +
    '<p:sldIdLst>' +
    slideRids.map((id, i) => `<p:sldId id="${String(256 + i)}" r:id="${id}"/>`).join('') +
    '</p:sldIdLst>' +
    `<p:sldSz cx="${String(SLIDE_WIDTH)}" cy="${String(SLIDE_HEIGHT)}"/>` +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    (embeddedFontXml === '' ? '' : `<p:embeddedFontLst>${embeddedFontXml}</p:embeddedFontLst>`) +
    '</p:presentation>';

  // ---- content types --------------------------------------------------------
  const CT = 'application/vnd.openxmlformats-officedocument';
  const overrides = [
    { part: '/ppt/presentation.xml', type: `${CT}.presentationml.presentation.main+xml` },
    { part: '/ppt/slideMasters/slideMaster1.xml', type: `${CT}.presentationml.slideMaster+xml` },
    { part: '/ppt/slideLayouts/slideLayout1.xml', type: `${CT}.presentationml.slideLayout+xml` },
    ...options.slides.map((_, i) => ({
      part: `/ppt/slides/slide${String(i + 1)}.xml`,
      type: `${CT}.presentationml.slide+xml`,
    })),
    { part: '/ppt/theme/theme1.xml', type: `${CT}.theme+xml` },
  ];

  /** One `Default` per distinct media extension; a missing one is a repair. */
  const mediaDefaults = [
    ...new Map(
      media.map((part) => {
        const extension = part.name.slice(part.name.lastIndexOf('.') + 1).toLowerCase();
        const type = MEDIA_TYPES[extension];
        if (type === undefined) throw new Error(`no content type for media "${part.name}"`);
        return [extension, type] as const;
      }),
    ),
  ];

  const contentTypes =
    DECLARATION +
    `<Types xmlns="${NS_CT}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    mediaDefaults
      .map(([extension, type]) => `<Default Extension="${extension}" ContentType="${type}"/>`)
      .join('') +
    // Omitting this one is the canonical "PowerPoint found a problem with
    // content" bug for embedded fonts.
    (fontParts.length > 0
      ? '<Default Extension="fntdata" ContentType="application/x-fontdata"/>'
      : '') +
    overrides.map((o) => `<Override PartName="${o.part}" ContentType="${o.type}"/>`).join('') +
    '</Types>';

  const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

  // `[Content_Types].xml` first, then `_rels/.rels`, then everything else.
  const entries: ZipEntry[] = [
    { name: '[Content_Types].xml', bytes: utf8(contentTypes) },
    {
      name: '_rels/.rels',
      bytes: utf8(
        rels([{ id: 'rId1', type: `${REL}/officeDocument`, target: 'ppt/presentation.xml' }]),
      ),
    },
    { name: 'ppt/presentation.xml', bytes: utf8(presentation) },
    { name: 'ppt/_rels/presentation.xml.rels', bytes: utf8(rels(presRels)) },
    { name: 'ppt/theme/theme1.xml', bytes: utf8(theme(options.fontScheme ?? DEFAULT_FONT_SCHEME)) },
    {
      name: 'ppt/slideMasters/slideMaster1.xml',
      bytes: utf8(slideMaster(options.clrMap ?? IDENTITY_CLR_MAP)),
    },
    {
      name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      bytes: utf8(
        rels([
          { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
          { id: 'rId2', type: `${REL}/theme`, target: '../theme/theme1.xml' },
        ]),
      ),
    },
    { name: 'ppt/slideLayouts/slideLayout1.xml', bytes: utf8(slideLayout()) },
    {
      name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      bytes: utf8(
        rels([
          { id: 'rId1', type: `${REL}/slideMaster`, target: '../slideMasters/slideMaster1.xml' },
        ]),
      ),
    },
  ];

  const mediaRels = media.map((part, i) => ({
    id: `rId${String(i + 2)}`,
    type: `${REL}/image`,
    target: `../media/${part.name}`,
  }));

  options.slides.forEach((body, i) => {
    const n = String(i + 1);
    entries.push({
      name: `ppt/slides/slide${n}.xml`,
      bytes: utf8(slide(body, options.backgrounds?.[i])),
    });
    entries.push({
      name: `ppt/slides/_rels/slide${n}.xml.rels`,
      bytes: utf8(
        rels([
          { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
          ...mediaRels,
        ]),
      ),
    });
  });

  for (const part of media) entries.push({ name: `ppt/media/${part.name}`, bytes: part.bytes });
  entries.push(...fontParts);
  return writeZip(entries);
}

/**
 * Escape text for an XML text node.
 *
 * Not optional, and not obvious: PowerPoint's response to a stray `<` in an
 * `a:t` is to refuse the whole package with "The file or directory is corrupted
 * and unreadable" - no part name, no line number, no clue that the problem is
 * one character in one text run.
 */
export function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A solid-filled rectangle with no outline, positioned in EMU. */
export function rect(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  fill: string,
): string {
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvSpPr/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    `<a:solidFill>${fill}</a:solidFill>` +
    '<a:ln><a:noFill/></a:ln>' +
    '</p:spPr></p:sp>'
  );
}

/** A text box. `runs` is already-built `a:r` markup. */
export function textBox(
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  runs: string,
): string {
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm><a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
    '</p:spPr><p:txBody>' +
    '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>' +
    `<a:p>${runs}</a:p>` +
    '</p:txBody></p:sp>'
  );
}

/** How a probe shape is positioned, shaped and filled. */
export interface ShapeOptions {
  readonly id: number;
  readonly name: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** A whole fill element: `<a:gradFill>...</a:gradFill>`, `<a:pattFill .../>`, `<a:noFill/>`. */
  readonly fill: string;
  /** `ST_ShapeType`. Defaults to `rect`. Ignored when `geom` is given. */
  readonly prst?: string | undefined;
  /**
   * A whole geometry element - `<a:prstGeom .../>` or `<a:custGeom>...</a:custGeom>`.
   *
   * C4 needs `prst="line"` and a hand-written `custGeom` with a corner of a
   * chosen angle, neither of which `prst` alone can express.
   */
  readonly geom?: string | undefined;
  /** Sixtieths of a degree, on `a:xfrm/@rot`. */
  readonly rot?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  /** A whole `a:ln` element. Defaults to no outline, which keeps edges clean for sampling. */
  readonly line?: string | undefined;
  /**
   * A whole `a:effectLst`. Goes after `a:ln` - `CT_ShapeProperties` is a
   * sequence and there is only one place it opens from.
   */
  readonly effect?: string | undefined;
}

/**
 * A shape with an arbitrary fill.
 *
 * `rect()` above takes a *colour* and wraps it in `a:solidFill`, which is all
 * experiments C and C2 needed. C3 fills shapes with gradients and patterns, so
 * it hands over the fill element whole.
 */
export function shapeXml(options: ShapeOptions): string {
  const { id, name, x, y, cx, cy, fill } = options;
  const flip =
    (options.flipH === true ? ' flipH="1"' : '') + (options.flipV === true ? ' flipV="1"' : '');
  const rot = options.rot === undefined || options.rot === 0 ? '' : ` rot="${String(options.rot)}"`;
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(id)}" name="${name}"/><p:cNvSpPr/><p:nvPr/>` +
    '</p:nvSpPr><p:spPr>' +
    `<a:xfrm${rot}${flip}>` +
    `<a:off x="${String(x)}" y="${String(y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/>` +
    '</a:xfrm>' +
    (options.geom ?? `<a:prstGeom prst="${options.prst ?? 'rect'}"><a:avLst/></a:prstGeom>`) +
    fill +
    (options.line ?? '<a:ln><a:noFill/></a:ln>') +
    (options.effect ?? '') +
    '</p:spPr></p:sp>'
  );
}

/** A `p:grpSp`. The child coordinate space is the whole subject of C6. */
export interface GroupOptions {
  readonly id: number;
  readonly name: string;
  /** `a:off` and `a:ext`, in EMU: the rectangle the group occupies on the slide. */
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /**
   * `a:chOff` and `a:chExt`, in EMU: the coordinate system the children are
   * written in. A group PowerPoint has just made sets these equal to `off` and
   * `ext`, so the children keep their slide coordinates verbatim; every
   * subsequent resize changes `ext` alone and leaves both the children and
   * `chExt` untouched, which is how the two come apart.
   */
  readonly chX: number;
  readonly chY: number;
  readonly chCx: number;
  readonly chCy: number;
  readonly rot?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  /** A whole fill element on `p:grpSpPr` - what `a:grpFill` children reach for. */
  readonly fill?: string | undefined;
  /**
   * Write no `a:chOff` and no `a:chExt` at all.
   *
   * Both are optional in `CT_GroupTransform2D`, and what a group without them
   * does to its children is a question only a hand-written package can ask -
   * PowerPoint always writes all four.
   */
  readonly noChild?: boolean | undefined;
  /** Already-built `p:sp` / `p:grpSp` markup. */
  readonly children: string;
}

export function groupXml(o: GroupOptions): string {
  const flip = (o.flipH === true ? ' flipH="1"' : '') + (o.flipV === true ? ' flipV="1"' : '');
  const rot = o.rot === undefined || o.rot === 0 ? '' : ` rot="${String(o.rot)}"`;
  const child =
    o.noChild === true
      ? ''
      : `<a:chOff x="${String(o.chX)}" y="${String(o.chY)}"/><a:chExt cx="${String(o.chCx)}" cy="${String(o.chCy)}"/>`;
  return (
    '<p:grpSp><p:nvGrpSpPr>' +
    `<p:cNvPr id="${String(o.id)}" name="${o.name}"/><p:cNvGrpSpPr/><p:nvPr/>` +
    '</p:nvGrpSpPr><p:grpSpPr>' +
    `<a:xfrm${rot}${flip}>` +
    `<a:off x="${String(o.x)}" y="${String(o.y)}"/><a:ext cx="${String(o.cx)}" cy="${String(o.cy)}"/>` +
    child +
    '</a:xfrm>' +
    (o.fill ?? '') +
    '</p:grpSpPr>' +
    o.children +
    '</p:grpSp>'
  );
}

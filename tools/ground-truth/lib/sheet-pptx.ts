/**
 * Build a PresentationML package with more than one master, more than one
 * layout, and a slide bound to whichever of them the probe wants.
 *
 * `pptx.ts` writes one master, one layout and N slides, which was everything
 * C1 through C4 needed: those experiments asked what a *fill* paints, and a fill
 * has no ancestors. C5 asks what a shape inherits, and inheritance cannot be
 * probed inside a package that has only one of each.
 *
 * So this is a second builder rather than a generalisation of the first. The
 * shapes differ enough - masters own themes, layouts own masters, slides own
 * layouts, and every one of those bindings is a relationship that a hostile
 * probe wants to break on purpose - that folding them together would put four
 * working experiments at risk to save one file.
 *
 * ## Every binding is a relationship, and every relationship is overridable
 *
 * The one architectural fact this sub-phase rests on is that a slide does not
 * name its layout: `ppt/slides/_rels/slideN.xml.rels` does. So the builder
 * takes the rels of every part as data, with a default that is correct and an
 * override that need not be. A probe that wants a slide bound to nothing, or to
 * two layouts, or to a layout its master has never heard of, says so here.
 */

import { writeZip, type ZipEntry } from './zip.ts';

export const EMU_PER_POINT = 12700;
/** 960 x 540 points, which is the 13.333in x 7.5in of a 16:9 deck. */
export const SLIDE_WIDTH_PT = 960;
export const SLIDE_HEIGHT_PT = 540;

const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = NS_R;

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_DECLS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

/** Points to EMU. Throws on a fraction, because a fractional probe coordinate
 *  is a rounding error waiting to be read back as a match failure. */
export function emu(points: number, what: string): number {
  const value = points * EMU_PER_POINT;
  if (!Number.isInteger(value)) {
    throw new Error(`${what}: ${String(points)}pt is not a whole number of EMU`);
  }
  return value;
}

/* -------------------------------------------------------------------------- */
/* the theme                                                                  */
/* -------------------------------------------------------------------------- */

/** The twelve `a:clrScheme` children, in the only order the schema allows. */
export interface Scheme {
  readonly dk1: string;
  readonly lt1: string;
  readonly dk2: string;
  readonly lt2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hlink: string;
  readonly folHlink: string;
}

/** The Office 2013 values, as C1 through C4 have used throughout. */
export const SCHEME_ONE: Scheme = {
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
};

/**
 * A second palette, sharing not one value with the first.
 *
 * Two masters whose `accent1` differ is the only way to ask whether a slide
 * resolves its theme through its own master or through the presentation's first
 * one - and the presentation part has a `theme` relationship of its own, which
 * is exactly the wrong thing to read.
 */
export const SCHEME_TWO: Scheme = {
  dk1: '111111',
  lt1: 'FDFDFD',
  dk2: '2E4A1E',
  lt2: 'DDE8D4',
  accent1: '1F9E8C',
  accent2: '8E44AD',
  accent3: 'D35400',
  accent4: '16A085',
  accent5: 'C0392B',
  accent6: '2C3E50',
  hlink: '0B7285',
  folHlink: '862E9C',
};

/**
 * The style matrix, built so that a resolved colour names the entry it came
 * from.
 *
 * Six fills, all of `phClr`, each with a different transform: the hue says
 * which colour was passed in and the luminance says which list entry applied
 * it. Three line widths, half a point apart in a ratio no rounding can blur.
 * With that, `Shape.Fill.ForeColor.RGB` and `Shape.Line.Weight` over COM answer
 * "which entry of which list" without a pixel being measured.
 */
export const FILL_STYLES: readonly string[] = [
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',
  '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="60000"/></a:schemeClr></a:solidFill>',
  '<a:solidFill><a:schemeClr val="phClr"><a:lumMod val="30000"/></a:schemeClr></a:solidFill>',
];

export const BG_FILL_STYLES: readonly string[] = [
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="90000"/></a:schemeClr></a:solidFill>',
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:solidFill>',
  '<a:solidFill><a:schemeClr val="phClr"><a:shade val="40000"/></a:schemeClr></a:solidFill>',
];

/** Half a point, two points, four and a half. Read back through `Line.Weight`. */
export const LINE_STYLE_WIDTHS: readonly number[] = [6350, 25400, 57150];

/**
 * A theme part: the colour scheme, and the two things 3.1 needs to vary.
 *
 * The font scheme matters because `+mj-lt`/`+mn-lt` is the last hop of the text
 * cascade, and `a:objectDefaults` because it is the second-to-last and no
 * implementation this project has read consults it at all.
 */
export interface ThemeSpec {
  readonly scheme: Scheme;
  /** `a:majorFont/a:latin/@typeface`. */
  readonly majorLatin?: string | undefined;
  readonly minorLatin?: string | undefined;
  /** A whole `a:objectDefaults`, written after `a:themeElements`. */
  readonly objectDefaults?: string | undefined;
}

function themeXml(spec: ThemeSpec, name: string): string {
  const scheme = spec.scheme;
  const slot = (k: keyof Scheme): string =>
    k === 'dk1'
      ? `<a:dk1><a:sysClr val="windowText" lastClr="${scheme.dk1}"/></a:dk1>`
      : k === 'lt1'
        ? `<a:lt1><a:sysClr val="window" lastClr="${scheme.lt1}"/></a:lt1>`
        : `<a:${k}><a:srgbClr val="${scheme[k]}"/></a:${k}>`;

  const clrScheme = (
    [
      'dk1',
      'lt1',
      'dk2',
      'lt2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ] as const
  )
    .map(slot)
    .join('');

  const font = (kind: 'major' | 'minor', latin: string): string =>
    `<a:${kind}Font><a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/></a:${kind}Font>`;

  const line = (w: number): string =>
    `<a:ln w="${String(w)}" cap="flat" cmpd="sng" algn="ctr">` +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';

  return (
    DECLARATION +
    `<a:theme xmlns:a="${NS_A}" name="${name}">` +
    '<a:themeElements>' +
    `<a:clrScheme name="${name}">${clrScheme}</a:clrScheme>` +
    `<a:fontScheme name="${name}">${font('major', spec.majorLatin ?? 'Calibri Light')}${font('minor', spec.minorLatin ?? 'Calibri')}</a:fontScheme>` +
    `<a:fmtScheme name="${name}">` +
    `<a:fillStyleLst>${FILL_STYLES.join('')}</a:fillStyleLst>` +
    `<a:lnStyleLst>${LINE_STYLE_WIDTHS.map(line).join('')}</a:lnStyleLst>` +
    '<a:effectStyleLst>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3) +
    '</a:effectStyleLst>' +
    `<a:bgFillStyleLst>${BG_FILL_STYLES.join('')}</a:bgFillStyleLst>` +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    (spec.objectDefaults ?? '') +
    '</a:theme>'
  );
}

/* -------------------------------------------------------------------------- */
/* shapes                                                                     */
/* -------------------------------------------------------------------------- */

/** A rectangle in points on the slide. */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface ShapeSpec {
  readonly id: number;
  readonly name: string;
  /** Absent means the shape states no geometry, which is the whole subject. */
  readonly rect?: Rect | undefined;
  /** `a:xfrm/@rot` in sixtieths of a degree. Needs `rect`. */
  readonly rot?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
  /** Raw `p:ph` attributes, e.g. `type="body" idx="1"`. `''` is a bare `<p:ph/>`. */
  readonly ph?: string | undefined;
  /** A whole fill element. Absent writes nothing, which is not the same as `<a:noFill/>`. */
  readonly fill?: string | undefined;
  /** A whole `a:ln`. Absent writes nothing. */
  readonly line?: string | undefined;
  /** A whole `p:style`. Absent writes nothing. */
  readonly style?: string | undefined;
  /** `ST_ShapeType`. Defaults to `rect`. */
  readonly prst?: string | undefined;
  /** A whole `a:bodyPr`. Defaults to an empty one. */
  readonly bodyPr?: string | undefined;
  /** A whole `a:lstStyle`. Defaults to an empty one, which declares nothing. */
  readonly lstStyle?: string | undefined;
  /** Whole `a:p` elements. Defaults to one empty paragraph. */
  readonly paragraphs?: readonly string[] | undefined;
}

export function shape(spec: ShapeSpec): string {
  const geom = `<a:prstGeom prst="${spec.prst ?? 'rect'}"><a:avLst/></a:prstGeom>`;
  const turn =
    (spec.rot === undefined || spec.rot === 0 ? '' : ` rot="${String(spec.rot)}"`) +
    (spec.flipH === true ? ' flipH="1"' : '') +
    (spec.flipV === true ? ' flipV="1"' : '');
  const xfrm =
    spec.rect === undefined
      ? ''
      : `<a:xfrm${turn}><a:off x="${String(emu(spec.rect.x, spec.name))}" y="${String(emu(spec.rect.y, spec.name))}"/>` +
        `<a:ext cx="${String(emu(spec.rect.w, spec.name))}" cy="${String(emu(spec.rect.h, spec.name))}"/></a:xfrm>`;
  const ph = spec.ph === undefined ? '' : spec.ph === '' ? '<p:ph/>' : `<p:ph ${spec.ph}/>`;
  return (
    '<p:sp><p:nvSpPr>' +
    `<p:cNvPr id="${String(spec.id)}" name="${spec.name}"/><p:cNvSpPr/>` +
    `<p:nvPr>${ph}</p:nvPr>` +
    '</p:nvSpPr><p:spPr>' +
    xfrm +
    geom +
    (spec.fill ?? '') +
    (spec.line ?? '') +
    '</p:spPr>' +
    (spec.style ?? '') +
    '<p:txBody>' +
    (spec.bodyPr ?? '<a:bodyPr/>') +
    (spec.lstStyle ?? '<a:lstStyle/>') +
    (spec.paragraphs ?? ['<a:p/>']).join('') +
    '</p:txBody>' +
    '</p:sp>'
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

/* -------------------------------------------------------------------------- */
/* the three sheets                                                           */
/* -------------------------------------------------------------------------- */

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

/** One relationship, for the probes that want to write their own. */
export interface RelSpec {
  readonly id: string;
  readonly type: string;
  readonly target: string;
}

export interface MasterSpec {
  /** Index into `SheetPackage.themes`. */
  readonly theme: number;
  readonly clrMap?: ClrMapAttrs | undefined;
  /** A whole `p:bg`. Defaults to `bgRef idx="1001"` with `bg1`, as PowerPoint writes. */
  readonly bg?: string | undefined;
  readonly shapes?: readonly string[] | undefined;
  /**
   * A whole `p:txStyles`, written after `p:sldLayoutIdLst` as `CT_SlideMaster`
   * sequences it. Absent means the master states no text styles at all, which
   * no PowerPoint-authored master does and which is exactly why it is worth
   * being able to write.
   */
  readonly txStyles?: string | undefined;
  /** Replaces the generated rels entirely. For hostile probes only. */
  readonly rels?: readonly RelSpec[] | undefined;
}

export interface LayoutSpec {
  /** Index into `SheetPackage.masters`. */
  readonly master: number;
  /** `ST_SlideLayoutType`. Defaults to `obj`. */
  readonly type?: string | undefined;
  readonly name?: string | undefined;
  readonly bg?: string | undefined;
  /** Defaults to `<a:masterClrMapping/>`. */
  readonly clrMapOvr?: string | undefined;
  readonly showMasterSp?: boolean | undefined;
  readonly shapes?: readonly string[] | undefined;
  readonly rels?: readonly RelSpec[] | undefined;
  /** Leave this layout out of its master's `sldLayoutIdLst`. */
  readonly unlisted?: boolean | undefined;
}

export interface SlideSpec {
  /** Index into `SheetPackage.layouts`. */
  readonly layout: number;
  readonly name?: string | undefined;
  readonly bg?: string | undefined;
  readonly clrMapOvr?: string | undefined;
  readonly showMasterSp?: boolean | undefined;
  readonly shapes?: readonly string[] | undefined;
  readonly rels?: readonly RelSpec[] | undefined;
}

export interface SheetPackage {
  readonly themes: readonly ThemeSpec[];
  readonly masters: readonly MasterSpec[];
  readonly layouts: readonly LayoutSpec[];
  readonly slides: readonly SlideSpec[];
  /**
   * A whole `p:defaultTextStyle` on `ppt/presentation.xml`, written after
   * `p:notesSz`. The seventh source of the text cascade, and the one the plan
   * claims a placeholder never reaches.
   */
  readonly defaultTextStyle?: string | undefined;
  /**
   * A whole `<p:kinsoku/>` element, written between `p:notesSz` and
   * `p:defaultTextStyle` - the position `CT_Presentation` gives it, after
   * `custDataLst` and before `defaultTextStyle`.
   *
   * Absent is a distinct case from present-and-empty: T3 needs a package that
   * states no kinsoku at all, to find out whether PowerPoint falls back to a
   * built-in list or does no filtering.
   */
  readonly kinsoku?: string | undefined;
  /**
   * `p:presentation/@strictFirstAndLastChars`. Omitted when undefined, so the
   * schema default applies and the deck can ask what that default is.
   */
  readonly strictFirstAndLastChars?: boolean | undefined;
  /**
   * Files to drop into `ppt/media/`, for the one bullet kind that is a picture.
   *
   * A part with no `Override` needs a `Default` for its extension, and omitting
   * that one is the canonical "PowerPoint found a problem" bug - so the
   * extension is derived from the name here rather than taken on trust, and a
   * name without one throws.
   */
  readonly media?: readonly { readonly name: string; readonly bytes: Uint8Array }[];
}

function relsXml(entries: readonly RelSpec[]): string {
  return (
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    entries
      .map((e) => `<Relationship Id="${e.id}" Type="${e.type}" Target="${e.target}"/>`)
      .join('') +
    '</Relationships>'
  );
}

export function buildSheetPackage(pkg: SheetPackage): Uint8Array {
  const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);
  const entries: ZipEntry[] = [];
  const overrides: { part: string; type: string }[] = [];
  const CT = 'application/vnd.openxmlformats-officedocument';

  // ---- themes -------------------------------------------------------------
  pkg.themes.forEach((theme, i) => {
    const n = String(i + 1);
    entries.push({
      name: `ppt/theme/theme${n}.xml`,
      bytes: utf8(themeXml(theme, `Ground Truth ${n}`)),
    });
    overrides.push({ part: `/ppt/theme/theme${n}.xml`, type: `${CT}.theme+xml` });
  });

  // ---- masters ------------------------------------------------------------
  // Which layouts each master owns, in declaration order. The `sldLayoutIdLst`
  // and the rels are built from the same list, and a probe may then break one
  // of them without touching the other.
  const layoutsOf = pkg.masters.map((_, m) =>
    pkg.layouts.map((l, i) => ({ l, i })).filter((e) => e.l.master === m),
  );

  // `sldMasterId/@id` and `sldLayoutId/@id` share one number space above
  // 2147483648, and PowerPoint allocates them from a single counter: a master,
  // then its own layouts, then the next master. Giving each master its own
  // sub-range looks tidier and collides on the second master, which PowerPoint
  // answers by repairing the file - a repaired deck being exactly the thing an
  // experiment cannot measure.
  let nextSheetId = 2147483648;
  const masterIds: number[] = [];
  const layoutIds = new Map<number, number>();
  pkg.masters.forEach((_, m) => {
    masterIds.push(nextSheetId++);
    for (const e of layoutsOf[m] ?? []) {
      if (e.l.unlisted !== true) layoutIds.set(e.i, nextSheetId++);
    }
  });

  pkg.masters.forEach((master, m) => {
    const n = String(m + 1);
    const owned = layoutsOf[m] ?? [];
    const listed = owned.filter((e) => e.l.unlisted !== true);
    const clrMap =
      '<p:clrMap ' +
      CLR_MAP_KEYS.map((k) => `${k}="${(master.clrMap ?? IDENTITY_CLR_MAP)[k]}"`).join(' ') +
      '/>';
    const bg = master.bg ?? '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
    const xml =
      DECLARATION +
      `<p:sldMaster ${NS_DECLS}>` +
      `<p:cSld>${bg}${spTreeHead()}${(master.shapes ?? []).join('')}</p:spTree></p:cSld>` +
      clrMap +
      '<p:sldLayoutIdLst>' +
      listed
        .map(
          (e, k) =>
            `<p:sldLayoutId id="${String(layoutIds.get(e.i) ?? 0)}" r:id="rId${String(k + 1)}"/>`,
        )
        .join('') +
      '</p:sldLayoutIdLst>' +
      (master.txStyles ?? '') +
      '</p:sldMaster>';
    entries.push({ name: `ppt/slideMasters/slideMaster${n}.xml`, bytes: utf8(xml) });
    overrides.push({
      part: `/ppt/slideMasters/slideMaster${n}.xml`,
      type: `${CT}.presentationml.slideMaster+xml`,
    });

    // The theme relationship lives here, on the master, and nowhere else that
    // matters. Reading it off `ppt/presentation.xml.rels` instead gives every
    // slide the first master's palette, which is right on every one-master deck
    // in the world and wrong on this one.
    const rels: RelSpec[] = listed.map((e, k) => ({
      id: `rId${String(k + 1)}`,
      type: `${REL}/slideLayout`,
      target: `../slideLayouts/slideLayout${String(e.i + 1)}.xml`,
    }));
    rels.push({
      id: `rId${String(listed.length + 1)}`,
      type: `${REL}/theme`,
      target: `../theme/theme${String(master.theme + 1)}.xml`,
    });
    entries.push({
      name: `ppt/slideMasters/_rels/slideMaster${n}.xml.rels`,
      bytes: utf8(relsXml(master.rels ?? rels)),
    });
  });

  // ---- layouts ------------------------------------------------------------
  pkg.layouts.forEach((layout, i) => {
    const n = String(i + 1);
    const attrs =
      ` type="${layout.type ?? 'obj'}" preserve="1"` +
      (layout.showMasterSp === false ? ' showMasterSp="0"' : '');
    const xml =
      DECLARATION +
      `<p:sldLayout ${NS_DECLS}${attrs}>` +
      `<p:cSld name="${layout.name ?? `Layout ${n}`}">` +
      (layout.bg ?? '') +
      spTreeHead() +
      (layout.shapes ?? []).join('') +
      '</p:spTree></p:cSld>' +
      `<p:clrMapOvr>${layout.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr>` +
      '</p:sldLayout>';
    entries.push({ name: `ppt/slideLayouts/slideLayout${n}.xml`, bytes: utf8(xml) });
    overrides.push({
      part: `/ppt/slideLayouts/slideLayout${n}.xml`,
      type: `${CT}.presentationml.slideLayout+xml`,
    });
    entries.push({
      name: `ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`,
      bytes: utf8(
        relsXml(
          layout.rels ?? [
            {
              id: 'rId1',
              type: `${REL}/slideMaster`,
              target: `../slideMasters/slideMaster${String(layout.master + 1)}.xml`,
            },
          ],
        ),
      ),
    });
  });

  // ---- slides -------------------------------------------------------------
  pkg.slides.forEach((slide, i) => {
    const n = String(i + 1);
    const attrs = slide.showMasterSp === false ? ' showMasterSp="0"' : '';
    const xml =
      DECLARATION +
      `<p:sld ${NS_DECLS}${attrs}>` +
      `<p:cSld name="${slide.name ?? `Slide ${n}`}">` +
      (slide.bg ?? '') +
      spTreeHead() +
      (slide.shapes ?? []).join('') +
      '</p:spTree></p:cSld>' +
      `<p:clrMapOvr>${slide.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr>` +
      '</p:sld>';
    entries.push({ name: `ppt/slides/slide${n}.xml`, bytes: utf8(xml) });
    overrides.push({
      part: `/ppt/slides/slide${n}.xml`,
      type: `${CT}.presentationml.slide+xml`,
    });
    entries.push({
      name: `ppt/slides/_rels/slide${n}.xml.rels`,
      bytes: utf8(
        relsXml(
          slide.rels ?? [
            {
              id: 'rId1',
              type: `${REL}/slideLayout`,
              target: `../slideLayouts/slideLayout${String(slide.layout + 1)}.xml`,
            },
          ],
        ),
      ),
    });
  });

  // ---- presentation -------------------------------------------------------
  const presRels: RelSpec[] = [];
  let rid = 0;
  const nextId = (): string => `rId${String(++rid)}`;

  const masterRids = pkg.masters.map((_, m) => {
    const id = nextId();
    presRels.push({
      id,
      type: `${REL}/slideMaster`,
      target: `slideMasters/slideMaster${String(m + 1)}.xml`,
    });
    return id;
  });
  const slideRids = pkg.slides.map((_, i) => {
    const id = nextId();
    presRels.push({ id, type: `${REL}/slide`, target: `slides/slide${String(i + 1)}.xml` });
    return id;
  });
  // PowerPoint writes this and it is a trap: it names theme1 whatever the
  // masters say. It is written here precisely so the resolver has the chance to
  // read the wrong one.
  presRels.push({ id: nextId(), type: `${REL}/theme`, target: 'theme/theme1.xml' });

  const presentation =
    DECLARATION +
    `<p:presentation ${NS_DECLS} saveSubsetFonts="0"` +
    (pkg.strictFirstAndLastChars === undefined
      ? ''
      : ` strictFirstAndLastChars="${pkg.strictFirstAndLastChars ? '1' : '0'}"`) +
    `>` +
    '<p:sldMasterIdLst>' +
    masterRids
      .map((id, m) => `<p:sldMasterId id="${String(masterIds[m] ?? 0)}" r:id="${id}"/>`)
      .join('') +
    '</p:sldMasterIdLst>' +
    '<p:sldIdLst>' +
    slideRids.map((id, i) => `<p:sldId id="${String(256 + i)}" r:id="${id}"/>`).join('') +
    '</p:sldIdLst>' +
    `<p:sldSz cx="${String(emu(SLIDE_WIDTH_PT, 'slide'))}" cy="${String(emu(SLIDE_HEIGHT_PT, 'slide'))}"/>` +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    (pkg.kinsoku ?? '') +
    (pkg.defaultTextStyle ?? '') +
    '</p:presentation>';

  entries.push({ name: 'ppt/presentation.xml', bytes: utf8(presentation) });
  entries.push({ name: 'ppt/_rels/presentation.xml.rels', bytes: utf8(relsXml(presRels)) });
  overrides.push({
    part: '/ppt/presentation.xml',
    type: `${CT}.presentationml.presentation.main+xml`,
  });

  // ---- media --------------------------------------------------------------
  const mediaTypes = new Map<string, string>();
  const MEDIA_CONTENT_TYPES: Readonly<Record<string, string>> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
  };
  for (const file of pkg.media ?? []) {
    const dot = file.name.lastIndexOf('.');
    if (dot <= 0) throw new Error(`media file ${file.name} has no extension to declare`);
    const ext = file.name.slice(dot + 1).toLowerCase();
    const type = MEDIA_CONTENT_TYPES[ext];
    if (type === undefined) throw new Error(`no content type known for media extension .${ext}`);
    mediaTypes.set(ext, type);
    entries.push({ name: `ppt/media/${file.name}`, bytes: file.bytes });
  }

  const contentTypes =
    DECLARATION +
    `<Types xmlns="${NS_CT}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    [...mediaTypes]
      .map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`)
      .join('') +
    overrides.map((o) => `<Override PartName="${o.part}" ContentType="${o.type}"/>`).join('') +
    '</Types>';

  return writeZip([
    { name: '[Content_Types].xml', bytes: utf8(contentTypes) },
    {
      name: '_rels/.rels',
      bytes: utf8(
        relsXml([{ id: 'rId1', type: `${REL}/officeDocument`, target: 'ppt/presentation.xml' }]),
      ),
    },
    ...entries,
  ]);
}

/**
 * Experiment C9 - the probes, and every reading of the table-style cascade they are scored against:
 * which parts reach a cell or a grid edge, which of them wins, and where the style's text sits in the
 * text cascade. PowerPoint draws only its 74 built-ins (ADR 0063), so they are the probes; ADR 0064.
 */

import {
  BUILTIN_TABLE_STYLES,
  builtinTableStyle,
  parseTheme,
  styleMatrixFill,
  styleMatrixLine,
  type TableStyle,
  type TableStylePart,
  type TableStylePartName,
  type Theme,
  type Themeable,
} from '../../../../../packages/model/dist/index.js';
import {
  resolveColor,
  toHexColor,
  type ClrMap,
  type Color,
  type ColorContext,
  type Fill as PaintFill,
  type Line as PaintLine,
  type Rgba,
} from '../../../../../packages/paint/dist/index.js';
import { parseXmlString } from '../../../../../packages/xml/dist/index.js';
import {
  IDENTITY_CLR_MAP,
  SCHEME_TWO,
  shape,
  themeXml,
  type ClrMapAttrs,
  type FormatSchemeSpec,
  type ThemeSpec,
} from '../../../lib/sheet-pptx.ts';
import { TABLE_URI } from '../grid/probes.ts';
import { SCHEME_THREE } from '../styles/probes.ts';

/* -------------------------------------------------------------------------- */
/* geometry and themes                                                        */
/* -------------------------------------------------------------------------- */

export const EMU_PER_POINT = 12700;
/** Every table's top-left corner, and its cells, in points. */
export const ORIGIN = { x: 72, y: 72 } as const;
export const CELL = { w: 72, h: 36 } as const;
/** Four pixels per point: every built-in line is a whole number of half points, so two pixels or more. */
export const SCALE = 4;
export const EXPORT = { width: 960 * SCALE, height: 540 * SCALE } as const;
/** How far either side of a grid line an edge profile reaches, in pixels: the widest line is 4.5 pt. */
export const PROFILE_REACH = 16;

export type ThemeKey = 'A' | 'B' | 'C';

export interface ThemeCase {
  readonly theme: ThemeSpec;
  readonly clrMap: ClrMapAttrs;
}

const CLR_MAP_A: ClrMapAttrs = {
  ...IDENTITY_CLR_MAP,
  bg1: 'lt2',
  tx1: 'dk2',
  bg2: 'lt1',
  tx2: 'dk1',
};

const phClr = (transform = ''): string => `<a:schemeClr val="phClr">${transform}</a:schemeClr>`;
const gradient = (from: string, to: string, angle: number): string =>
  `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${from}</a:gs><a:gs pos="100000">${to}</a:gs></a:gsLst>` +
  `<a:lin ang="${String(angle * 60000)}" scaled="0"/></a:gradFill>`;
const shadow = (hex: string, direction: number): string =>
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="0" dist="76200" dir="${String(direction * 60000)}" algn="tl" rotWithShape="0">` +
  `<a:srgbClr val="${hex}"/></a:outerShdw></a:effectLst></a:effectStyle>`;

/** Theme C's style matrix: gradients and hard shadows where stock themes have them, and a dashed line. */
export const FORMAT_C: FormatSchemeSpec = {
  fills: [
    `<a:solidFill>${phClr('<a:lumMod val="80000"/>')}</a:solidFill>`,
    gradient(phClr('<a:lumMod val="20000"/>'), phClr(), 0),
    gradient(phClr(), phClr('<a:lumMod val="50000"/>'), 90),
  ],
  lines: [6350, 25400, 57150].map(
    (w, k) =>
      `<a:ln w="${String(w)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${phClr()}</a:solidFill>` +
      `<a:prstDash val="${k === 1 ? 'dash' : 'solid'}"/></a:ln>`,
  ),
  effects: [
    shadow('C9C0E1', 0),
    '<a:effectStyle><a:effectLst/></a:effectStyle>',
    shadow('C9C0E3', 90),
  ],
  bgFills: [
    `<a:solidFill>${phClr('<a:tint val="90000"/>')}</a:solidFill>`,
    `<a:solidFill>${phClr('<a:tint val="40000"/>')}</a:solidFill>`,
    `<a:solidFill>${phClr('<a:shade val="40000"/>')}</a:solidFill>`,
  ],
};

/**
 * A: twelve distinct slots, `tx1` apart from `dk1` and `bg1` from `lt1`. B: another palette on the
 * identity map. C: A with gradient fills and shadows. None writes an `a:sysClr` (ADR 0021, decision 9).
 */
export const THEMES: Readonly<Record<ThemeKey, ThemeCase>> = {
  A: {
    theme: {
      scheme: SCHEME_THREE,
      rgbDarkLight: true,
      majorLatin: 'Georgia',
      minorLatin: 'Verdana',
    },
    clrMap: CLR_MAP_A,
  },
  B: { theme: { scheme: SCHEME_TWO, rgbDarkLight: true }, clrMap: IDENTITY_CLR_MAP },
  C: {
    theme: {
      scheme: SCHEME_THREE,
      rgbDarkLight: true,
      majorLatin: 'Georgia',
      minorLatin: 'Verdana',
      formatScheme: FORMAT_C,
    },
    clrMap: CLR_MAP_A,
  },
};

/** A white page under every table, whatever `bg1` maps to. */
export const MASTER_BG =
  '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
export const PAGE = 'FFFFFF';

export interface Palette {
  readonly theme: Theme;
  readonly colors: ColorContext;
}

const PALETTES = new Map<ThemeKey, Palette>();

/** The theme a deck was built with, parsed from the bytes `buildSheetPackage` writes. */
export function paletteOf(key: ThemeKey): Palette {
  const known = PALETTES.get(key);
  if (known !== undefined) return known;
  const { theme, clrMap } = THEMES[key];
  const xml = themeXml(theme, 'Ground Truth 1');
  if (xml.includes('a:sysClr')) throw new Error(`theme ${key} writes an a:sysClr`);
  const parsed = parseTheme(parseXmlString(xml).root, '/ppt/theme/theme1.xml');
  const palette = { theme: parsed, colors: { scheme: parsed.scheme, map: clrMap as ClrMap } };
  PALETTES.set(key, palette);
  return palette;
}

/* -------------------------------------------------------------------------- */
/* tables                                                                     */
/* -------------------------------------------------------------------------- */

export const FLAGS = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'] as const;
export type Flag = (typeof FLAGS)[number];
export type Flags = Readonly<Record<Flag, boolean>>;

/** Bit k of `mask` is `FLAGS[k]`. */
export function flagsOf(mask: number): Flags {
  return Object.fromEntries(FLAGS.map((f, k) => [f, (mask & (1 << k)) !== 0])) as Record<
    Flag,
    boolean
  >;
}

export function maskOf(flags: Flags): number {
  return FLAGS.reduce((mask, f, k) => (flags[f] ? mask | (1 << k) : mask), 0);
}

export const ALL = 63;
export const HEADER_BANDED = 0b010001;

export interface CellSpec {
  /** Attributes of `a:tc`, verbatim, e.g. ` rowSpan="2"`. */
  readonly attrs?: string;
  /** The whole `a:tcPr`. */
  readonly tcPr?: string;
  /** The whole `a:txBody`. */
  readonly body?: string;
}

export interface TableSpec {
  /** What the slide asks, for the probes that ask one thing each. */
  readonly name?: string;
  readonly rows: number;
  readonly cols: number;
  /** A GUID, or `null` for an `a:tblPr` naming none. */
  readonly style: string | null;
  readonly flags: Flags;
  readonly rtl?: boolean;
  /** An `EG_FillProperties` element inside `a:tblPr`. */
  readonly tblPrFill?: string;
  /** An `a:effectLst` inside `a:tblPr`. */
  readonly tblPrEffects?: string;
  /** Raw `p:ph` attributes for the frame, e.g. `idx="1"`. */
  readonly ph?: string;
  /** The frame's `a:ext` as a multiple of the grid's size. */
  readonly extScale?: number;
  /** Keyed `"r,c"`, zero-based. */
  readonly cells?: Readonly<Record<string, CellSpec>>;
  /** Rows written with fewer `a:tc` than the grid has columns, by row: PowerPoint pads them (ADR 0056). */
  readonly short?: Readonly<Record<string, number>>;
}

/** How many `a:tc` row `r` of `spec` writes. */
export const cellsInRow = (spec: TableSpec, r: number): number =>
  spec.short?.[String(r)] ?? spec.cols;

export const TEXT_BODY =
  '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>X</a:t></a:r></a:p></a:txBody>';

export function tblPrXml(spec: TableSpec): string {
  const attrs =
    (spec.rtl === true ? ' rtl="1"' : '') +
    FLAGS.filter((f) => spec.flags[f])
      .map((f) => ` ${f}="1"`)
      .join('');
  const id = spec.style === null ? '' : `<a:tableStyleId>${spec.style}</a:tableStyleId>`;
  const inner = (spec.tblPrFill ?? '') + (spec.tblPrEffects ?? '') + id;
  return inner === '' ? `<a:tblPr${attrs}/>` : `<a:tblPr${attrs}>${inner}</a:tblPr>`;
}

const emu = (pt: number): string => String(Math.round(pt * EMU_PER_POINT));

export function tableFrame(name: string, spec: TableSpec): string {
  const rows = Array.from({ length: spec.rows }, (_, r) => {
    const cells = Array.from({ length: cellsInRow(spec, r) }, (_, c) => {
      const cell = spec.cells?.[`${String(r)},${String(c)}`] ?? {};
      return `<a:tc${cell.attrs ?? ''}>${cell.body ?? TEXT_BODY}${cell.tcPr ?? '<a:tcPr/>'}</a:tc>`;
    });
    return `<a:tr h="${emu(CELL.h)}">${cells.join('')}</a:tr>`;
  });
  const scale = spec.extScale ?? 1;
  const nvPr = spec.ph === undefined ? '<p:nvPr/>' : `<p:nvPr><p:ph ${spec.ph}/></p:nvPr>`;
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="2" name="${name}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>${nvPr}` +
    '</p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${emu(ORIGIN.x)}" y="${emu(ORIGIN.y)}"/>` +
    `<a:ext cx="${emu(spec.cols * CELL.w * scale)}" cy="${emu(spec.rows * CELL.h * scale)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${TABLE_URI}"><a:tbl>${tblPrXml(spec)}` +
    `<a:tblGrid>${`<a:gridCol w="${emu(CELL.w)}"/>`.repeat(spec.cols)}</a:tblGrid>${rows.join('')}</a:tbl>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

/* -------------------------------------------------------------------------- */
/* the text levels                                                            */
/* -------------------------------------------------------------------------- */

/** Colours, sizes and faces each text level declares: none a style or a theme has. */
export const LEVELS = {
  run: { hex: 'C9E001', sz: 1100, face: 'Courier New' },
  para: { hex: 'C9E002', sz: 1300, face: 'Consolas' },
  list: { hex: 'C9E003', sz: 1500, face: 'Tahoma' },
  layoutPh: { hex: 'C9E006', sz: 1600, face: 'Arial' },
  masterPh: { hex: 'C9E007', sz: 1800, face: 'Segoe UI' },
  bodyStyle: { hex: 'C9E008', sz: 2000, face: 'Palatino Linotype' },
  defaultTextStyle: { hex: 'C9E004', sz: 1700, face: 'Times New Roman' },
  otherStyle: { hex: 'C9E005', sz: 1900, face: 'Trebuchet MS' },
} as const;
export type TextLevel = keyof typeof LEVELS;

const solid = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;

/** A `defRPr` for `level`: colour, size, face, and `b="0"` so a style's bold shows by losing. */
function defRPr(level: TextLevel): string {
  const { hex, sz, face } = LEVELS[level];
  return `<a:defRPr sz="${String(sz)}" b="0">${solid(hex)}<a:latin typeface="${face}"/></a:defRPr>`;
}

/** A list style whose `lvl`-th level (1-based) declares `level`, under the element name `tag`. */
export function listStyle(level: TextLevel, tag = 'a:lstStyle', lvl = 1, bullet = false): string {
  const bu = bullet ? '<a:buFont typeface="Arial"/><a:buChar char="•"/>' : '';
  return `<${tag}><a:lvl${String(lvl)}pPr>${bu}${defRPr(level)}</a:lvl${String(lvl)}pPr></${tag}>`;
}

export type CellLevel = 'run' | 'para' | 'list';

/** One rung of the ladder: the cell-level levels it declares, and the paragraph's `@lvl`. */
export interface Rung {
  readonly levels: readonly CellLevel[];
  readonly lvl: number;
}

/** Every cell-level combination the ladder descends through, and a paragraph one level down with and without a list. */
export const LADDER: readonly Rung[] = [
  { levels: ['run', 'para', 'list'], lvl: 0 },
  { levels: ['para', 'list'], lvl: 0 },
  { levels: ['list'], lvl: 0 },
  { levels: [], lvl: 0 },
  { levels: ['list'], lvl: 1 },
  { levels: [], lvl: 1 },
];

export function ladderBody(rung: Rung): string {
  const list = rung.levels.includes('list')
    ? listStyle('list', 'a:lstStyle', rung.lvl + 1)
    : '<a:lstStyle/>';
  const lvl = rung.lvl === 0 ? '' : ` lvl="${String(rung.lvl)}"`;
  const pPr = rung.levels.includes('para')
    ? `<a:pPr${lvl}>${defRPr('para')}</a:pPr>`
    : lvl === ''
      ? ''
      : `<a:pPr${lvl}/>`;
  const { hex, sz, face } = LEVELS.run;
  const rPr = rung.levels.includes('run')
    ? `<a:rPr lang="en-US" sz="${String(sz)}" b="0">${solid(hex)}<a:latin typeface="${face}"/></a:rPr>`
    : '<a:rPr lang="en-US"/>';
  return `<a:txBody><a:bodyPr/>${list}<a:p>${pPr}<a:r>${rPr}<a:t>X</a:t></a:r></a:p></a:txBody>`;
}

/** The floor every sweep deck writes: a face and a colour no style part states, at both termini. */
export const FLOOR = { face: 'Courier New', slot: 'hlink' } as const;
const floorRPr = `<a:defRPr><a:solidFill><a:schemeClr val="${FLOOR.slot}"/></a:solidFill><a:latin typeface="${FLOOR.face}"/></a:defRPr>`;
export const FLOOR_DEFAULT_TEXT_STYLE = `<p:defaultTextStyle><a:lvl1pPr>${floorRPr}</a:lvl1pPr></p:defaultTextStyle>`;
export const FLOOR_TX_STYLES = `<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle><a:lvl1pPr>${floorRPr}</a:lvl1pPr></p:otherStyle></p:txStyles>`;

/* -------------------------------------------------------------------------- */
/* the decks                                                                  */
/* -------------------------------------------------------------------------- */

export type Group =
  | 'control'
  | 'sweep'
  | 'ends'
  | 'sides'
  | 'direct'
  | 'merge'
  | 'rtl'
  | 'tblpr'
  | 'background'
  | 'dash'
  | 'text';

/** `sweep` reads a cell's fill and text; `full` adds every side, the text's size and bullet, and the rect. */
export type ReadMode = 'sweep' | 'full';

export interface DeckSpec {
  readonly id: string;
  readonly group: Group;
  readonly theme: ThemeKey;
  readonly read: ReadMode;
  /** A whole `p:defaultTextStyle`. */
  readonly defaultTextStyle?: string;
  /** A whole `p:txStyles` for the master. */
  readonly txStyles?: string;
  /** Whole shapes on the master and on the layout, for the placeholder ladder. */
  readonly masterShapes?: readonly string[];
  readonly layoutShapes?: readonly string[];
  /** PowerPoint also saves a copy, so the analysis can read what it keeps. */
  readonly resave?: boolean;
  readonly slides: readonly TableSpec[];
}

export interface BuiltinRef {
  readonly id: string;
  readonly name: string;
}

/** Each built-in's XML as PowerPoint writes it, by GUID. */
export const BUILTIN_TABLE_STYLE_XML: ReadonlyMap<string, string> = new Map(
  BUILTIN_TABLE_STYLES.map((s) => [s.id, s.xml]),
);

export const BUILTINS: readonly BuiltinRef[] = BUILTIN_TABLE_STYLES.map((s) => ({
  id: s.id,
  name: s.name,
}));

export function builtinNamed(name: string): BuiltinRef {
  const found = BUILTINS.find((s) => s.name === name);
  if (found === undefined) throw new Error(`no built-in table style named "${name}"`);
  return found;
}

/** A GUID no built-in has: PowerPoint draws its default grid (ADR 0063). */
export const UNKNOWN_STYLE = '{C9000000-0000-4000-8000-000000000001}';

export const REFERENCE = {
  medium1: 'Medium Style 1',
  medium2: 'Medium Style 2 - Accent 1',
  medium3: 'Medium Style 3 - Accent 2',
  dark1: 'Dark Style 1',
  themed1: 'Themed Style 1 - Accent 1',
  themed2: 'Themed Style 2 - Accent 1',
  light2: 'Light Style 2',
  light2a1: 'Light Style 2 - Accent 1',
  noGrid: 'No Style, No Grid',
  tableGrid: 'No Style, Table Grid',
} as const;
const idOf = (key: keyof typeof REFERENCE): string => builtinNamed(REFERENCE[key]).id;

const two = (n: number): string => String(n).padStart(2, '0');
const ALL_MASKS = Array.from({ length: 64 }, (_, m) => m);
/** The sixteen combinations of the four end flags, bands on. */
export const END_MASKS: readonly number[] = Array.from({ length: 16 }, (_, m) => m | 0b110000);
/** The table shapes the 5x5 sweep cannot show: one row or column, and corners that touch. */
export const END_SHAPES: readonly (readonly [number, number])[] = [
  [1, 1],
  [2, 2],
  [1, 3],
  [3, 1],
];
/** Eight combinations that between them turn every part on somewhere. */
export const SIDE_MASKS: readonly number[] = [
  0,
  ALL,
  HEADER_BANDED,
  0b001111,
  0b110000,
  0b010100,
  0b100010,
  0b000101,
];

function slidesFor(
  style: string | null,
  shapes: readonly (readonly [number, number])[],
  masks: readonly number[],
): TableSpec[] {
  return shapes.flatMap(([rows, cols]) =>
    masks.map((mask) => ({ rows, cols, style, flags: flagsOf(mask) })),
  );
}

/** Direct `a:tcPr` colours no style or theme produces. */
export const DIRECT = {
  fill: 'C9D101',
  fillHeader: 'C9D102',
  right: 'C9D201',
  conflictA: 'C9D301',
  conflictB: 'C9D302',
  diagonal: 'C9D601',
  outer: 'C9D701',
  header: 'C9D801',
  below: 'C9D901',
  noWidth: 'C9DA01',
  lower: 'C9DB01',
  covered: 'C9A001',
  coveredFill: 'C9A002',
  segment: 'C9A003',
  segmentCovered: 'C9A004',
  bottomAnchor: 'C9A005',
  authoredAnchor: 'C9A006',
  authoredBelow: 'C9A007',
  rightAnchor: 'C9A008',
  rightBeside: 'C9A009',
  rightBesideCovered: 'C9A00A',
  wideRight: 'C9A00B',
  wideBottom: 'C9A00C',
  belowCovered: 'C9A00D',
  topWide: 'C9A00E',
  leftTall: 'C9A00F',
  wideBelow: 'C9A010',
  besideMergedLeft: 'C9A011',
  besideMergedRight: 'C9A012',
  bottomEdgeWide: 'C9A013',
  bottomEdgeCovered: 'C9A014',
  rightEdgeTall: 'C9A015',
  rightEdgeCovered: 'C9A016',
  wideBottomAlike: 'C9A017',
  dashed: 'C9A018',
  padded: 'C9A019',
  paddedBeside: 'C9A01A',
  diagonalAnchor: 'C9A01B',
  diagonalCovered: 'C9A01C',
  rtlLeft: 'C9B001',
  rtlRight: 'C9B002',
  rtlDown: 'C9B003',
  rtlUp: 'C9B004',
} as const;

export const line = (tag: string, pt: number, hex: string): string =>
  `<a:${tag} w="${emu(pt)}" cap="flat" cmpd="sng" algn="ctr">${solid(hex)}<a:prstDash val="solid"/><a:round/></a:${tag}>`;

export interface DirectSlide {
  readonly name: string;
  /** Replaces the deck's flags for this slide. */
  readonly mask?: number;
  readonly cells: Record<string, CellSpec>;
}

/** One direct-formatting question per slide, on the cells named. */
export const DIRECT_SLIDES: readonly DirectSlide[] = [
  {
    name: 'fill',
    cells: {
      '1,1': { tcPr: `<a:tcPr>${solid(DIRECT.fill)}</a:tcPr>` },
      '2,2': { tcPr: '<a:tcPr><a:noFill/></a:tcPr>' },
      '0,0': { tcPr: `<a:tcPr>${solid(DIRECT.fillHeader)}</a:tcPr>` },
    },
  },
  { name: 'right', cells: { '2,2': { tcPr: `<a:tcPr>${line('lnR', 3, DIRECT.right)}</a:tcPr>` } } },
  {
    name: 'bottom-none',
    cells: { '2,2': { tcPr: '<a:tcPr><a:lnB><a:noFill/></a:lnB></a:tcPr>' } },
  },
  {
    name: 'conflict-heavy-first',
    cells: {
      '2,2': { tcPr: `<a:tcPr>${line('lnR', 3, DIRECT.conflictA)}</a:tcPr>` },
      '2,3': { tcPr: `<a:tcPr>${line('lnL', 1, DIRECT.conflictB)}</a:tcPr>` },
    },
  },
  {
    name: 'conflict-heavy-second',
    cells: {
      '2,2': { tcPr: `<a:tcPr>${line('lnR', 1, DIRECT.conflictA)}</a:tcPr>` },
      '2,3': { tcPr: `<a:tcPr>${line('lnL', 3, DIRECT.conflictB)}</a:tcPr>` },
    },
  },
  {
    name: 'conflict-vertical',
    cells: {
      '2,2': { tcPr: `<a:tcPr>${line('lnB', 1, DIRECT.conflictA)}</a:tcPr>` },
      '3,2': { tcPr: `<a:tcPr>${line('lnT', 1, DIRECT.conflictB)}</a:tcPr>` },
    },
  },
  { name: 'width-only', cells: { '2,2': { tcPr: '<a:tcPr><a:lnB w="50800"/></a:tcPr>' } } },
  {
    name: 'fill-no-width',
    cells: { '2,2': { tcPr: `<a:tcPr><a:lnB>${solid(DIRECT.noWidth)}</a:lnB></a:tcPr>` } },
  },
  {
    name: 'dash-only',
    cells: { '2,2': { tcPr: '<a:tcPr><a:lnB><a:prstDash val="dash"/></a:lnB></a:tcPr>' } },
  },
  { name: 'compound-only', cells: { '2,2': { tcPr: '<a:tcPr><a:lnB cmpd="dbl"/></a:tcPr>' } } },
  { name: 'empty-line', cells: { '2,2': { tcPr: '<a:tcPr><a:lnB/></a:tcPr>' } } },
  {
    name: 'width-over-double',
    mask: ALL,
    cells: { '3,2': { tcPr: '<a:tcPr><a:lnB w="25400"/></a:tcPr>' } },
  },
  {
    name: 'lower-only',
    cells: { '2,2': { tcPr: `<a:tcPr>${line('lnT', 1, DIRECT.lower)}</a:tcPr>` } },
  },
  {
    name: 'right-only',
    cells: { '2,3': { tcPr: `<a:tcPr>${line('lnL', 1, DIRECT.lower)}</a:tcPr>` } },
  },
  {
    name: 'diagonals',
    cells: {
      '2,2': { tcPr: `<a:tcPr>${line('lnTlToBr', 2, DIRECT.diagonal)}</a:tcPr>` },
      '2,3': { tcPr: `<a:tcPr>${line('lnBlToTr', 2, DIRECT.diagonal)}</a:tcPr>` },
    },
  },
  {
    name: 'outer',
    cells: {
      '0,0': {
        tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.outer)}${line('lnT', 2, DIRECT.outer)}</a:tcPr>`,
      },
    },
  },
  {
    name: 'header-bottom',
    cells: { '0,1': { tcPr: `<a:tcPr>${line('lnB', 1, DIRECT.header)}</a:tcPr>` } },
  },
  {
    name: 'below-header',
    cells: { '1,1': { tcPr: `<a:tcPr>${line('lnT', 1, DIRECT.below)}</a:tcPr>` } },
  },
];

const covered = (h: boolean, v: boolean, tcPr?: string): CellSpec => ({
  attrs: (h ? ' hMerge="1"' : '') + (v ? ' vMerge="1"' : ''),
  ...(tcPr === undefined ? {} : { tcPr }),
});

export interface MergeSlide {
  readonly name: string;
  readonly mask: number;
  readonly cells: Record<string, CellSpec>;
}

/** Spans that cross the parts' regions, in PowerPoint's own form (ADR 0056), and direct lines on them. */
export const MERGE_SLIDES: readonly MergeSlide[] = [
  {
    name: 'header-into-band',
    mask: 0b110101,
    cells: { '0,1': { attrs: ' rowSpan="2"' }, '1,1': covered(false, true) },
  },
  {
    name: 'firstcol-into-band',
    mask: 0b110101,
    cells: { '1,0': { attrs: ' gridSpan="2"' }, '1,1': covered(true, false) },
  },
  {
    name: 'block-into-corner',
    mask: ALL,
    cells: {
      '3,3': { attrs: ' rowSpan="2" gridSpan="2"' },
      '3,4': { attrs: ' rowSpan="2" hMerge="1"' },
      '4,3': { attrs: ' gridSpan="2" vMerge="1"' },
      '4,4': covered(true, true),
    },
  },
  {
    name: 'band-pair',
    mask: 0b110001,
    cells: { '1,1': { attrs: ' rowSpan="2"' }, '2,1': covered(false, true) },
  },
  {
    name: 'whole-row',
    mask: 0b110011,
    cells: {
      '2,0': { attrs: ' gridSpan="5"' },
      '2,1': covered(true, false),
      '2,2': covered(true, false),
      '2,3': covered(true, false),
      '2,4': covered(true, false),
    },
  },
  {
    name: 'first-to-last-row',
    mask: ALL,
    cells: {
      '0,0': { attrs: ' rowSpan="5"' },
      '1,0': covered(false, true),
      '2,0': covered(false, true),
      '3,0': covered(false, true),
      '4,0': covered(false, true),
    },
  },
  {
    name: 'covered-direct',
    mask: HEADER_BANDED,
    cells: {
      '1,1': { attrs: ' rowSpan="2"' },
      '2,1': covered(
        false,
        true,
        `<a:tcPr>${line('lnB', 1, DIRECT.covered)}${solid(DIRECT.coveredFill)}</a:tcPr>`,
      ),
    },
  },
  {
    name: 'segment-anchor',
    mask: HEADER_BANDED,
    cells: {
      '1,1': { attrs: ' rowSpan="2"', tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.segment)}</a:tcPr>` },
      '2,1': covered(false, true),
    },
  },
  {
    name: 'segment-covered',
    mask: HEADER_BANDED,
    cells: {
      '1,1': { attrs: ' rowSpan="2"' },
      '2,1': covered(false, true, `<a:tcPr>${line('lnR', 2, DIRECT.segmentCovered)}</a:tcPr>`),
    },
  },
  {
    name: 'bottom-anchor',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.bottomAnchor)}</a:tcPr>`,
      },
      '2,1': covered(false, true),
    },
  },
  {
    // The shape PowerPoint writes for a merged cell's bottom border (author-merged-bottom).
    name: 'bottom-authored',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.authoredAnchor)}</a:tcPr>`,
      },
      '2,1': covered(false, true),
      '3,1': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.authoredBelow)}</a:tcPr>` },
    },
  },
  {
    name: 'below-covered',
    mask: HEADER_BANDED,
    cells: {
      '1,1': { attrs: ' rowSpan="2"' },
      '2,1': covered(false, true),
      '3,1': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.belowCovered)}</a:tcPr>` },
    },
  },
  {
    // The shape PowerPoint writes for a merged cell's right border (author-merged-right).
    name: 'right-authored',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.rightAnchor)}</a:tcPr>`,
      },
      '1,2': { tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.rightBeside)}</a:tcPr>` },
      '2,1': covered(false, true),
      '2,2': { tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.rightBesideCovered)}</a:tcPr>` },
    },
  },
  {
    name: 'wide-right',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' gridSpan="2"',
        tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.wideRight)}</a:tcPr>`,
      },
      '1,2': covered(true, false),
    },
  },
  {
    name: 'wide-bottom',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' gridSpan="2"',
        tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.wideBottom)}</a:tcPr>`,
      },
      '1,2': covered(true, false),
    },
  },
  {
    name: 'top-wide',
    mask: HEADER_BANDED,
    cells: {
      '0,1': { attrs: ' gridSpan="2"', tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.topWide)}</a:tcPr>` },
      '0,2': covered(true, false),
    },
  },
  {
    name: 'left-tall',
    mask: HEADER_BANDED,
    cells: {
      '1,0': { attrs: ' rowSpan="2"', tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.leftTall)}</a:tcPr>` },
      '2,0': covered(false, true),
    },
  },
  {
    // A wide cell's lnT under two owners, one writing it alike and one writing nothing.
    name: 'wide-below',
    mask: HEADER_BANDED,
    cells: {
      '1,1': { tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.wideBelow)}</a:tcPr>` },
      '2,1': {
        attrs: ' gridSpan="2"',
        tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.wideBelow)}</a:tcPr>`,
      },
      '2,2': covered(true, false),
    },
  },
  {
    // Two tall cells side by side: the right one is the cell after on an interior edge.
    name: 'beside-merged',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.besideMergedLeft)}</a:tcPr>`,
      },
      '2,1': covered(false, true),
      '1,2': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.besideMergedRight)}</a:tcPr>`,
      },
      '2,2': covered(false, true),
    },
  },
  {
    // The table's own bottom edge under a wide cell, where no cell after exists.
    name: 'bottom-edge-wide',
    mask: HEADER_BANDED,
    cells: {
      '4,1': {
        attrs: ' gridSpan="2"',
        tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.bottomEdgeWide)}</a:tcPr>`,
      },
      '4,2': covered(true, false, `<a:tcPr>${line('lnB', 2, DIRECT.bottomEdgeCovered)}</a:tcPr>`),
    },
  },
  {
    // The table's own right edge beside a tall cell.
    name: 'right-edge-tall',
    mask: HEADER_BANDED,
    cells: {
      '2,4': {
        attrs: ' rowSpan="2"',
        tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.rightEdgeTall)}</a:tcPr>`,
      },
      '3,4': covered(false, true, `<a:tcPr>${line('lnR', 2, DIRECT.rightEdgeCovered)}</a:tcPr>`),
    },
  },
  // Merges whose anchor's band differs from a covered position's: pooled or the anchor's.
  {
    name: 'band-rows-odd',
    mask: HEADER_BANDED,
    cells: { '2,1': { attrs: ' rowSpan="2"' }, '3,1': covered(false, true) },
  },
  {
    name: 'band-rows-three',
    mask: HEADER_BANDED,
    cells: {
      '1,2': { attrs: ' rowSpan="3"' },
      '2,2': covered(false, true),
      '3,2': covered(false, true),
    },
  },
  {
    name: 'band-cols-odd',
    mask: 0b100001,
    cells: { '2,1': { attrs: ' gridSpan="2"' }, '2,2': covered(true, false) },
  },
  {
    name: 'band-block',
    mask: 0b110000,
    cells: {
      '1,1': { attrs: ' rowSpan="2" gridSpan="2"' },
      '1,2': { attrs: ' rowSpan="2" hMerge="1"' },
      '2,1': { attrs: ' gridSpan="2" vMerge="1"' },
      '2,2': covered(true, true),
    },
  },
  {
    // A wide cell's bottom as PowerPoint's writer would put it: the anchor's lnB, an lnT below each column.
    name: 'wide-bottom-alike',
    mask: HEADER_BANDED,
    cells: {
      '1,1': {
        attrs: ' gridSpan="2"',
        tcPr: `<a:tcPr>${line('lnB', 2, DIRECT.wideBottomAlike)}</a:tcPr>`,
      },
      '1,2': covered(true, false),
      '2,1': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.wideBottomAlike)}</a:tcPr>` },
      '2,2': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.wideBottomAlike)}</a:tcPr>` },
    },
  },
];

function ladderTable(style: string | null, ph?: string): TableSpec {
  const cells: Record<string, CellSpec> = {};
  for (let r = 0; r < 2; r++) {
    LADDER.forEach((rung, c) => {
      cells[`${String(r)},${String(c)}`] = { body: ladderBody(rung) };
    });
  }
  return {
    rows: 2,
    cols: LADDER.length,
    style,
    flags: flagsOf(0b000001),
    cells,
    ...(ph === undefined ? {} : { ph }),
  };
}

export interface TextPackage {
  readonly id: string;
  readonly levels: readonly TextLevel[];
  /** The frame's `p:ph` attributes; absent is a frame that is no placeholder. */
  readonly ph?: string;
  /** The layout's and master's placeholder for the frame, `p:ph` attributes, with no list style. */
  readonly bare?: { readonly layout: string; readonly master: string };
}

/** Which of the package-level levels each text package declares. */
export const TEXT_PACKAGES: readonly TextPackage[] = [
  { id: 'text-both', levels: ['defaultTextStyle', 'otherStyle'] },
  { id: 'text-default', levels: ['defaultTextStyle'] },
  { id: 'text-other', levels: ['otherStyle'] },
  { id: 'text-none', levels: [] },
  {
    id: 'text-ph',
    ph: 'idx="1"',
    levels: ['layoutPh', 'masterPh', 'bodyStyle', 'defaultTextStyle', 'otherStyle'],
  },
  {
    id: 'text-ph-tbl',
    ph: 'type="tbl" idx="1"',
    levels: ['layoutPh', 'masterPh', 'bodyStyle', 'defaultTextStyle', 'otherStyle'],
  },
  // A master with no p:txStyles, so the built-in floor is reached, under three placeholder types.
  {
    id: 'text-ph-bare',
    ph: 'idx="1"',
    levels: [],
    bare: { layout: 'idx="1"', master: 'type="body" idx="1"' },
  },
  {
    id: 'text-ph-tbl-bare',
    ph: 'type="tbl" idx="1"',
    levels: [],
    bare: { layout: 'type="tbl" idx="1"', master: 'type="body" idx="1"' },
  },
  {
    id: 'text-ph-title-bare',
    ph: 'type="title"',
    levels: [],
    bare: { layout: 'type="title"', master: 'type="title"' },
  },
];

function textDeck(pkg: TextPackage): DeckSpec {
  const has = (level: TextLevel): boolean => pkg.levels.includes(level);
  const txStyles =
    has('otherStyle') || has('bodyStyle')
      ? '<p:txStyles><p:titleStyle/>' +
        (has('bodyStyle') ? listStyle('bodyStyle', 'p:bodyStyle', 1, true) : '<p:bodyStyle/>') +
        (has('otherStyle') ? listStyle('otherStyle', 'p:otherStyle') : '<p:otherStyle/>') +
        '</p:txStyles>'
      : undefined;
  const rect = { x: ORIGIN.x, y: ORIGIN.y, w: 360, h: 180 };
  return {
    id: pkg.id,
    group: 'text',
    theme: 'A',
    read: 'full',
    ...(has('defaultTextStyle')
      ? { defaultTextStyle: listStyle('defaultTextStyle', 'p:defaultTextStyle') }
      : {}),
    ...(txStyles === undefined ? {} : { txStyles }),
    ...(has('masterPh')
      ? {
          masterShapes: [
            shape({
              id: 3,
              name: 'Body',
              ph: 'type="body" idx="1"',
              rect,
              lstStyle: listStyle('masterPh'),
            }),
          ],
        }
      : {}),
    ...(has('layoutPh')
      ? {
          layoutShapes: [
            shape({ id: 3, name: 'Content', ph: 'idx="1"', rect, lstStyle: listStyle('layoutPh') }),
          ],
        }
      : {}),
    ...(pkg.bare === undefined
      ? {}
      : {
          masterShapes: [shape({ id: 3, name: 'Placeholder', ph: pkg.bare.master, rect })],
          layoutShapes: [shape({ id: 3, name: 'Placeholder', ph: pkg.bare.layout, rect })],
        }),
    slides: [idOf('medium2'), null, idOf('tableGrid')].map((style) => ladderTable(style, pkg.ph)),
  };
}

/** Every deck, in the order `read.ps1` opens them: the control first, and again last. */
export function allDecks(): DeckSpec[] {
  const floor = { defaultTextStyle: FLOOR_DEFAULT_TEXT_STYLE, txStyles: FLOOR_TX_STYLES };
  const decks: DeckSpec[] = [
    {
      id: 'control',
      group: 'control',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: (['medium2', 'themed1', 'dark1'] as const).map((key) => ({
        rows: 5,
        cols: 5,
        style: idOf(key),
        flags: flagsOf(ALL),
      })),
    },
  ];

  const sweepStyles: readonly (readonly [string, string | null])[] = [
    ...BUILTINS.map((s, i): [string, string | null] => [`sweep-${two(i + 1)}`, s.id]),
    ['sweep-none', null],
    ['sweep-unknown', UNKNOWN_STYLE],
  ];
  for (const [id, style] of sweepStyles) {
    decks.push({
      id,
      group: 'sweep',
      theme: 'A',
      read: 'sweep',
      ...floor,
      slides: slidesFor(style, [[5, 5]], ALL_MASKS),
    });
  }
  BUILTINS.forEach((style, i) => {
    decks.push({
      id: `ends-${two(i + 1)}`,
      group: 'ends',
      theme: 'A',
      read: 'sweep',
      ...floor,
      slides: slidesFor(style.id, END_SHAPES, END_MASKS),
    });
  });
  const representatives = new Set(
    BUILTINS.filter(
      (s) => !/ - Accent [2-6]|Accent 3\/Accent 4|Accent 5\/Accent 6/.test(s.name),
    ).map((s) => s.id),
  );
  BUILTINS.forEach((style, i) => {
    decks.push({
      id: `sides-${two(i + 1)}`,
      group: 'sides',
      theme: 'B',
      read: representatives.has(style.id) ? 'full' : 'sweep',
      ...floor,
      slides: slidesFor(style.id, [[5, 5]], SIDE_MASKS),
    });
  });

  const directStyles: readonly (readonly [string, string | null, number])[] = [
    ['medium2', idOf('medium2'), HEADER_BANDED],
    ['medium1', idOf('medium1'), HEADER_BANDED],
    ['dark1', idOf('dark1'), ALL],
    ['themed1', idOf('themed1'), 0b010101],
    ['tablegrid', idOf('tableGrid'), HEADER_BANDED],
    ['none', null, HEADER_BANDED],
  ];
  for (const [key, style, mask] of directStyles) {
    decks.push({
      id: `direct-${key}`,
      group: 'direct',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: DIRECT_SLIDES.map((s) => ({
        name: s.name,
        rows: 5,
        cols: 5,
        style,
        flags: flagsOf(s.mask ?? mask),
        cells: s.cells,
      })),
    });
  }

  for (const key of ['medium2', 'medium3', 'dark1', 'themed2'] as const) {
    decks.push({
      id: `merge-${key}`,
      group: 'merge',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: MERGE_SLIDES.map((s) => ({
        name: s.name,
        rows: 5,
        cols: 5,
        style: idOf(key),
        flags: flagsOf(s.mask),
        cells: s.cells,
      })),
    });
  }

  for (const key of ['dark1', 'themed1', 'medium3', 'medium2'] as const) {
    decks.push({
      id: `rtl-${key}`,
      group: 'rtl',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: [
        ...[ALL, 0b101100, 0b100000, 0b000101].map((mask) => ({
          rows: 5,
          cols: 6,
          style: idOf(key),
          flags: flagsOf(mask),
          rtl: true,
        })),
        {
          rows: 5,
          cols: 6,
          style: idOf(key),
          flags: flagsOf(0),
          rtl: true,
          cells: {
            '2,2': { tcPr: `<a:tcPr>${line('lnL', 2, DIRECT.rtlLeft)}</a:tcPr>` },
            '2,4': { tcPr: `<a:tcPr>${line('lnR', 2, DIRECT.rtlRight)}</a:tcPr>` },
          },
        },
      ],
    });
  }

  decks.push({
    id: 'tblpr',
    group: 'tblpr',
    theme: 'A',
    read: 'full',
    ...floor,
    slides: [null, idOf('noGrid'), idOf('themed1'), idOf('medium2')].flatMap((style) =>
      [solid('C9F001'), '<a:noFill/>'].map((fill) => ({
        rows: 5,
        cols: 5,
        style,
        flags: flagsOf(HEADER_BANDED),
        tblPrFill: fill,
      })),
    ),
  });

  decks.push({
    id: 'background',
    group: 'background',
    theme: 'C',
    read: 'full',
    ...floor,
    slides: [
      ...(['themed1', 'themed2', 'light2', 'light2a1'] as const).flatMap((key) =>
        [0, HEADER_BANDED, ALL].map((mask) => ({
          rows: 5,
          cols: 5,
          style: idOf(key),
          flags: flagsOf(mask),
        })),
      ),
      ...(['themed2', 'medium2'] as const).flatMap((key) => [
        {
          rows: 5,
          cols: 5,
          style: idOf(key),
          flags: flagsOf(HEADER_BANDED),
          tblPrFill: gradient('<a:srgbClr val="C9C001"/>', '<a:srgbClr val="C9C002"/>', 0),
        },
        {
          rows: 5,
          cols: 5,
          style: idOf(key),
          flags: flagsOf(HEADER_BANDED),
          tblPrEffects:
            '<a:effectLst><a:outerShdw blurRad="0" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="C9C003"/></a:outerShdw></a:effectLst>',
        },
      ]),
      ...(['themed1', 'themed2'] as const).map((key) => ({
        rows: 5,
        cols: 5,
        style: idOf(key),
        flags: flagsOf(0),
        extScale: 1.5,
      })),
    ],
  });

  decks.push(...lineAndEffectDecks(floor));
  for (const pkg of TEXT_PACKAGES) decks.push(textDeck(pkg));
  return decks;
}

const dashed = (tag: string, pt: number, hex: string, dash: string): string =>
  `<a:${tag} w="${emu(pt)}" cap="flat" cmpd="sng" algn="ctr">${solid(hex)}<a:prstDash val="${dash}"/><a:round/></a:${tag}>`;
const shadowList = (degrees: number): string =>
  `<a:effectLst><a:outerShdw blurRad="0" dist="76200" dir="${String(degrees * 60000)}" algn="tl" rotWithShape="0"><a:srgbClr val="C9C004"/></a:outerShdw></a:effectLst>`;

/** Dashed lines, a padded row, diagonals, the background's effect, and a deck PowerPoint saves. */
function lineAndEffectDecks(floor: Pick<DeckSpec, 'defaultTextStyle' | 'txStyles'>): DeckSpec[] {
  const styled = (key: keyof typeof REFERENCE | null): string | null =>
    key === null ? null : idOf(key);
  const table = (
    name: string,
    style: string | null,
    cells: Record<string, CellSpec>,
    extra: Partial<TableSpec> = {},
  ): TableSpec => ({
    name,
    rows: 5,
    cols: 5,
    style,
    flags: flagsOf(HEADER_BANDED),
    cells,
    ...extra,
  });
  return [
    {
      id: 'dash',
      group: 'dash',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: (['medium2', 'dark1', null] as const).flatMap((key) =>
        (
          [
            ['dash', `<a:tcPr>${dashed('lnB', 1, DIRECT.dashed, 'dash')}</a:tcPr>`],
            ['long-dash', `<a:tcPr>${dashed('lnR', 2, DIRECT.dashed, 'lgDash')}</a:tcPr>`],
            ['dot-only', '<a:tcPr><a:lnB><a:prstDash val="sysDot"/></a:lnB></a:tcPr>'],
          ] as const
        ).map(([name, tcPr]) => table(name, styled(key), { '2,2': { tcPr } })),
      ),
    },
    {
      id: 'direct-padded',
      group: 'direct',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: (['medium2', 'dark1', null] as const).map((key) =>
        table(
          'padded-above',
          styled(key),
          {
            '2,3': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.padded)}</a:tcPr>` },
            '2,1': { tcPr: `<a:tcPr>${line('lnT', 2, DIRECT.paddedBeside)}</a:tcPr>` },
          },
          { short: { '1': 3 } },
        ),
      ),
    },
    {
      id: 'merge-diagonals',
      group: 'merge',
      theme: 'A',
      read: 'full',
      resave: true,
      ...floor,
      slides: (['medium2', 'dark1'] as const).flatMap((key) => [
        table('diagonal-wide', styled(key), {
          '1,1': {
            attrs: ' gridSpan="2"',
            tcPr: `<a:tcPr>${line('lnTlToBr', 2, DIRECT.diagonalAnchor)}</a:tcPr>`,
          },
          '1,2': covered(
            true,
            false,
            `<a:tcPr>${line('lnBlToTr', 2, DIRECT.diagonalCovered)}</a:tcPr>`,
          ),
        }),
        table('diagonal-tall', styled(key), {
          '1,1': {
            attrs: ' rowSpan="2"',
            tcPr: `<a:tcPr>${line('lnBlToTr', 2, DIRECT.diagonalAnchor)}</a:tcPr>`,
          },
          '2,1': covered(
            false,
            true,
            `<a:tcPr>${line('lnTlToBr', 2, DIRECT.diagonalCovered)}</a:tcPr>`,
          ),
        }),
        table('covered-saved', styled(key), {
          '1,1': { attrs: ' rowSpan="2"' },
          '2,1': covered(
            false,
            true,
            `<a:tcPr>${line('lnB', 1, DIRECT.covered)}${solid(DIRECT.coveredFill)}</a:tcPr>`,
          ),
        }),
      ]),
    },
    {
      id: 'rtl-diagonals',
      group: 'rtl',
      theme: 'A',
      read: 'full',
      ...floor,
      slides: (['dark1', 'medium2'] as const).map((key) => ({
        ...table('rtl-diagonals', styled(key), {
          '2,1': { tcPr: `<a:tcPr>${line('lnTlToBr', 2, DIRECT.rtlDown)}</a:tcPr>` },
          '2,4': { tcPr: `<a:tcPr>${line('lnBlToTr', 2, DIRECT.rtlUp)}</a:tcPr>` },
        }),
        cols: 6,
        flags: flagsOf(0),
        rtl: true,
      })),
    },
    {
      id: 'background-effects',
      group: 'background',
      theme: 'C',
      read: 'full',
      ...floor,
      slides: [
        table('shadow-right', styled('themed2'), {}, { tblPrEffects: shadowList(0) }),
        table('shadow-below', styled('themed1'), {}, { tblPrEffects: shadowList(90) }),
        table('no-effects', styled('themed1'), {}, { tblPrEffects: '<a:effectLst/>' }),
        table('frame-wider', styled('themed1'), {}, { flags: flagsOf(ALL), extScale: 1.25 }),
        table('frame-wider', styled('themed2'), {}, { flags: flagsOf(ALL), extScale: 1.25 }),
      ],
    },
  ];
}

/* -------------------------------------------------------------------------- */
/* the readings                                                               */
/* -------------------------------------------------------------------------- */

/** Which parts reach a cell. `spec` is ECMA-376's text; the rest are the readings one writes instead. */
export const APPLICABILITY = [
  'spec',
  'cornersByPosition',
  'cornersOnEither',
  'bandsFromZero',
  'rowBandsOverEnds',
  'colBandsOverEnds',
  'rowBandsOnEnds',
  'colBandsOnEnds',
] as const;
export type Applicability = (typeof APPLICABILITY)[number];

/** Lowest first: a later part beats an earlier one. */
export const ORDERS = {
  schema: [
    'wholeTbl',
    'band1H',
    'band2H',
    'band1V',
    'band2V',
    'lastCol',
    'firstCol',
    'lastRow',
    'seCell',
    'swCell',
    'firstRow',
    'neCell',
    'nwCell',
  ],
  spreadsheetml: [
    'wholeTbl',
    'band1V',
    'band2V',
    'band1H',
    'band2H',
    'lastCol',
    'firstCol',
    'firstRow',
    'lastRow',
    'nwCell',
    'neCell',
    'swCell',
    'seCell',
  ],
  colsOverRows: [
    'wholeTbl',
    'band1H',
    'band2H',
    'band1V',
    'band2V',
    'firstRow',
    'lastRow',
    'firstCol',
    'lastCol',
    'nwCell',
    'neCell',
    'swCell',
    'seCell',
  ],
  bandsOverEnds: [
    'wholeTbl',
    'lastCol',
    'firstCol',
    'lastRow',
    'firstRow',
    'band1H',
    'band2H',
    'band1V',
    'band2V',
    'nwCell',
    'neCell',
    'swCell',
    'seCell',
  ],
} as const satisfies Readonly<Record<string, readonly TableStylePartName[]>>;
export type Order = keyof typeof ORDERS;

/** Per property, or the highest part holding the element takes all of it. */
export const MERGES = ['perProperty', 'wholeElement'] as const;
export type Merge = (typeof MERGES)[number];

/**
 * How a grid edge's line is chosen. `perEdge`: every part of either cell competes, a part covering
 * both offering its inside edge. The next five resolve each cell's own side first and then combine
 * the two; the last two map a part's edges onto cells differently and compete as `perEdge` does.
 */
export const EDGE_MODELS = [
  'perEdge',
  'heavier',
  'visible',
  'upperLeft',
  'lowerRight',
  'stacked',
  'cellRelative',
  'tableRelative',
] as const;
export type EdgeModel = (typeof EDGE_MODELS)[number];

/** One translucent fill wins, or every applicable fill is painted in order. */
export const FILL_PAINTS = ['winner', 'stacked'] as const;
export type FillPaint = (typeof FILL_PAINTS)[number];

export type TextColourSource = 'explicit' | 'fontRef';

export interface Reading {
  readonly applicability: Applicability;
  readonly order: Order;
  readonly merge: Merge;
}

export interface Shape {
  readonly rows: number;
  readonly cols: number;
  readonly flags: Flags;
}

/** The parts reaching cell (r, c) under one reading of applicability. */
export function partsAt(
  reading: Applicability,
  table: Shape,
  r: number,
  c: number,
): Set<TableStylePartName> {
  const { rows, cols, flags } = table;
  const parts = new Set<TableStylePartName>(['wholeTbl']);
  const firstRow = flags.firstRow && r === 0;
  const lastRow = flags.lastRow && r === rows - 1;
  const firstCol = flags.firstCol && c === 0;
  const lastCol = flags.lastCol && c === cols - 1;
  if (firstRow) parts.add('firstRow');
  if (lastRow) parts.add('lastRow');
  if (firstCol) parts.add('firstCol');
  if (lastCol) parts.add('lastCol');

  // `over` counts from zero across the ends; `on` reaches the ends but counts as ECMA counts.
  const band = (
    on: boolean,
    index: number,
    ends: boolean,
    skip: boolean,
    how: 'over' | 'on' | null,
  ): number | null => {
    if (!on) return null;
    if (how === 'over') return index;
    if (ends && how !== 'on') return null;
    return reading === 'bandsFromZero' ? index : index - (skip ? 1 : 0);
  };
  const reach = (over: Applicability, onEnds: Applicability): 'over' | 'on' | null =>
    reading === over ? 'over' : reading === onEnds ? 'on' : null;
  const h = band(
    flags.bandRow,
    r,
    firstRow || lastRow,
    flags.firstRow,
    reach('rowBandsOverEnds', 'rowBandsOnEnds'),
  );
  if (h !== null) parts.add(Math.abs(h) % 2 === 0 ? 'band1H' : 'band2H');
  const v = band(
    flags.bandCol,
    c,
    firstCol || lastCol,
    flags.firstCol,
    reach('colBandsOverEnds', 'colBandsOnEnds'),
  );
  if (v !== null) parts.add(Math.abs(v) % 2 === 0 ? 'band1V' : 'band2V');

  const corner = (atRow: boolean, atCol: boolean, rowFlag: boolean, colFlag: boolean): boolean => {
    if (!atRow || !atCol) return false;
    if (reading === 'cornersByPosition') return true;
    if (reading === 'cornersOnEither') return rowFlag || colFlag;
    return rowFlag && colFlag;
  };
  const top = r === 0;
  const bottom = r === rows - 1;
  const left = c === 0;
  const right = c === cols - 1;
  if (corner(top, left, flags.firstRow, flags.firstCol)) parts.add('nwCell');
  if (corner(top, right, flags.firstRow, flags.lastCol)) parts.add('neCell');
  if (corner(bottom, left, flags.lastRow, flags.firstCol)) parts.add('swCell');
  if (corner(bottom, right, flags.lastRow, flags.lastCol)) parts.add('seCell');
  return parts;
}

const PARTS_CACHE = new Map<string, Set<TableStylePartName>[][]>();

/** `partsAt` for every cell of a shape, memoised: the sweep asks the same shapes of every style. */
export function partsGrid(reading: Applicability, table: Shape): Set<TableStylePartName>[][] {
  const key = `${reading}|${String(table.rows)}x${String(table.cols)}|${String(maskOf(table.flags))}`;
  let grid = PARTS_CACHE.get(key);
  if (grid === undefined) {
    grid = Array.from({ length: table.rows }, (_, r) =>
      Array.from({ length: table.cols }, (_, c) => partsAt(reading, table, r, c)),
    );
    PARTS_CACHE.set(key, grid);
  }
  return grid;
}

/** The parts, highest first. */
function ranked(parts: Iterable<TableStylePartName>, order: Order): TableStylePartName[] {
  const list: readonly TableStylePartName[] = ORDERS[order];
  return [...parts].sort((a, b) => list.indexOf(b) - list.indexOf(a));
}

function partOf(style: TableStyle, name: TableStylePartName): TableStylePart | undefined {
  return style.parts[name];
}

/** The fill a cell takes: the highest part that states one, or holds a `tcStyle`. */
export function fillWinner(
  style: TableStyle,
  parts: Iterable<TableStylePartName>,
  reading: Reading,
): { part: TableStylePartName; fill: Themeable<PaintFill> } | null {
  for (const part of ranked(parts, reading.order)) {
    const cell = partOf(style, part)?.cell;
    if (cell === undefined) continue;
    if (cell.fill !== undefined) return { part, fill: cell.fill };
    if (reading.merge === 'wholeElement') return null;
  }
  return null;
}

/** Every applicable part's fill, lowest first: what a cell paints if translucent fills stack. */
export function fillStack(
  style: TableStyle,
  parts: Iterable<TableStylePartName>,
  order: Order,
): Themeable<PaintFill>[] {
  return ranked(parts, order)
    .reverse()
    .map((part) => partOf(style, part)?.cell?.fill)
    .filter((f): f is Themeable<PaintFill> => f !== undefined);
}

export interface TextWinner {
  readonly color: Color | null;
  readonly bold: boolean;
  /** The part that states the face, or `null` where none does. */
  readonly font: TableStylePartName | null;
}

/** The text colour, bold and face a cell takes from the style; `null` where no part says. */
export function textWinner(
  style: TableStyle,
  parts: Iterable<TableStylePartName>,
  reading: Reading,
  source: TextColourSource,
): TextWinner {
  let color: Color | null | undefined;
  let bold: boolean | undefined;
  let font: TableStylePartName | null | undefined;
  for (const part of ranked(parts, reading.order)) {
    const text = partOf(style, part)?.text;
    if (text === undefined) continue;
    const whole = reading.merge === 'wholeElement';
    const stated =
      source === 'explicit'
        ? text.color
        : text.font?.kind === 'ref'
          ? (text.font.ref.color ?? undefined)
          : undefined;
    if (color === undefined && (stated !== undefined || whole)) color = stated ?? null;
    if (bold === undefined && (text.b !== undefined || whole)) bold = text.b === 'on';
    if (font === undefined && (text.font !== undefined || whole))
      font = text.font === undefined ? null : part;
    if (color !== undefined && bold !== undefined && font !== undefined) break;
  }
  return { color: color ?? null, bold: bold ?? false, font: font ?? null };
}

export type Side = 'top' | 'bottom' | 'left' | 'right';
type EdgeRole = Side | 'insideH' | 'insideV';

const OPPOSITE: Readonly<Record<Side, Side>> = {
  top: 'bottom',
  bottom: 'top',
  left: 'right',
  right: 'left',
};
const INSIDE: Readonly<Record<Side, 'insideH' | 'insideV'>> = {
  top: 'insideH',
  bottom: 'insideH',
  left: 'insideV',
  right: 'insideV',
};

/** A grid edge: `h` is the top of row `r` at column `c` (r = rows is the bottom); `v` the left of column `c`. */
export interface Edge {
  readonly axis: 'h' | 'v';
  readonly r: number;
  readonly c: number;
}

/** Every grid edge of a table, horizontal first. */
export function edgesOf(table: Shape): Edge[] {
  const out: Edge[] = [];
  for (let r = 0; r <= table.rows; r++)
    for (let c = 0; c < table.cols; c++) out.push({ axis: 'h', r, c });
  for (let r = 0; r < table.rows; r++)
    for (let c = 0; c <= table.cols; c++) out.push({ axis: 'v', r, c });
  return out;
}

/** The cells either side of an edge: upper or left first, `null` beyond the table. */
function cellsOf(table: Shape, edge: Edge): [[number, number] | null, [number, number] | null] {
  if (edge.axis === 'h') {
    return [
      edge.r > 0 ? [edge.r - 1, edge.c] : null,
      edge.r < table.rows ? [edge.r, edge.c] : null,
    ];
  }
  return [edge.c > 0 ? [edge.r, edge.c - 1] : null, edge.c < table.cols ? [edge.r, edge.c] : null];
}

interface Claim {
  readonly part: TableStylePartName;
  readonly role: EdgeRole;
}

type StyleLine = { readonly part: TableStylePartName; readonly line: Themeable<PaintLine> };

/** The highest claim that states its edge; with `wholeElement` the highest holding a `tcBdr` decides. */
function claimWinner(
  style: TableStyle,
  claims: readonly Claim[],
  reading: Reading,
): StyleLine | null {
  const byPart = new Map(claims.map((c) => [c.part, c.role]));
  for (const part of ranked(byPart.keys(), reading.order)) {
    const role = byPart.get(part);
    const borders = partOf(style, part)?.cell?.borders;
    if (role === undefined || borders === undefined) continue;
    const stated = borders[role];
    if (stated !== undefined) return { part, line: stated };
    if (reading.merge === 'wholeElement') return null;
  }
  return null;
}

/** One cell's own claim on its side facing `edge`, region-relative: its parts only. */
function sideWinner(
  style: TableStyle,
  grid: Set<TableStylePartName>[][],
  cell: [number, number],
  other: [number, number] | null,
  side: Side,
  reading: Reading,
): StyleLine | null {
  const own = grid[cell[0]]?.[cell[1]] ?? new Set<TableStylePartName>();
  const theirs =
    other === null ? new Set<TableStylePartName>() : (grid[other[0]]?.[other[1]] ?? new Set());
  const claims = [...own].map((part) => ({ part, role: theirs.has(part) ? INSIDE[side] : side }));
  return claimWinner(style, claims, reading);
}

/** A line as the pixels would show it: colour, width in points, compound. */
export interface DrawnLine {
  readonly rgb: string;
  readonly weight: number;
  readonly cmpd: string;
}

const hexOf = (rgba: Rgba): string => toHexColor(rgba).toUpperCase();

export function colorHex(color: Color, palette: Palette, phClr?: Rgba): string {
  return hexOf(
    resolveColor(color, { ...palette.colors, ...(phClr === undefined ? {} : { phClr }) }),
  );
}

/** A style line resolved against the theme, or `null` for `a:noFill`. */
export function drawnLine(line: Themeable<PaintLine>, palette: Palette): DrawnLine | null {
  let resolved: PaintLine | null;
  let ph: Rgba | undefined;
  if (line.kind === 'value') resolved = line.value;
  else {
    resolved = styleMatrixLine(line.ref, palette.theme);
    ph = line.ref.color === null ? undefined : resolveColor(line.ref.color, palette.colors);
  }
  if (resolved === null) return null;
  const fill = resolved.fill;
  if (fill === null || fill.type === 'none') return null;
  if (fill.type !== 'solid')
    throw new Error(`a built-in line fill of type ${fill.type} is not predicted`);
  if (resolved.w === null) throw new Error('a built-in line with no width is not predicted');
  if (resolved.dash !== null && resolved.dash.kind !== 'preset')
    throw new Error('a custom dash is not predicted');
  return {
    rgb: colorHex(fill.color, palette, ph),
    weight: resolved.w / EMU_PER_POINT,
    cmpd: resolved.cmpd ?? 'sng',
  };
}

const heavierOf = (a: DrawnLine | null, b: DrawnLine | null): DrawnLine | null =>
  a === null ? b : b === null ? a : b.weight > a.weight ? b : a;

/** The edge each end part faces the rest of the table with. */
const KIND_ROLE: Readonly<Partial<Record<TableStylePartName, Side>>> = {
  firstRow: 'bottom',
  lastRow: 'top',
  firstCol: 'right',
  lastCol: 'left',
};

/**
 * `perEdge` over explicit part sets either side of an edge, `null` beyond the table: a part in
 * both offers its inside edge, a part on one side its edge facing the other.
 */
export function edgeLinesBetween(
  style: TableStyle,
  before: ReadonlySet<TableStylePartName> | null,
  after: ReadonlySet<TableStylePartName> | null,
  axis: 'h' | 'v',
  reading: Reading,
  palette: Palette,
  roles: 'geometry' | 'kind' = 'geometry',
): DrawnLine[] {
  const sideA: Side = axis === 'h' ? 'bottom' : 'right';
  const claims: Claim[] = [];
  for (const part of new Set([...(before ?? []), ...(after ?? [])])) {
    const inA = before?.has(part) === true;
    const inB = after?.has(part) === true;
    const geometric = inA && inB ? INSIDE[sideA] : inA ? sideA : OPPOSITE[sideA];
    // An end part meeting a cell outside it offers the edge its name says, whichever way it runs.
    const named = before !== null && after !== null && inA !== inB ? KIND_ROLE[part] : undefined;
    claims.push({ part, role: roles === 'kind' && named !== undefined ? named : geometric });
  }
  const winner = claimWinner(style, claims, reading);
  const drawn = winner === null ? null : drawnLine(winner.line, palette);
  return drawn === null ? [] : [drawn];
}

/**
 * The lines one edge model draws on a grid edge, bottom first: empty for none, two for `stacked`
 * where both neighbours draw.
 */
export function edgeLines(
  style: TableStyle,
  table: Shape,
  edge: Edge,
  reading: Reading,
  model: EdgeModel,
  palette: Palette,
): DrawnLine[] {
  const grid = partsGrid(reading.applicability, table);
  const [a, b] = cellsOf(table, edge);
  const sideA: Side = edge.axis === 'h' ? 'bottom' : 'right';
  const sideB: Side = OPPOSITE[sideA];
  const resolve = (w: StyleLine | null): DrawnLine | null =>
    w === null ? null : drawnLine(w.line, palette);

  if (model === 'perEdge' || model === 'cellRelative' || model === 'tableRelative') {
    const claims: Claim[] = [];
    const add = (
      cell: [number, number] | null,
      other: [number, number] | null,
      side: Side,
    ): void => {
      if (cell === null) return;
      const own = grid[cell[0]]?.[cell[1]] ?? new Set<TableStylePartName>();
      const theirs =
        other === null ? new Set<TableStylePartName>() : (grid[other[0]]?.[other[1]] ?? new Set());
      for (const part of own) {
        let role: EdgeRole;
        if (model === 'cellRelative') role = side;
        else if (model === 'tableRelative') role = other === null ? side : INSIDE[side];
        else role = theirs.has(part) ? INSIDE[side] : side;
        if (!claims.some((c) => c.part === part && c.role === role)) claims.push({ part, role });
      }
    };
    add(a, b, sideA);
    add(b, a, sideB);
    const drawn = resolve(claimWinner(style, claims, reading));
    return drawn === null ? [] : [drawn];
  }

  const winA = a === null ? null : sideWinner(style, grid, a, b, sideA, reading);
  const winB = b === null ? null : sideWinner(style, grid, b, a, sideB, reading);
  const lineA = resolve(winA);
  const lineB = resolve(winB);
  let drawn: (DrawnLine | null)[];
  if (model === 'stacked') drawn = [lineA, lineB];
  else if (model === 'heavier') drawn = [heavierOf(lineA, lineB)];
  else if (model === 'visible') {
    const rankOf = (w: StyleLine | null): number =>
      w === null ? -1 : (ORDERS[reading.order] as readonly TableStylePartName[]).indexOf(w.part);
    drawn = [
      lineA === null ? lineB : lineB === null ? lineA : rankOf(winB) > rankOf(winA) ? lineB : lineA,
    ];
  } else if (model === 'upperLeft') drawn = [a === null ? lineB : lineA];
  else drawn = [b === null ? lineA : lineB];
  return drawn.filter((l): l is DrawnLine => l !== null);
}

/* -------------------------------------------------------------------------- */
/* fills and text as COM and the pixels report them                           */
/* -------------------------------------------------------------------------- */

/** `Shape.Fill` as COM reports it: a colour and a transparency, or nothing. */
export interface FillReport {
  readonly visible: boolean;
  readonly rgb: string;
  /** `1 - alpha`, to the thousandth: every built-in alpha is a whole per cent. */
  readonly transparency: number;
}

const NO_FILL: FillReport = { visible: false, rgb: '', transparency: 1 };

function fillReport(fill: PaintFill, palette: Palette, phClr?: Rgba): FillReport {
  if (fill.type === 'none') return NO_FILL;
  if (fill.type !== 'solid') throw new Error(`a fill of type ${fill.type} is not predicted`);
  const rgba = resolveColor(fill.color, {
    ...palette.colors,
    ...(phClr === undefined ? {} : { phClr }),
  });
  return { visible: true, rgb: hexOf(rgba), transparency: Math.round((1 - rgba.a) * 1000) / 1000 };
}

/** A style fill, with a `fillRef` followed into the theme and invoked with its colour. */
export function predictFill(fill: Themeable<PaintFill> | null, palette: Palette): FillReport {
  if (fill === null) return NO_FILL;
  if (fill.kind === 'value') return fillReport(fill.value, palette);
  const resolved = styleMatrixFill(fill.ref, palette.theme);
  if (resolved === null) return NO_FILL;
  const ph = fill.ref.color === null ? undefined : resolveColor(fill.ref.color, palette.colors);
  return fillReport(resolved, palette, ph);
}

export const fillText = (f: FillReport): string =>
  f.visible ? `${f.rgb}/${String(f.transparency)}` : 'none';

/** A text report: colour, bold, face. */
export const textText = (rgb: string, bold: boolean, face: string): string =>
  `${rgb}/${bold ? 'b' : '-'}/${face}`;

/** The style parts a theme font reference names: every built-in says `minor`. */
function faceOf(style: TableStyle, part: TableStylePartName, palette: Palette): string {
  const font = partOf(style, part)?.text?.font;
  if (font === undefined) throw new Error(`${part} states no face`);
  if (font.kind === 'value') {
    const latin = font.value.latin;
    if (latin === null) throw new Error(`${part} names no latin face`);
    return latin;
  }
  if (font.ref.idx === 'none') throw new Error(`${part} has fontRef idx="none"`);
  const face = palette.theme.fonts[font.ref.idx].latin;
  if (face === null) throw new Error(`the theme names no ${font.ref.idx} face`);
  return face;
}

/** The floor's colour and face, which a cell shows where its style says nothing. */
export function floorText(palette: Palette): { rgb: string; face: string } {
  return {
    rgb: colorHex({ space: 'scheme', name: FLOOR.slot, transforms: [] }, palette),
    face: FLOOR.face,
  };
}

/** A built-in, parsed once per process: the analysis reads, never edits, what it returns. */
const PARSED = new Map<string, TableStyle>();
export function styleById(id: string): TableStyle {
  let style = PARSED.get(id);
  if (style === undefined) {
    style = builtinTableStyle(id);
    if (style === undefined) throw new Error(`${id} is no built-in`);
    PARSED.set(id, style);
  }
  return style;
}

/** One reading's prediction of what COM reports for a cell's fill, or its first character's text. */
export function predictCom(
  style: TableStyle,
  table: Shape,
  r: number,
  c: number,
  unit: 'fill' | 'text',
  reading: Reading,
  source: TextColourSource,
  palette: Palette,
): string {
  const parts = partsGrid(reading.applicability, table)[r]?.[c];
  if (parts === undefined) throw new Error(`no cell ${String(r)},${String(c)}`);
  if (unit === 'fill')
    return fillText(predictFill(fillWinner(style, parts, reading)?.fill ?? null, palette));
  const text = textWinner(style, parts, reading, source);
  const floor = floorText(palette);
  return textText(
    text.color === null ? floor.rgb : colorHex(text.color, palette),
    text.bold,
    text.font === null ? floor.face : faceOf(style, text.font, palette),
  );
}

export const lineText = (lines: readonly DrawnLine[]): string =>
  lines.length === 0
    ? 'none'
    : lines.map((l) => `${String(l.weight)}/${l.rgb}/${l.cmpd}`).join('+');

/** A key naming one observation: a cell's COM fill or text, or a grid edge's pixels. */
export const cellKey = (deck: string, slide: number, r: number, c: number, unit: string): string =>
  `${deck}#${String(slide)}@${String(r)},${String(c)}:${unit}`;
export const edgeKey = (deck: string, slide: number, edge: Edge): string =>
  `${deck}#${String(slide)}@${edge.axis}${String(edge.r)},${String(edge.c)}`;

export interface Scenario {
  readonly reading: Reading;
  readonly edges: EdgeModel;
  readonly source: TextColourSource;
}

export const BASELINE: Scenario = {
  reading: { applicability: 'spec', order: 'schema', merge: 'perProperty' },
  edges: 'perEdge',
  source: 'explicit',
};

/** Every styled table with no direct formatting in `decks`: what `predictions` covers. */
export function* styledTables(
  decks: readonly DeckSpec[],
): Generator<{ deck: DeckSpec; slide: number; table: TableSpec; style: TableStyle }> {
  for (const deck of decks) {
    for (const [slide, table] of deck.slides.entries()) {
      if (table.style === null || table.cells !== undefined || table.rtl === true) continue;
      if (BUILTINS.every((s) => s.id !== table.style)) continue;
      if (
        table.tblPrFill !== undefined ||
        table.tblPrEffects !== undefined ||
        table.extScale !== undefined
      )
        continue;
      yield { deck, slide, table, style: styleById(table.style) };
    }
  }
}

/** Every COM unit and every grid edge of every styled table, and what one scenario predicts. */
export function* predictions(
  decks: readonly DeckSpec[],
  scenario: Scenario,
): Generator<[string, string]> {
  for (const { deck, slide, table, style } of styledTables(decks)) {
    const palette = paletteOf(deck.theme);
    for (let r = 0; r < table.rows; r++) {
      for (let c = 0; c < table.cols; c++) {
        for (const unit of ['fill', 'text'] as const) {
          yield [
            cellKey(deck.id, slide, r, c, unit),
            predictCom(style, table, r, c, unit, scenario.reading, scenario.source, palette),
          ];
        }
      }
    }
    for (const edge of edgesOf(table)) {
      yield [
        edgeKey(deck.id, slide, edge),
        lineText(edgeLines(style, table, edge, scenario.reading, scenario.edges, palette)),
      ];
    }
  }
}

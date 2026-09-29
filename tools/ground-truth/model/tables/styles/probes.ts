/**
 * Experiment C8 - the probes: which style an `a:tableStyleId` draws, where PowerPoint looks for
 * it, and what it writes back. One question per package, so a repair names its cause; ADR 0063.
 */

import { entry, readZip } from '../../../lib/zip.ts';
import { COL, FRAME, ROW, TABLE_URI } from '../grid/probes.ts';

/**
 * PowerPoint's own names for its built-in styles, in its gallery's order: the `style` union of
 * Office.js `PowerPoint.TableStyleSettings`,
 * https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.tablestylesettings
 */
export const OFFICE_JS_TABLE_STYLES = [
  'NoStyleNoGrid',
  'ThemedStyle1Accent1',
  'ThemedStyle1Accent2',
  'ThemedStyle1Accent3',
  'ThemedStyle1Accent4',
  'ThemedStyle1Accent5',
  'ThemedStyle1Accent6',
  'NoStyleTableGrid',
  'ThemedStyle2Accent1',
  'ThemedStyle2Accent2',
  'ThemedStyle2Accent3',
  'ThemedStyle2Accent4',
  'ThemedStyle2Accent5',
  'ThemedStyle2Accent6',
  'LightStyle1',
  'LightStyle1Accent1',
  'LightStyle1Accent2',
  'LightStyle1Accent3',
  'LightStyle1Accent4',
  'LightStyle1Accent5',
  'LightStyle1Accent6',
  'LightStyle2',
  'LightStyle2Accent1',
  'LightStyle2Accent2',
  'LightStyle2Accent3',
  'LightStyle2Accent4',
  'LightStyle2Accent5',
  'LightStyle2Accent6',
  'LightStyle3',
  'LightStyle3Accent1',
  'LightStyle3Accent2',
  'LightStyle3Accent3',
  'LightStyle3Accent4',
  'LightStyle3Accent5',
  'LightStyle3Accent6',
  'MediumStyle1',
  'MediumStyle1Accent1',
  'MediumStyle1Accent2',
  'MediumStyle1Accent3',
  'MediumStyle1Accent4',
  'MediumStyle1Accent5',
  'MediumStyle1Accent6',
  'MediumStyle2',
  'MediumStyle2Accent1',
  'MediumStyle2Accent2',
  'MediumStyle2Accent3',
  'MediumStyle2Accent4',
  'MediumStyle2Accent5',
  'MediumStyle2Accent6',
  'MediumStyle3',
  'MediumStyle3Accent1',
  'MediumStyle3Accent2',
  'MediumStyle3Accent3',
  'MediumStyle3Accent4',
  'MediumStyle3Accent5',
  'MediumStyle3Accent6',
  'MediumStyle4',
  'MediumStyle4Accent1',
  'MediumStyle4Accent2',
  'MediumStyle4Accent3',
  'MediumStyle4Accent4',
  'MediumStyle4Accent5',
  'MediumStyle4Accent6',
  'DarkStyle1',
  'DarkStyle1Accent1',
  'DarkStyle1Accent2',
  'DarkStyle1Accent3',
  'DarkStyle1Accent4',
  'DarkStyle1Accent5',
  'DarkStyle1Accent6',
  'DarkStyle2',
  'DarkStyle2Accent1',
  'DarkStyle2Accent2',
  'DarkStyle2Accent3',
] as const;

export type OfficeJsKey = (typeof OFFICE_JS_TABLE_STYLES)[number];

export interface RosterName {
  readonly family: string;
  /** `plain`, `accent1`..`accent6`, or Dark Style 2's `pair1`..`pair3`. */
  readonly variant: string;
  /** The name the gallery shows, as the mapping below predicts it. */
  readonly ui: string;
}

/** The UI name an Office.js key stands for; the run holds this mapping to the gallery. */
export function rosterName(key: OfficeJsKey): RosterName {
  if (key === 'NoStyleNoGrid')
    return { family: 'No Style', variant: 'noGrid', ui: 'No Style, No Grid' };
  if (key === 'NoStyleTableGrid') {
    return { family: 'No Style', variant: 'defaultGrid', ui: 'No Style, Table Grid' };
  }
  const match = /^(Themed|Light|Medium|Dark)Style(\d)(?:Accent(\d))?$/.exec(key);
  if (match === null) throw new Error(`no mapping for ${key}`);
  const [, kind = '', n = '', accent] = match;
  const family = `${kind} Style ${n}`;
  if (accent === undefined) return { family, variant: 'plain', ui: family };
  const a = Number(accent);
  if (family === 'Dark Style 2') {
    return {
      family,
      variant: `pair${accent}`,
      ui: `${family} - Accent ${String(2 * a - 1)}/Accent ${String(2 * a)}`,
    };
  }
  return { family, variant: `accent${accent}`, ui: `${family} - Accent ${accent}` };
}

/* -------------------------------------------------------------------------- */
/* what a probe is                                                            */
/* -------------------------------------------------------------------------- */

/** Whole-table fills no built-in style can produce: every definition uses theme colours only. */
export const SIGNATURE = {
  A: 'C8A001',
  B: 'C8B002',
  C: 'C8C003',
  D: 'C8D004',
  E: 'C8E005',
  F: 'C8F006',
} as const;

/** Custom GUIDs, upper case as `ST_Guid` requires; the last digits say which signature they carry. */
export const CUSTOM = {
  A: '{C8000000-0000-4000-8000-00000000A001}',
  A2: '{C8000000-0000-4000-8000-00000000A002}',
  C: '{C8000000-0000-4000-8000-00000000C003}',
  F: '{C8000000-0000-4000-8000-00000000F006}',
  B: '{C8000000-0000-4000-8000-00000000B002}',
  M1: '{C8000000-0000-4000-8000-0000000000E1}',
  M2: '{C8000000-0000-4000-8000-0000000000E2}',
  M3: '{C8000000-0000-4000-8000-0000000000E3}',
  P1: '{C8000000-0000-4000-8000-0000000000D1}',
  P2: '{C8000000-0000-4000-8000-0000000000D2}',
} as const;

/** The three built-ins the resolution probes use, and the stock `@def`, which is `G1`. */
export const REFERENCE = {
  G1: 'Medium Style 2 - Accent 1',
  G2: 'Dark Style 1',
  G3: 'Medium Style 3 - Accent 2',
} as const;

/** Questions whose drawing cannot separate their candidates, so each probe's result is recorded. */
export const RECORD_ONLY: ReadonlySet<string> = new Set([
  'custom-banded',
  'lexical-custom',
  'location',
  'form',
]);

export type Theme = 's1' | 's2' | 's3';

/** A palette whose `dk1` is no system colour and is far from black, so the grid's colour shows. */
export const SCHEME_THREE = {
  dk1: '6B3A10',
  lt1: 'FFF4DC',
  dk2: '1F3A5F',
  lt2: 'E4ECF5',
  accent1: '9C27B0',
  accent2: '00897B',
  accent3: 'F4511E',
  accent4: '3949AB',
  accent5: 'C0CA33',
  accent6: '6D4C41',
  hlink: '1565C0',
  folHlink: '6A1B9A',
} as const;

/**
 * What a slide shows, as a candidate predicts it: a named control's pixels, all sixteen cell
 * centres one exact colour, or nothing but the background.
 */
export type Observation = `control:${string}` | `signature:${string}` | 'bare';

export interface SlideSpec {
  /** The whole `a:tblPr`; `null` writes none. */
  readonly tblPr: string | null;
  /** Every cell's `a:tcPr`; defaults to an empty one. */
  readonly tcPr?: string;
  /** Each candidate of the probe's question, and what it predicts here. */
  readonly predict: Readonly<Record<string, Observation>>;
}

export interface StylesPartSpec {
  readonly xml: string;
  readonly partName?: string;
  readonly rel?: boolean;
  readonly part?: boolean;
}

export interface Probe {
  readonly id: string;
  readonly question: string;
  readonly theme: Theme;
  /** A control's pixels are what other slides are compared against; it is asked nothing. */
  readonly control?: true;
  /** Schema-invalid or a duplicate; expected to be repaired or refused, and alone in its package. */
  readonly hostile?: true;
  readonly slides: readonly SlideSpec[];
  readonly part: StylesPartSpec | null;
}

/* -------------------------------------------------------------------------- */
/* markup                                                                     */
/* -------------------------------------------------------------------------- */

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** Four 108-pt columns by four 36-pt rows at one inch in, every cell empty. */
export function tableFrame(name: string, tblPr: string | null, tcPr = '<a:tcPr/>'): string {
  const cell = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1200"/></a:p></a:txBody>${tcPr}</a:tc>`;
  const row = `<a:tr h="${String(ROW)}">${cell.repeat(4)}</a:tr>`;
  const grid = `<a:gridCol w="${String(COL)}"/>`.repeat(4);
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="2" name="${name}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
    '</p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${String(FRAME.x)}" y="${String(FRAME.y)}"/>` +
    `<a:ext cx="${String(4 * COL)}" cy="${String(4 * ROW)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${TABLE_URI}">` +
    `<a:tbl>${tblPr ?? ''}<a:tblGrid>${grid}</a:tblGrid>${row.repeat(4)}</a:tbl>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

/** An `a:tblPr` naming `id` (or nothing), with the flags written verbatim. */
export function tblPr(id: string | null, flags = ''): string {
  const ref = id === null ? '' : `<a:tableStyleId>${id}</a:tableStyleId>`;
  return ref === '' ? `<a:tblPr${flags}/>` : `<a:tblPr${flags}>${ref}</a:tblPr>`;
}

/** A whole `ppt/tableStyles.xml`. */
export function styleList(def: string | null, ...styles: readonly string[]): string {
  const attr = def === null ? '' : ` def="${def}"`;
  const body = styles.join('');
  return (
    DECLARATION +
    (body === ''
      ? `<a:tblStyleLst xmlns:a="${NS_A}"${attr}/>`
      : `<a:tblStyleLst xmlns:a="${NS_A}"${attr}>${body}</a:tblStyleLst>`)
  );
}

/** A style whose whole table is one flat fill and nothing else, or `wholeTbl` written verbatim. */
export function flatStyle(
  id: string | null,
  name: string,
  fill: string,
  wholeTbl?: string,
): string {
  const attr = id === null ? '' : ` styleId="${id}"`;
  const whole =
    wholeTbl ??
    `<a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>`;
  return `<a:tblStyle${attr} styleName="${name}">${whole}</a:tblStyle>`;
}

/** The same style under another id. */
export function renamed(xml: string, id: string): string {
  const next = xml.replace(/^<a:tblStyle styleId="[^"]+"/, `<a:tblStyle styleId="${id}"`);
  if (next === xml) throw new Error('no styleId to rename');
  return next;
}

const WHOLE = /<a:wholeTbl>(.*?)<\/a:wholeTbl>/;

/** PowerPoint's own style with only its whole-table fill replaced by `fill`. */
export function withWholeFill(xml: string, fill: string): string {
  const whole = WHOLE.exec(xml)?.[1];
  if (whole === undefined) throw new Error('no a:wholeTbl');
  const flat = `<a:fill><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:fill>`;
  const next = /<a:fill>.*?<\/a:fill>/.test(whole)
    ? whole.replace(/<a:fill>.*?<\/a:fill>/, flat)
    : whole.replace('</a:tcBdr>', `</a:tcBdr>${flat}`);
  if (next === whole) throw new Error('no place for a whole-table fill');
  return xml.replace(WHOLE, `<a:wholeTbl>${next}</a:wholeTbl>`);
}

/** PowerPoint's own style with its whole `a:wholeTbl` replaced by one flat fill. */
export function withWholeReplaced(xml: string, fill: string): string {
  const flat = `<a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>`;
  if (!WHOLE.test(xml)) throw new Error('no a:wholeTbl');
  return xml.replace(WHOLE, flat);
}

/* -------------------------------------------------------------------------- */
/* what PowerPoint authored                                                   */
/* -------------------------------------------------------------------------- */

export interface GalleryItem {
  readonly seq: number;
  readonly uiaName: string;
  readonly styleId: string;
}

export interface AuthorLog {
  readonly gallery: readonly GalleryItem[];
}

/** The table every probe draws, in points, and the export that samples it: one pixel per point. */
export const TABLE_RECT = { x: 72, y: 72, w: 432, h: 144 } as const;
export const EXPORT = { width: 960, height: 540 } as const;

const utf8 = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

/** Every `a:tblStyle` in a part with a `@styleId`, in document order, duplicates kept. */
export function stylesIn(part: string): [id: string, xml: string][] {
  return [...part.matchAll(/<a:tblStyle styleId="([^"]+)"[^>]*>.*?<\/a:tblStyle>/g)].map((m) => [
    m[1] ?? '',
    m[0],
  ]);
}

/** Every `a:tblStyle` in a part, keyed by `@styleId`, as the exact bytes PowerPoint wrote. */
export function definitionsIn(part: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const [id, xml] of stylesIn(part)) {
    if (out.has(id)) throw new Error(`${id} is defined twice`);
    out.set(id, xml);
  }
  return out;
}

/** The table-style part of a deck PowerPoint saved, found through the presentation's relationship. */
export function tableStylesOf(deck: Uint8Array): string | null {
  const entries = readZip(deck);
  const rels = entry(entries, 'ppt/_rels/presentation.xml.rels');
  if (rels === undefined) throw new Error('no ppt/_rels/presentation.xml.rels');
  const target = /<Relationship [^>]*Type="[^"]*\/tableStyles"[^>]*Target="([^"]+)"/.exec(
    utf8(rels),
  )?.[1];
  if (target === undefined) return null;
  const part = entry(entries, `ppt/${target}`);
  if (part === undefined)
    throw new Error(`the tableStyles relationship names ppt/${target}, which is missing`);
  return utf8(part);
}

/** The gallery's styles, each as PowerPoint serialised it in `part`. */
export function catalogueFrom(log: AuthorLog, part: string): Catalogue {
  const written = definitionsIn(part);
  const definitions = new Map<string, string>();
  const ids = new Map<string, string>();
  for (const item of log.gallery) {
    const xml = written.get(item.styleId);
    if (xml === undefined)
      throw new Error(`PowerPoint did not write ${item.uiaName} (${item.styleId})`);
    definitions.set(item.styleId, xml);
    ids.set(item.uiaName, item.styleId);
  }
  return { definitions, ids };
}

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

export interface Catalogue {
  /** GUID to PowerPoint's own serialisation, in gallery order. */
  readonly definitions: ReadonlyMap<string, string>;
  /** Gallery name to GUID. */
  readonly ids: ReadonlyMap<string, string>;
}

const one = (tblPr: string | null, predict: Record<string, Observation>): SlideSpec[] => [
  { tblPr, predict },
];

/** A GUID with one hex digit changed, so it is certainly nobody's. */
function offByOne(id: string): string {
  const last = id.at(-2);
  if (last === undefined) throw new Error(`not a GUID: ${id}`);
  return `${id.slice(0, -2)}${last === '0' ? '1' : '0'}}`;
}

export function allProbes(catalogue: Catalogue): Probe[] {
  const idOf = (name: string): string => {
    const id = catalogue.ids.get(name);
    if (id === undefined) throw new Error(`the gallery has no "${name}"`);
    return id;
  };
  const V = (id: string): string => {
    const xml = catalogue.definitions.get(id);
    if (xml === undefined) throw new Error(`PowerPoint wrote no definition of ${id}`);
    return xml;
  };
  const G1 = idOf(REFERENCE.G1);
  const G2 = idOf(REFERENCE.G2);
  const G3 = idOf(REFERENCE.G3);
  const G = { G1, G2, G3 } as const;
  const all = [...catalogue.definitions.keys()];
  const g1Slide = all.indexOf(G1);
  if (g1Slide < 0) throw new Error('G1 is not in the catalogue');
  const gridSlide = all.indexOf(idOf('No Style, Table Grid'));
  const grid: Observation = `control:sweep-ctl#${String(gridSlide)}`;
  const grid2: Observation = `control:sweep2-ctl#${String(gridSlide)}`;
  const UNKNOWN = [
    '{00000000-0000-0000-0000-00000000C700}',
    '{C8000000-0000-4000-8000-00000000FFFF}',
    offByOne(G2),
  ];
  const off = (id: string | null): string => tblPr(id);
  const banded = (id: string | null): string => tblPr(id, ' bandRow="1"');
  const X = CUSTOM;
  const S = SIGNATURE;
  const probes: Probe[] = [];
  const add = (probe: Probe): void => {
    probes.push(probe);
  };

  // ---- controls -----------------------------------------------------------------------------
  for (const [k, id] of Object.entries(G)) {
    add({
      id: `ctl-${k}`,
      question: 'control',
      theme: 's1',
      control: true,
      slides: one(off(id), {}),
      part: { xml: styleList(id, V(id)) },
    });
    add({
      id: `ctl2-${k}`,
      question: 'control',
      theme: 's2',
      control: true,
      slides: one(off(id), {}),
      part: { xml: styleList(id, V(id)) },
    });
    const merge = { G1: X.M1, G2: X.M2, G3: X.M3 }[k] ?? '';
    add({
      id: `ctl-merge-${k}-B`,
      question: 'custom-id',
      theme: 's1',
      slides: one(off(merge), {
        package: `signature:${S.B}`,
        bare: 'bare',
        def: `signature:${S.B}`,
        fixedG1: 'control:ctl-G1',
        defaultGrid: grid,
      }),
      part: { xml: styleList(merge, renamed(withWholeFill(V(id), S.B), merge)) },
    });
  }
  add({
    id: 'ctl-sig-B',
    question: 'custom-id',
    theme: 's1',
    slides: one(off(X.B), {
      package: `signature:${S.B}`,
      bare: 'bare',
      def: `signature:${S.B}`,
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(X.B, flatStyle(X.B, 'C8 Flat B', S.B)) },
  });
  for (const [k, id] of [
    ['G1', G1],
    ['G2', G2],
  ] as const) {
    const merge = k === 'G1' ? X.M1 : X.M2;
    const part = k === 'G1' ? X.P1 : X.P2;
    add({
      id: `ctl-${k}-banded`,
      question: 'control',
      theme: 's1',
      control: true,
      slides: one(banded(id), {}),
      part: { xml: styleList(id, V(id)) },
    });
    add({
      id: `ctl-merge-${k}-B-banded`,
      question: 'custom-banded',
      theme: 's1',
      slides: one(banded(merge), { defaultGrid: grid, bare: 'bare' }),
      part: { xml: styleList(merge, renamed(withWholeFill(V(id), S.B), merge)) },
    });
    add({
      id: `ctl-partmerge-${k}-B-banded`,
      question: 'custom-banded',
      theme: 's1',
      slides: one(banded(part), { defaultGrid: grid, bare: 'bare' }),
      part: { xml: styleList(part, renamed(withWholeReplaced(V(id), S.B), part)) },
    });
  }
  const sweepSlides = (predict: (k: number) => Record<string, Observation>): SlideSpec[] =>
    all.map((id, k) => ({ tblPr: tblPr(id, ' firstRow="1" bandRow="1"'), predict: predict(k) }));
  const every = all.map(V);
  add({
    id: 'sweep-ctl',
    question: 'control',
    theme: 's1',
    control: true,
    slides: sweepSlides(() => ({})),
    part: { xml: styleList(G1, ...every) },
  });
  add({
    id: 'sweep2-ctl',
    question: 'control',
    theme: 's2',
    control: true,
    slides: sweepSlides(() => ({})),
    part: { xml: styleList(G1, ...every) },
  });

  // ---- a built-in id the package does not define --------------------------------------------
  const UNDEF = 'undefined-builtin';
  for (const [k, id] of Object.entries(G)) {
    add({
      id: `undef-nopart-${k}`,
      question: UNDEF,
      theme: 's1',
      slides: one(off(id), {
        builtin: `control:ctl-${k}`,
        bare: 'bare',
        def: 'bare',
        fixedG1: 'control:ctl-G1',
        defaultGrid: grid,
      }),
      part: null,
    });
    add({
      id: `undef-nopart2-${k}`,
      question: UNDEF,
      theme: 's2',
      slides: one(off(id), {
        builtin: `control:ctl2-${k}`,
        bare: 'bare',
        def: 'bare',
        fixedG1: 'control:ctl2-G1',
        defaultGrid: grid2,
      }),
      part: null,
    });
  }
  for (const [k, id, other, ok] of [
    ['G2', G2, G3, 'G3'],
    ['G3', G3, G2, 'G2'],
  ] as const) {
    add({
      id: `undef-empty-${k}`,
      question: UNDEF,
      theme: 's1',
      slides: one(off(id), {
        builtin: `control:ctl-${k}`,
        bare: 'bare',
        def: `control:ctl-${ok}`,
        fixedG1: 'control:ctl-G1',
        defaultGrid: grid,
      }),
      part: { xml: styleList(other) },
    });
  }
  add({
    id: 'sweep-nopart',
    question: UNDEF,
    theme: 's1',
    slides: sweepSlides((k) => ({
      builtin: `control:sweep-ctl#${String(k)}`,
      bare: 'bare',
      def: 'bare',
      fixedG1: `control:sweep-ctl#${String(g1Slide)}`,
      defaultGrid: grid,
    })),
    part: null,
  });
  add({
    id: 'sweep-empty',
    question: UNDEF,
    theme: 's1',
    slides: sweepSlides((k) => ({
      builtin: `control:sweep-ctl#${String(k)}`,
      bare: 'bare',
      def: `control:sweep-ctl#${String(g1Slide)}`,
      fixedG1: `control:sweep-ctl#${String(g1Slide)}`,
      defaultGrid: grid,
    })),
    part: { xml: styleList(G1) },
  });
  add({
    id: 'sweep2-nopart',
    question: UNDEF,
    theme: 's2',
    slides: sweepSlides((k) => ({
      builtin: `control:sweep2-ctl#${String(k)}`,
      bare: 'bare',
      def: 'bare',
      fixedG1: `control:sweep2-ctl#${String(g1Slide)}`,
      defaultGrid: grid2,
    })),
    part: null,
  });

  // ---- a built-in id the package defines differently ----------------------------------------
  const MOD = 'modified-builtin';
  for (const [k, id] of Object.entries(G)) {
    add({
      id: `mod-${k}`,
      question: MOD,
      theme: 's1',
      slides: one(off(id), {
        package: `signature:${S.B}`,
        builtin: `control:ctl-${k}`,
        defaultGrid: grid,
      }),
      part: { xml: styleList(id, flatStyle(id, REFERENCE[k as keyof typeof REFERENCE], S.B)) },
    });
  }
  add({
    id: 'mod-renamed-G1',
    question: MOD,
    theme: 's1',
    slides: one(off(G1), {
      package: `signature:${S.B}`,
      builtin: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G1, flatStyle(G1, 'C8 Modified', S.B)) },
  });
  add({
    id: 'mod-subtle-G1',
    question: MOD,
    theme: 's1',
    slides: one(off(G1), {
      package: `signature:${S.B}`,
      builtin: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G1, withWholeFill(V(G1), S.B)) },
  });
  for (const [k, id] of [
    ['G1', G1],
    ['G2', G2],
  ] as const) {
    add({
      id: `mod-banded-${k}`,
      question: MOD,
      theme: 's1',
      slides: one(banded(id), {
        package: `signature:${S.B}`,
        builtin: `control:ctl-${k}-banded`,
        defaultGrid: grid,
      }),
      part: { xml: styleList(id, flatStyle(id, REFERENCE[k], S.B)) },
    });
  }

  // ---- a custom id, and whether it is looked up by id or by name ----------------------------
  const CUSTOM_Q = 'custom-id';
  add({
    id: 'custom-X',
    question: CUSTOM_Q,
    theme: 's1',
    slides: one(off(X.A), {
      package: `signature:${S.A}`,
      bare: 'bare',
      def: 'control:ctl-G2',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G2, V(G2), flatStyle(X.A, 'C8 Custom', S.A)) },
  });
  add({
    id: 'custom-X2',
    question: CUSTOM_Q,
    theme: 's1',
    slides: one(off(X.A2), {
      package: `signature:${S.A}`,
      bare: 'bare',
      def: 'control:ctl-G3',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G3, V(G3), flatStyle(X.A2, 'C8 Custom Two', S.A)) },
  });
  add({
    id: 'custom-name-X',
    question: 'lookup',
    theme: 's1',
    slides: one(off(X.F), {
      byId: `signature:${S.F}`,
      byName: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G2, V(G2), flatStyle(X.F, REFERENCE.G1, S.F)) },
  });

  // ---- an id nothing defines, and no id at all ----------------------------------------------
  const UNKNOWN_Q = 'unknown-id';
  UNKNOWN.forEach((u, i) => {
    const n = String(i + 1);
    add({
      id: `unknown-nopart-U${n}`,
      question: UNKNOWN_Q,
      theme: 's1',
      slides: one(off(u), {
        bare: 'bare',
        def: 'bare',
        fixedG1: 'control:ctl-G1',
        defaultGrid: grid,
      }),
      part: null,
    });
    add({
      id: `unknown-empty-U${n}`,
      question: UNKNOWN_Q,
      theme: 's1',
      slides: one(off(u), {
        bare: 'bare',
        def: 'control:ctl-G2',
        fixedG1: 'control:ctl-G1',
        defaultGrid: grid,
      }),
      part: { xml: styleList(G2) },
    });
  });
  add({
    id: 'unknown-defC',
    question: UNKNOWN_Q,
    theme: 's1',
    slides: one(off(UNKNOWN[0] ?? ''), {
      bare: 'bare',
      def: `signature:${S.C}`,
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(X.C, flatStyle(X.C, 'C8 Default', S.C)) },
  });
  const NOID = 'no-id';
  add({
    id: 'noid-nopart',
    question: NOID,
    theme: 's1',
    slides: one(off(null), {
      bare: 'bare',
      def: 'bare',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: null,
  });
  add({
    id: 'noid-defG2',
    question: NOID,
    theme: 's1',
    slides: one(off(null), {
      bare: 'bare',
      def: 'control:ctl-G2',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G2, V(G2)) },
  });
  add({
    id: 'noid-defC',
    question: NOID,
    theme: 's1',
    slides: one(off(null), {
      bare: 'bare',
      def: `signature:${S.C}`,
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(X.C, flatStyle(X.C, 'C8 Default', S.C)) },
  });
  add({
    id: 'noid-no-tblpr',
    question: NOID,
    theme: 's1',
    slides: one(null, {
      bare: 'bare',
      def: 'control:ctl-G2',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G2, V(G2)) },
  });

  // ---- how a GUID is compared ----------------------------------------------------------------
  const CASE = 'lexical-case';
  const BRACES = 'lexical-braces';
  const SPACE = 'lexical-space';
  const unmatched = {
    bare: 'bare',
    def: 'control:ctl-G3',
    fixedG1: 'control:ctl-G1',
    defaultGrid: grid,
  } as const;
  const lexParts = { xml: styleList(G3, V(G3)) };
  const mixed = G2.slice(0, 10).toLowerCase() + G2.slice(10);
  add({
    id: 'lex-lower-G2',
    question: CASE,
    theme: 's1',
    hostile: true,
    slides: one(off(G2.toLowerCase()), { matched: 'control:ctl-G2', ...unmatched }),
    part: lexParts,
  });
  add({
    id: 'lex-mixed-G2',
    question: CASE,
    theme: 's1',
    hostile: true,
    slides: one(off(mixed), { matched: 'control:ctl-G2', ...unmatched }),
    part: lexParts,
  });
  add({
    id: 'lex-nobrace-G2',
    question: BRACES,
    theme: 's1',
    hostile: true,
    slides: one(off(G2.slice(1, -1)), { matched: 'control:ctl-G2', ...unmatched }),
    part: lexParts,
  });
  add({
    id: 'lex-space-G2',
    question: SPACE,
    theme: 's1',
    hostile: true,
    slides: one(off(` ${G2} `), { matched: 'control:ctl-G2', ...unmatched }),
    part: lexParts,
  });
  add({
    id: 'lex-newline-G2',
    question: SPACE,
    theme: 's1',
    hostile: true,
    slides: one(off(`\n${G2}\n`), { matched: 'control:ctl-G2', ...unmatched }),
    part: lexParts,
  });
  const customMatched = { matched: `signature:${S.A}`, defaultGrid: grid } as const;
  add({
    id: 'lex-lower-ref-X',
    question: 'lexical-custom',
    theme: 's1',
    hostile: true,
    slides: one(off(X.A.toLowerCase()), customMatched),
    part: { xml: styleList(G3, V(G3), flatStyle(X.A, 'C8 Custom', S.A)) },
  });
  add({
    id: 'lex-lower-def-X',
    question: 'lexical-custom',
    theme: 's1',
    hostile: true,
    slides: one(off(X.A), customMatched),
    part: { xml: styleList(G3, V(G3), flatStyle(X.A.toLowerCase(), 'C8 Custom', S.A)) },
  });

  add({
    id: 'lex-lower-G3',
    question: CASE,
    theme: 's1',
    hostile: true,
    slides: one(off(G3.toLowerCase()), {
      matched: 'control:ctl-G3',
      bare: 'bare',
      def: 'control:ctl-G2',
      fixedG1: 'control:ctl-G1',
      defaultGrid: grid,
    }),
    part: { xml: styleList(G2, V(G2)) },
  });

  // ---- the grid a style that does not resolve is drawn as, in the second theme ---------------
  const FALLBACK = 'fallback-theme';
  const black = (edge: string): string =>
    `<a:${edge} w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:${edge}>`;
  add({
    id: 'ctl2-black-grid',
    question: 'control',
    theme: 's2',
    control: true,
    slides: [
      {
        tblPr: off(null),
        tcPr: `<a:tcPr>${['lnL', 'lnR', 'lnT', 'lnB'].map(black).join('')}</a:tcPr>`,
        predict: {},
      },
    ],
    part: null,
  });
  const fallback = {
    tableGrid: grid2,
    blackGrid: 'control:ctl2-black-grid',
    bare: 'bare',
  } as const;
  add({
    id: 'unknown-nopart2-U1',
    question: FALLBACK,
    theme: 's2',
    slides: one(off(UNKNOWN[0] ?? ''), fallback),
    part: null,
  });
  add({
    id: 'noid-nopart2',
    question: FALLBACK,
    theme: 's2',
    slides: one(off(null), fallback),
    part: null,
  });
  add({
    id: 'custom-X-s2',
    question: FALLBACK,
    theme: 's2',
    slides: one(off(X.A), fallback),
    part: { xml: styleList(X.A, flatStyle(X.A, 'C8 Custom', S.A)) },
  });

  // ---- the grid's colour: black, the theme's dk1, or tx1 through the clrMap --------------------
  const COLOUR = 'fallback-colour';
  const edges = (colour: string): string =>
    `<a:tcPr>${['lnL', 'lnR', 'lnT', 'lnB']
      .map((edge) => `<a:${edge} w="12700"><a:solidFill>${colour}</a:solidFill></a:${edge}>`)
      .join('')}</a:tcPr>`;
  for (const [name, colour] of [
    ['black', '<a:srgbClr val="000000"/>'],
    ['dk1', '<a:schemeClr val="dk1"/>'],
    ['tx1', '<a:schemeClr val="tx1"/>'],
  ] as const) {
    add({
      id: `ctl3-${name}-grid`,
      question: 'control',
      theme: 's3',
      control: true,
      slides: [{ tblPr: off(null), tcPr: edges(colour), predict: {} }],
      part: null,
    });
  }
  const colour = {
    black: 'control:ctl3-black-grid',
    dk1: 'control:ctl3-dk1-grid',
    tx1: 'control:ctl3-tx1-grid',
    bare: 'bare',
  } as const;
  add({
    id: 'unknown-nopart3-U1',
    question: COLOUR,
    theme: 's3',
    slides: one(off(UNKNOWN[0] ?? ''), colour),
    part: null,
  });
  add({
    id: 'noid-nopart3',
    question: COLOUR,
    theme: 's3',
    slides: one(off(null), colour),
    part: null,
  });
  add({
    id: 'custom-X-s3',
    question: COLOUR,
    theme: 's3',
    slides: one(off(X.A), colour),
    part: { xml: styleList(X.A, flatStyle(X.A, 'C8 Custom', S.A)) },
  });

  // ---- one GUID defined twice ----------------------------------------------------------------
  const DUP = 'duplicate';
  add({
    id: 'dup-X',
    question: DUP,
    theme: 's1',
    hostile: true,
    slides: one(off(X.A), { first: `signature:${S.D}`, last: `signature:${S.E}`, ignored: grid }),
    part: {
      xml: styleList(G2, V(G2), flatStyle(X.A, 'C8 First', S.D), flatStyle(X.A, 'C8 Last', S.E)),
    },
  });
  add({
    id: 'dup-G1',
    question: DUP,
    theme: 's1',
    hostile: true,
    slides: one(off(G1), {
      first: `signature:${S.D}`,
      last: `signature:${S.E}`,
      ignored: 'control:ctl-G1',
    }),
    part: { xml: styleList(G1, flatStyle(G1, 'C8 First', S.D), flatStyle(G1, 'C8 Last', S.E)) },
  });

  // ---- a style written inline -----------------------------------------------------------------
  const INLINE = 'inline';
  const inline = (id: string, fill: string, ref = ''): string =>
    `<a:tblPr>${ref}${flatStyle(id, 'C8 Inline', fill)
      .replace(/^<a:tblStyle/, '<a:tableStyle')
      .replace(/<\/a:tblStyle>$/, '</a:tableStyle>')}</a:tblPr>`;
  add({
    id: 'inline-X',
    question: INLINE,
    theme: 's1',
    slides: one(inline(X.A, S.A), {
      content: `signature:${S.A}`,
      part: grid,
      byStyleId: grid,
      ignored: grid,
    }),
    part: null,
  });
  add({
    id: 'inline-G1-mod',
    question: INLINE,
    theme: 's1',
    slides: one(inline(G1, S.B), {
      content: `signature:${S.B}`,
      part: grid,
      byStyleId: 'control:ctl-G1',
      ignored: grid,
    }),
    part: null,
  });
  add({
    id: 'inline-G2-mod',
    question: INLINE,
    theme: 's1',
    slides: one(inline(G2, S.B), {
      content: `signature:${S.B}`,
      part: grid,
      byStyleId: 'control:ctl-G2',
      ignored: grid,
    }),
    part: null,
  });
  add({
    id: 'inline-vs-part',
    question: INLINE,
    theme: 's1',
    slides: one(inline(X.A, S.A), {
      content: `signature:${S.A}`,
      part: `signature:${S.C}`,
      byStyleId: grid,
      ignored: grid,
    }),
    part: { xml: styleList(X.A, flatStyle(X.A, 'C8 Part', S.C)) },
  });
  add({
    id: 'inline-and-id',
    question: INLINE,
    theme: 's1',
    hostile: true,
    slides: one(inline(X.A, S.A, `<a:tableStyleId>${G2}</a:tableStyleId>`), {
      content: `signature:${S.A}`,
      part: grid,
      byStyleId: grid,
      ignored: 'control:ctl-G2',
    }),
    part: null,
  });

  // ---- where the part is found ---------------------------------------------------------------
  const WHERE = 'location';
  add({
    id: 'rel-renamed',
    question: WHERE,
    theme: 's1',
    slides: one(off(X.A), { byRelationship: `signature:${S.A}`, byPath: grid }),
    part: {
      xml: styleList(X.A, flatStyle(X.A, 'C8 Elsewhere', S.A)),
      partName: '/ppt/custom/ts.xml',
    },
  });
  add({
    id: 'orphan-part',
    question: WHERE,
    theme: 's1',
    slides: one(off(X.A), { byRelationship: grid, byPath: `signature:${S.A}` }),
    part: { xml: styleList(X.A, flatStyle(X.A, 'C8 Orphan', S.A)), rel: false },
  });
  add({
    id: 'rel-missing-target',
    question: WHERE,
    theme: 's1',
    hostile: true,
    slides: one(off(X.A), { byRelationship: grid, byPath: `signature:${S.A}` }),
    part: { xml: styleList(X.A), part: false },
  });
  add({
    id: 'wrong-root',
    question: WHERE,
    theme: 's1',
    hostile: true,
    slides: one(off(X.A), { byRelationship: grid, byPath: `signature:${S.A}` }),
    part: {
      xml:
        DECLARATION +
        flatStyle(X.A, 'C8 Root', S.A).replace('<a:tblStyle ', `<a:tblStyle xmlns:a="${NS_A}" `),
    },
  });

  // ---- attributes and elements the schema requires or forbids ---------------------------------
  const FORM = 'form';
  const formProbe = (id: string, part: string): void =>
    add({
      id,
      question: FORM,
      theme: 's1',
      hostile: true,
      slides: one(off(X.A), { applied: `signature:${S.A}`, defaultGrid: grid }),
      part: { xml: part },
    });
  formProbe('def-missing', styleList(null, flatStyle(X.A, 'C8 No Def', S.A)));
  formProbe(
    'b-yes',
    styleList(
      X.A,
      flatStyle(
        X.A,
        'C8 Bold Yes',
        S.A,
        `<a:wholeTbl><a:tcTxStyle b="yes"/><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="${S.A}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>`,
      ),
    ),
  );
  formProbe(
    'edge-empty',
    styleList(
      X.A,
      flatStyle(
        X.A,
        'C8 Empty Edge',
        S.A,
        `<a:wholeTbl><a:tcStyle><a:tcBdr><a:left/></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="${S.A}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>`,
      ),
    ),
  );
  formProbe(
    'fill-empty',
    styleList(
      X.A,
      `<a:tblStyle styleId="${X.A}" styleName="C8 Empty Fill"><a:tblBg><a:fill/></a:tblBg><a:wholeTbl><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="${S.A}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>`,
    ),
  );
  formProbe('styleid-missing', styleList(X.A, flatStyle(null, 'C8 No Id', S.A)));

  const ids = new Set<string>();
  for (const probe of probes) {
    if (ids.has(probe.id)) throw new Error(`duplicate probe id ${probe.id}`);
    ids.add(probe.id);
  }
  return probes;
}

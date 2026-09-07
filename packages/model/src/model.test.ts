/**
 * The model, against what PowerPoint actually did.
 *
 * The fixture is imported and the rules are **re-derived from it**, not compared
 * against a summary of it. Every match case carries the placeholders that were
 * on the slide and on the layout and the box the shape landed in; the test feeds
 * those same placeholders through this package's own matcher and checks the
 * prediction. So a matcher that drifts fails on 34 measurements rather than on
 * one sentence somebody wrote about them.
 */

import { PartStore, REL_TYPE } from '@pptx-studio/opc';
import { resolveColor, toHexColor, type Rgba } from '@pptx-studio/paint';
import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/sheets.json' with { type: 'json' };
import pictures from '../../../corpus/ground-truth/pictures.json' with { type: 'json' };

import { resolveBackground, resolveBackgroundColor } from './resolve/background.js';
import { loadDocument } from './document.js';
import { ModelError } from './errors.js';
import {
  DEFAULT_PLACEHOLDER_IDX,
  DEFAULT_PLACEHOLDER_TYPE,
  inheritanceChain,
  masterPlaceholderType,
  matchInLayout,
  matchInMaster,
  normalizePlaceholder,
} from './resolve/placeholder.js';
import { parseSheet, parseTheme } from './parse/sheet.js';
import { colorMapOf, resolve, resolveXfrm, sheetChain, themeOf } from './resolve/resolve.js';
import {
  STYLE_MATRIX_OFFSET,
  resolveAppearance,
  resolveSolidFill,
  styleMatrixTarget,
} from './style.js';
import type { Placeholder, PlaceholderType, Shape, Sheet } from './types.js';

/* -------------------------------------------------------------------------- */
/* building sheets by hand                                                    */
/* -------------------------------------------------------------------------- */

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SP_TREE_HEAD =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';

const CLR_MAP =
  '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"/>';

let nextId = 100;

interface PhSpec {
  readonly type?: string | undefined;
  readonly idx?: number | undefined;
  readonly box?: string | undefined;
  readonly fill?: string | undefined;
  readonly style?: string | undefined;
  readonly line?: string | undefined;
  readonly name?: string | undefined;
}

const BOXES = fixture.boxes as Readonly<
  Record<string, { x: number; y: number; w: number; h: number }>
>;

function phXml(spec: PhSpec): string {
  const attrs =
    (spec.type === undefined ? '' : ` type="${spec.type}"`) +
    (spec.idx === undefined ? '' : ` idx="${String(spec.idx)}"`);
  const box = spec.box === undefined ? undefined : BOXES[spec.box];
  const xfrm =
    box === undefined
      ? ''
      : `<a:xfrm><a:off x="${String(box.x * 12700)}" y="${String(box.y * 12700)}"/>` +
        `<a:ext cx="${String(box.w * 12700)}" cy="${String(box.h * 12700)}"/></a:xfrm>`;
  const name = spec.name ?? `${spec.type ?? 'ph'}#${String(spec.idx ?? 0)}`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${String(nextId++)}" name="${name}"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spec.fill ?? ''}${spec.line ?? ''}</p:spPr>` +
    `${spec.style ?? ''}<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`
  );
}

interface SheetXmlOptions {
  readonly shapes?: readonly string[] | undefined;
  readonly bg?: string | undefined;
  readonly clrMapOvr?: string | undefined;
  readonly showMasterSp?: boolean | undefined;
  readonly type?: string | undefined;
}

function sheetXml(kind: 'slide' | 'layout' | 'master', options: SheetXmlOptions = {}): string {
  const tag = kind === 'slide' ? 'p:sld' : kind === 'layout' ? 'p:sldLayout' : 'p:sldMaster';
  const attrs =
    (options.showMasterSp === false ? ' showMasterSp="0"' : '') +
    (kind === 'layout' ? ` type="${options.type ?? 'obj'}"` : '');
  const tail =
    kind === 'master'
      ? CLR_MAP
      : `<p:clrMapOvr>${options.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr>`;
  return (
    `<${tag} ${NS}${attrs}><p:cSld name="${kind}">${options.bg ?? ''}${SP_TREE_HEAD}` +
    `${(options.shapes ?? []).join('')}</p:spTree></p:cSld>${tail}</${tag}>`
  );
}

function parse(kind: 'slide' | 'layout' | 'master', options: SheetXmlOptions = {}) {
  const document = parseXmlString(sheetXml(kind, options));
  return parseSheet(document.root, `/ppt/${kind}.xml`);
}

/** The theme the C5 decks used, rebuilt from the scheme the fixture records. */
const SCHEME = fixture.schemes.one as Readonly<Record<string, string>>;

function themeXml(scheme: Readonly<Record<string, string>> = SCHEME): string {
  const slot = (k: string): string =>
    `<a:${k}><a:srgbClr val="${scheme[k] ?? '000000'}"/></a:${k}>`;
  const slots = [
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
  ]
    .map(slot)
    .join('');
  const line = (w: number): string =>
    `<a:ln w="${String(w)}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;
  const matrix = fixture.themeStyleMatrix;
  return (
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">` +
    '<a:themeElements>' +
    `<a:clrScheme name="T">${slots}</a:clrScheme>` +
    '<a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
    '<a:fmtScheme name="T">' +
    `<a:fillStyleLst>${matrix.fillStyleLst.join('')}</a:fillStyleLst>` +
    `<a:lnStyleLst>${matrix.lnStyleWidths.map(line).join('')}</a:lnStyleLst>` +
    '<a:effectStyleLst>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3) +
    '</a:effectStyleLst>' +
    `<a:bgFillStyleLst>${matrix.bgFillStyleLst.join('')}</a:bgFillStyleLst>` +
    '</a:fmtScheme></a:themeElements></a:theme>'
  );
}

function theme(scheme: Readonly<Record<string, string>> = SCHEME) {
  return parseTheme(parseXmlString(themeXml(scheme)).root, '/ppt/theme/theme1.xml');
}

/** A three-link chain, wired the way `document.ts` wires one. */
function chain(
  slideOptions: SheetXmlOptions,
  layoutOptions: SheetXmlOptions,
  masterOptions: SheetXmlOptions = {},
  scheme: Readonly<Record<string, string>> = SCHEME,
): { slide: Sheet; layout: Sheet; master: Sheet } {
  const master: Sheet = { ...parse('master', masterOptions), parent: null, theme: theme(scheme) };
  const layout: Sheet = { ...parse('layout', layoutOptions), parent: master, theme: null };
  const slide: Sheet = { ...parse('slide', slideOptions), parent: layout, theme: null };
  return { slide, layout, master };
}

function onlyShape(sheet: Sheet): Shape {
  expect(sheet.shapes.length).toBeGreaterThan(0);
  return sheet.shapes[0]!;
}

/** The box a resolved `xfrm` sits in, named the way the fixture names them. */
function boxOf(sheet: Sheet, shape: Shape): string {
  const xfrm = resolveXfrm(shape, sheet);
  if (xfrm === undefined) return 'orphan';
  const points = {
    x: xfrm.value.x / 12700,
    y: xfrm.value.y / 12700,
    w: xfrm.value.cx / 12700,
    h: xfrm.value.cy / 12700,
  };
  for (const [name, b] of Object.entries(BOXES)) {
    if (b.x === points.x && b.y === points.y && b.w === points.w && b.h === points.h) return name;
  }
  return `rect(${String(points.x)},${String(points.y)},${String(points.w)},${String(points.h)})`;
}

const hex = (rgba: Rgba | null): string | null => (rgba === null ? null : toHexColor(rgba));

/* -------------------------------------------------------------------------- */

describe('the first hop: slide to layout', () => {
  const cases = fixture.findings.firstHopCases;

  it('has cases to check', () => {
    expect(cases.length).toBe(34);
  });

  it.each(cases.map((c) => [c.id, c] as const))(
    'reproduces %s',
    (_id, probe: (typeof cases)[number]) => {
      const layout = parse('layout', {
        shapes: probe.layout.map((ph) =>
          phXml({ type: ph.type, idx: ph.idx, box: ph.box, name: ph.box }),
        ),
      });
      const slide = parse('slide', {
        shapes: [phXml({ type: probe.slide.type, idx: probe.slide.idx, name: 'probe' })],
      });
      const layoutSheet: Sheet = { ...layout, parent: null, theme: null };
      const slideSheet: Sheet = { ...slide, parent: layoutSheet, theme: null };

      expect(boxOf(slideSheet, onlyShape(slideSheet))).toBe(probe.landed);
    },
  );

  it('matches on idx and not on the pair, scored over every case', () => {
    // The same scoring the analysis ran, repeated here so the claim is a test
    // rather than a note: a `(type, idx)` matcher is wrong on 21 of 34.
    let idxOnly = 0;
    let typeAndIdx = 0;
    for (const probe of cases) {
      const byIdx = probe.layout.find((ph) => ph.idx === probe.slide.idx);
      const byPair = probe.layout.find(
        (ph) => ph.idx === probe.slide.idx && ph.type === probe.slide.type,
      );
      if ((byIdx?.box ?? 'orphan') === probe.landed) idxOnly += 1;
      if ((byPair?.box ?? 'orphan') === probe.landed) typeAndIdx += 1;
    }
    expect(idxOnly).toBe(cases.length);
    expect(typeAndIdx).toBeLessThan(cases.length);
    expect(fixture.findings.slideToLayoutKey).toBe('idx');
  });
});

describe('the second hop: layout to master', () => {
  const cases = fixture.findings.secondHopCases;

  it.each(cases.map((c) => [c.id, c] as const))(
    'reproduces %s',
    (_id, probe: (typeof cases)[number]) => {
      const master = parse('master', {
        shapes: probe.master.map((ph) =>
          phXml({ type: ph.type, idx: ph.idx, box: ph.box, name: ph.box }),
        ),
      });
      const layout = parse('layout', {
        shapes: [phXml({ type: probe.layout.type, idx: probe.layout.idx, name: 'layoutPh' })],
      });
      const masterSheet: Sheet = { ...master, parent: null, theme: null };
      const layoutSheet: Sheet = { ...layout, parent: masterSheet, theme: null };
      expect(boxOf(layoutSheet, onlyShape(layoutSheet))).toBe(probe.landed);
    },
  );

  it('folds every content type to body and ctrTitle to title', () => {
    const folded = fixture.findings.layoutToMasterNormalisation as Readonly<Record<string, string>>;
    for (const [from, to] of Object.entries(folded)) {
      expect(masterPlaceholderType(from as PlaceholderType)).toBe(to);
    }
    // The five a master may carry are their own targets.
    for (const type of ['title', 'body', 'dt', 'ftr', 'sldNum', 'hdr'] as PlaceholderType[]) {
      expect(masterPlaceholderType(type)).toBe(type);
    }
  });

  it('takes the first placeholder of the type, not the one whose idx agrees', () => {
    expect(fixture.findings.layoutToMasterTie).toBe('first');
    const master = parse('master', {
      shapes: [
        phXml({ type: 'body', idx: 1, box: 'M1', name: 'M1' }),
        phXml({ type: 'body', idx: 5, box: 'M2', name: 'M2' }),
      ],
    });
    const layout = parse('layout', { shapes: [phXml({ type: 'body', idx: 5 })] });
    const masterSheet: Sheet = { ...master, parent: null, theme: null };
    const layoutSheet: Sheet = { ...layout, parent: masterSheet, theme: null };
    expect(boxOf(layoutSheet, onlyShape(layoutSheet))).toBe('M1');
  });

  it('asks with the placeholder of the sheet it is leaving', () => {
    // A slide `obj` at idx 0 reaching a layout `ctrTitle` lands on the master's
    // *title*. If the second hop used the slide's own type it would fold to
    // `body` and land on M1 instead. Measured as `h3-ctrTitle`.
    const master = parse('master', {
      shapes: [
        phXml({ type: 'title', box: 'M0', name: 'M0' }),
        phXml({ type: 'body', idx: 1, box: 'M1', name: 'M1' }),
      ],
    });
    const layout = parse('layout', { shapes: [phXml({ type: 'ctrTitle' })] });
    const slide = parse('slide', { shapes: [phXml({ idx: 0 })] });
    const masterSheet: Sheet = { ...master, parent: null, theme: null };
    const layoutSheet: Sheet = { ...layout, parent: masterSheet, theme: null };
    const slideSheet: Sheet = { ...slide, parent: layoutSheet, theme: null };
    expect(boxOf(slideSheet, onlyShape(slideSheet))).toBe('M0');
    expect(inheritanceChain(onlyShape(slideSheet), slideSheet).map((l) => l.origin)).toEqual([
      'shape',
      'layoutPh',
      'masterPh',
    ]);
  });
});

describe('placeholder defaults', () => {
  it('a bare p:ph is obj at idx 0', () => {
    expect(fixture.findings.barePlaceholder.typeReadBack).toBe('obj');
    expect(DEFAULT_PLACEHOLDER_TYPE).toBe('obj');
    expect(DEFAULT_PLACEHOLDER_IDX).toBe(0);
    const sheet = parse('slide', { shapes: [phXml({})] });
    const ph = sheet.shapes[0]!.placeholder as Placeholder;
    expect(ph.type).toBeNull();
    expect(ph.idx).toBeNull();
    expect(normalizePlaceholder(ph)).toEqual({ type: 'obj', idx: 0 });
  });

  it('keeps what the file said, so a round trip can put it back', () => {
    const sheet = parse('slide', {
      shapes: [phXml({ type: 'body', idx: 3 })],
    });
    const ph = sheet.shapes[0]!.placeholder as Placeholder;
    expect(ph.type).toBe('body');
    expect(ph.idx).toBe(3);
  });

  it('records @sz and @orient without matching on them', () => {
    const document = parseXmlString(
      sheetXml('layout', {
        shapes: [
          '<p:sp><p:nvSpPr><p:cNvPr id="9" name="dt"/><p:cNvSpPr/>' +
            '<p:nvPr><p:ph type="dt" sz="half" orient="vert" idx="10"/></p:nvPr></p:nvSpPr>' +
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>',
        ],
      }),
    );
    const sheet = parseSheet(document.root, '/ppt/layout.xml');
    const ph = sheet.shapes[0]!.placeholder as Placeholder;
    expect(ph.size).toBe('half');
    expect(ph.orient).toBe('vert');
    // A slide placeholder with neither attribute still matches it, on idx.
    const layoutSheet: Sheet = { ...sheet, parent: null, theme: null };
    const slide = parse('slide', { shapes: [phXml({ type: 'dt', idx: 10 })] });
    expect(matchInLayout(slide.shapes[0]!.placeholder!, layoutSheet)?.name).toBe('dt');
  });

  it('refuses an idx that is not an unsignedInt', () => {
    for (const idx of ['-1', '4294967296', 'x']) {
      expect(() =>
        parseSheet(
          parseXmlString(
            sheetXml('slide', {
              shapes: [
                `<p:sp><p:nvSpPr><p:cNvPr id="9" name="p"/><p:cNvSpPr/>` +
                  `<p:nvPr><p:ph type="body" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`,
              ],
            }),
          ).root,
          '/ppt/slide.xml',
        ),
      ).toThrow(ModelError);
    }
  });
});

describe('orphans', () => {
  it('a placeholder that matched nothing resolves to no geometry at all', () => {
    expect(fixture.findings.orphanRect).toEqual({ x: 0, y: 0, w: 0, h: 0 });
    const layout = parse('layout', { shapes: [phXml({ type: 'body', idx: 1, box: 'A1' })] });
    const slide = parse('slide', { shapes: [phXml({ type: 'title', idx: 9 })] });
    const layoutSheet: Sheet = { ...layout, parent: null, theme: null };
    const slideSheet: Sheet = { ...slide, parent: layoutSheet, theme: null };
    expect(resolveXfrm(onlyShape(slideSheet), slideSheet)).toBeUndefined();
    expect(matchInLayout(onlyShape(slideSheet).placeholder!, layoutSheet)).toBeNull();
  });

  it('a layout with no placeholders orphans everything', () => {
    const layoutSheet: Sheet = { ...parse('layout'), parent: null, theme: null };
    const slide = parse('slide', { shapes: [phXml({ type: 'title' })] });
    const slideSheet: Sheet = { ...slide, parent: layoutSheet, theme: null };
    expect(inheritanceChain(onlyShape(slideSheet), slideSheet)).toHaveLength(1);
  });
});

describe('what descends besides geometry', () => {
  const green = '<a:solidFill><a:srgbClr val="00A000"/></a:solidFill>';
  const red = '<a:solidFill><a:srgbClr val="A00000"/></a:solidFill>';
  const styleXml = (fill: number, ln: number, colour: string): string =>
    '<p:style>' +
    `<a:lnRef idx="${String(ln)}"><a:schemeClr val="${colour}"/></a:lnRef>` +
    `<a:fillRef idx="${String(fill)}"><a:schemeClr val="${colour}"/></a:fillRef>` +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

  it('a fill comes down from the layout placeholder', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 1 })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A1', fill: green })] },
    );
    const appearance = resolveAppearance(onlyShape(slide), slide);
    expect(appearance.fillOrigin).toBe('layoutPh');
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.descends.fillFromLayout,
    );
  });

  it('an explicit fill on the slide beats it', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 1, fill: red })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A1', fill: green })] },
    );
    expect(resolveAppearance(onlyShape(slide), slide).fillOrigin).toBe('shape');
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.descends.fillOverridden,
    );
  });

  it('a p:style comes down, and its fillRef resolves against the theme', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 2 })] },
      { shapes: [phXml({ type: 'body', idx: 2, box: 'A2', style: styleXml(1, 2, 'accent2') })] },
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.descends.styleFromLayout.fill,
    );
    const line = resolveAppearance(onlyShape(slide), slide).line;
    expect(line?.w).toBe(fixture.findings.descends.styleFromLayout.weight * 12700);
  });

  it('both reach a slide from the master through a layout that declares nothing', () => {
    const { slide } = chain(
      { shapes: [phXml({ idx: 1 })] },
      { shapes: [phXml({ type: 'body', idx: 1 })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'M1', style: styleXml(2, 2, 'accent2') })] },
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.descends.styleFromMasterThroughEmptyLayout.fill,
    );
    expect(resolveAppearance(onlyShape(slide), slide).line?.w).toBe(
      fixture.findings.descends.styleFromMasterThroughEmptyLayout.weight * 12700,
    );
  });

  it('a shape with no fill, no style and no placeholder paints nothing', () => {
    const { slide } = chain(
      {
        shapes: [
          '<p:sp><p:nvSpPr><p:cNvPr id="9" name="plain"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
            '<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>',
        ],
      },
      {},
    );
    const appearance = resolveAppearance(onlyShape(slide), slide);
    expect(appearance.fill).toBeNull();
    expect(appearance.line).toBeNull();
    expect(fixture.findings.descends.nothingAnywhere).toEqual({ filled: false, lined: false });
  });

  it('an explicit a:noFill is not the same as no fill element', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 1, fill: '<a:noFill/>' })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A1', fill: green })] },
    );
    expect(resolveAppearance(onlyShape(slide), slide).fill).toEqual({ type: 'none' });
  });
});

describe('the style matrix', () => {
  const styleXml = (fill: number, colour: string): string =>
    '<p:style>' +
    `<a:lnRef idx="1"><a:schemeClr val="${colour}"/></a:lnRef>` +
    `<a:fillRef idx="${String(fill)}"><a:schemeClr val="${colour}"/></a:fillRef>` +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

  function fillAt(idx: number, colour = 'accent1'): string | null {
    const { slide } = chain({ shapes: [phXml({ style: styleXml(idx, colour) })] }, {});
    return hex(resolveSolidFill(onlyShape(slide), slide));
  }

  it('indexes both lists with the same 1000 offset', () => {
    expect(STYLE_MATRIX_OFFSET).toBe(1000);
    expect(styleMatrixTarget(0)).toEqual({ list: 'none' });
    expect(styleMatrixTarget(1000)).toEqual({ list: 'none' });
    expect(styleMatrixTarget(1)).toEqual({ list: 'main', at: 0 });
    expect(styleMatrixTarget(1001)).toEqual({ list: 'background', at: 0 });
  });

  it.each(fixture.findings.styleMatrix.fillStyleLst.map((e) => [e.idx, e.hex] as const))(
    'fillRef %i paints %s',
    (idx, want) => {
      expect(fillAt(idx)).toBe(want);
    },
  );

  it.each(fixture.findings.styleMatrix.bgFillStyleLst.map((e) => [e.idx, e.hex] as const))(
    'fillRef %i paints %s',
    (idx, want) => {
      expect(fillAt(idx)).toBe(want);
    },
  );

  it('the line list is one-based and its widths are the theme', () => {
    const widths = fixture.themeStyleMatrix.lnStyleWidths;
    fixture.findings.styleMatrix.lnStyleLst.forEach((entry, i) => {
      expect(entry.weight * 12700).toBe(widths[i]);
    });
  });

  it('idx 0 is no fill and no line at all', () => {
    expect(fixture.findings.styleMatrix.zeroIsNone).toEqual({ fill: true, line: true });
    const { slide } = chain(
      {
        shapes: [
          phXml({
            style:
              '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>' +
              '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
              '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
              '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
          }),
        ],
      },
      {},
    );
    const appearance = resolveAppearance(onlyShape(slide), slide);
    expect(appearance.fill).toBeNull();
    expect(appearance.line).toBeNull();
  });

  it('clamps an out-of-range index rather than throwing', () => {
    expect(fixture.findings.styleMatrix.outOfRangeClampsToLast).toBe(true);
    expect(fillAt(4)).toBe(fixture.findings.styleMatrix.fillStyleLst[2]!.hex);
  });

  it('phClr is the colour the reference carries, transforms and all', () => {
    // The same entry, invoked with three different colours.
    expect(fillAt(1, 'accent2')).toBe(SCHEME['accent2']);
    expect(fillAt(1, 'accent6')).toBe(SCHEME['accent6']);
    // And a literal, which is what a `fillRef` with an `srgbClr` child means.
    const { slide } = chain(
      {
        shapes: [
          phXml({
            style:
              '<p:style><a:lnRef idx="1"><a:srgbClr val="FF0000"/></a:lnRef>' +
              '<a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef>' +
              '<a:effectRef idx="0"><a:srgbClr val="FF0000"/></a:effectRef>' +
              '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
          }),
        ],
      },
      {},
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe('FF0000');
  });

  it('loses to an explicit spPr fill, and to an explicit noFill', () => {
    // Mutation P: making the p:style win unconditionally killed nothing until
    // these two landed. Both are measured - s-override and s-nofill-override.
    const withFill = chain(
      {
        shapes: [
          phXml({
            fill: '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>',
            style: styleXml(1, 'accent1'),
          }),
        ],
      },
      {},
    );
    expect(hex(resolveSolidFill(onlyShape(withFill.slide), withFill.slide))).toBe('123456');
    expect(resolveAppearance(onlyShape(withFill.slide), withFill.slide).fillOrigin).toBe('shape');

    const withNoFill = chain(
      { shapes: [phXml({ fill: '<a:noFill/>', style: styleXml(1, 'accent1') })] },
      {},
    );
    expect(resolveAppearance(onlyShape(withNoFill.slide), withNoFill.slide).fill).toEqual({
      type: 'none',
    });

    // And an inherited spPr fill beats the style too, since the spPr route is
    // searched down the whole chain before the style route is consulted at all.
    const inherited = chain(
      { shapes: [phXml({ type: 'body', idx: 1, style: styleXml(1, 'accent1') })] },
      {
        shapes: [
          phXml({
            type: 'body',
            idx: 1,
            box: 'A1',
            fill: '<a:solidFill><a:srgbClr val="00A000"/></a:solidFill>',
          }),
        ],
      },
    );
    expect(hex(resolveSolidFill(onlyShape(inherited.slide), inherited.slide))).toBe('00A000');
  });

  it('refuses a numeric fontRef idx, which PowerPoint repairs', () => {
    expect(() =>
      parse('slide', {
        shapes: [
          phXml({
            style:
              '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>' +
              '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
              '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
              '<a:fontRef idx="1"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
          }),
        ],
      }),
    ).toThrow(ModelError);
  });
});

describe('the background', () => {
  const solid = (h: string): string =>
    `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${h}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
  const ref = (idx: number, colour = 'accent1'): string =>
    `<p:bg><p:bgRef idx="${String(idx)}"><a:schemeClr val="${colour}"/></p:bgRef></p:bg>`;

  it.each(fixture.findings.background.byIndex.map((e) => [e.idx, e.hex] as const))(
    'p:bgRef %i paints %s',
    (idx, want) => {
      const { slide } = chain({}, {}, { bg: ref(idx) });
      expect(hex(resolveBackgroundColor(slide))).toBe(want);
    },
  );

  it('reaches the same six entries a:fillRef reaches', () => {
    expect(fixture.findings.background.sameTableAsFillRef).toBe(true);
    const viaBg = fixture.findings.background.byIndex.map((e) => e.hex);
    const viaFill = [
      ...fixture.findings.styleMatrix.fillStyleLst.map((e) => e.hex),
      ...fixture.findings.styleMatrix.bgFillStyleLst.map((e) => e.hex),
    ];
    expect(viaBg).toEqual(viaFill);
  });

  it('idx 0 and idx 1000 paint nothing', () => {
    for (const idx of [0, 1000]) {
      const { slide } = chain({}, {}, { bg: ref(idx) });
      expect(resolveBackground(slide)?.fill).toEqual({ type: 'none' });
    }
  });

  it('takes the nearest sheet that declares one', () => {
    const precedence = fixture.findings.background.precedence;
    const withLayoutBg = chain({}, { bg: solid('00FF00') }, { bg: solid('FF0000') });
    expect(hex(resolveBackgroundColor(withLayoutBg.slide))).toBe(precedence.layoutBeatsMaster);
    expect(resolveBackground(withLayoutBg.slide)?.sheet.kind).toBe('layout');

    const fromMaster = chain({}, {}, { bg: solid('FF0000') });
    expect(hex(resolveBackgroundColor(fromMaster.slide))).toBe(precedence.masterWhenNeither);

    const own = chain({ bg: solid('0000FF') }, { bg: solid('00FF00') }, { bg: solid('FF0000') });
    expect(hex(resolveBackgroundColor(own.slide))).toBe(precedence.slideBeatsLayout);
    expect(resolveBackground(own.slide)?.explicit).toBe(true);
  });

  it('honours an explicit noFill rather than inheriting through it', () => {
    expect(fixture.findings.background.precedence.slideNoFillIsHonoured).toBe(true);
    const { slide } = chain(
      { bg: '<p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg>' },
      { bg: solid('00FF00') },
      { bg: solid('FF0000') },
    );
    expect(resolveBackground(slide)?.fill).toEqual({ type: 'none' });
  });
});

describe('the colour map', () => {
  const override = (accent1: string, bg1: string): string =>
    `<a:overrideClrMapping bg1="${bg1}" tx1="lt1" bg2="dk2" tx2="lt2" accent1="${accent1}"` +
    ' accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
    ' hlink="hlink" folHlink="folHlink"/>';
  const swatch = (name: string): string =>
    phXml({ box: 'A0', fill: `<a:solidFill><a:schemeClr val="${name}"/></a:solidFill>` });

  it("a layout's override reaches a slide that says masterClrMapping", () => {
    const { slide } = chain(
      { shapes: [swatch('accent1')] },
      { clrMapOvr: override('accent6', 'dk1') },
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.colourMap.layoutOverrideReachesSlide,
    );
  });

  it("a slide's own override beats its layout's", () => {
    const { slide } = chain(
      { shapes: [swatch('accent1')], clrMapOvr: override('accent4', 'lt2') },
      { clrMapOvr: override('accent6', 'dk1') },
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.colourMap.slideOverrideBeatsLayout,
    );
  });

  it('dk1 bypasses the map, as it bypasses the master map', () => {
    const { slide } = chain(
      { shapes: [swatch('dk1')], clrMapOvr: override('accent4', 'lt2') },
      { clrMapOvr: override('accent6', 'dk1') },
    );
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(
      fixture.findings.colourMap.unmappedNamesBypass,
    );
  });

  it('falls through to the master map when nobody overrides', () => {
    const { slide } = chain({ shapes: [swatch('accent1')] }, {});
    expect(hex(resolveSolidFill(onlyShape(slide), slide))).toBe(fixture.findings.colourMap.control);
    expect(colorMapOf(slide)?.accent1).toBe('accent1');
  });

  it("a bgRef colour uses the asking sheet's map, not the declaring sheet's", () => {
    const { slide } = chain(
      { clrMapOvr: override('accent4', 'lt2') },
      { clrMapOvr: override('accent6', 'dk1') },
      { bg: '<p:bg><p:bgRef idx="1"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' },
    );
    // The theme's first fill entry is bare `phClr`, so the background is exactly
    // whatever `bg1` resolved to - `lt2` under the slide's own override.
    expect(hex(resolveBackgroundColor(slide))).toBe(
      fixture.findings.colourMap.backgroundUsesSheetMap,
    );
  });
});

describe('the chain comes from the rels', () => {
  const CT = 'application/vnd.openxmlformats-officedocument';
  const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

  function twoMasterPackage(): PartStore {
    const store = PartStore.create();
    const add = (name: string, type: string, xml: string): void => {
      store.addPart(name, type, utf8(xml));
    };

    add(
      '/ppt/presentation.xml',
      `${CT}.presentationml.presentation.main+xml`,
      `<p:presentation ${NS}><p:sldMasterIdLst>` +
        '<p:sldMasterId id="2147483648" r:id="rId1"/><p:sldMasterId id="2147483650" r:id="rId2"/>' +
        '</p:sldMasterIdLst><p:sldIdLst>' +
        '<p:sldId id="256" r:id="rId3"/><p:sldId id="257" r:id="rId4"/>' +
        '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>' +
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
    );
    add('/ppt/theme/theme1.xml', `${CT}.theme+xml`, themeXml(SCHEME));
    add('/ppt/theme/theme2.xml', `${CT}.theme+xml`, themeXml(fixture.schemes.two));
    for (const n of [1, 2]) {
      add(
        `/ppt/slideMasters/slideMaster${String(n)}.xml`,
        `${CT}.presentationml.slideMaster+xml`,
        sheetXml('master', {
          bg: '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
        }),
      );
      add(
        `/ppt/slideLayouts/slideLayout${String(n)}.xml`,
        `${CT}.presentationml.slideLayout+xml`,
        sheetXml('layout', { type: 'blank' }),
      );
      add(
        `/ppt/slides/slide${String(n)}.xml`,
        `${CT}.presentationml.slide+xml`,
        sheetXml('slide', {
          shapes: [
            phXml({
              box: 'A0',
              fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
            }),
          ],
        }),
      );
    }

    const root = store.rootRelationships();
    root.addWithId('rId1', REL_TYPE.officeDocument, 'ppt/presentation.xml');

    const pres = store.relationships('/ppt/presentation.xml');
    pres.addWithId('rId1', REL_TYPE.slideMaster, 'slideMasters/slideMaster1.xml');
    pres.addWithId('rId2', REL_TYPE.slideMaster, 'slideMasters/slideMaster2.xml');
    pres.addWithId('rId3', REL_TYPE.slide, 'slides/slide1.xml');
    pres.addWithId('rId4', REL_TYPE.slide, 'slides/slide2.xml');
    // The trap: PowerPoint writes this, it always names theme1, and reading it
    // gives every slide the first master's palette.
    pres.addWithId('rId5', REL_TYPE.theme, 'theme/theme1.xml');

    for (const n of [1, 2]) {
      const master = store.relationships(`/ppt/slideMasters/slideMaster${String(n)}.xml`);
      master.addWithId('rId1', REL_TYPE.slideLayout, `../slideLayouts/slideLayout${String(n)}.xml`);
      master.addWithId('rId2', REL_TYPE.theme, `../theme/theme${String(n)}.xml`);
      const layout = store.relationships(`/ppt/slideLayouts/slideLayout${String(n)}.xml`);
      layout.addWithId('rId1', REL_TYPE.slideMaster, `../slideMasters/slideMaster${String(n)}.xml`);
      const slide = store.relationships(`/ppt/slides/slide${String(n)}.xml`);
      slide.addWithId('rId1', REL_TYPE.slideLayout, `../slideLayouts/slideLayout${String(n)}.xml`);
    }
    return store;
  }

  it('gives each master its own theme, not the one the presentation names', () => {
    const document = loadDocument(twoMasterPackage());
    expect(document.problems).toEqual([]);
    expect(document.slides).toHaveLength(2);
    expect(document.masters).toHaveLength(2);
    expect(document.layouts).toHaveLength(2);

    const [first, second] = document.slides;
    expect(themeOf(first!)?.partName).toBe('/ppt/theme/theme1.xml');
    expect(themeOf(second!)?.partName).toBe('/ppt/theme/theme2.xml');

    expect(hex(resolveSolidFill(onlyShape(first!), first!))).toBe(
      fixture.findings.themes.perMaster.oneExpected,
    );
    expect(hex(resolveSolidFill(onlyShape(second!), second!))).toBe(
      fixture.findings.themes.perMaster.twoExpected,
    );
    expect(fixture.findings.themes.presentationRelIsNotTheAnswer).toBe(true);
  });

  it('reads the slide order from p:sldIdLst and the binding from the rels', () => {
    const document = loadDocument(twoMasterPackage());
    expect(document.slides.map((s) => s.partName)).toEqual([
      '/ppt/slides/slide1.xml',
      '/ppt/slides/slide2.xml',
    ]);
    expect(sheetChain(document.slides[1]!).map((s) => s.partName)).toEqual([
      '/ppt/slides/slide2.xml',
      '/ppt/slideLayouts/slideLayout2.xml',
      '/ppt/slideMasters/slideMaster2.xml',
    ]);
    expect(document.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
  });

  it('crosses the numbering and still follows the rels', () => {
    // Master 1 owns slideLayout2 and master 2 owns slideLayout1, and each slide
    // is bound to the layout of the *other* number. Nothing about the part names
    // lines up any more, which is the point: PowerPoint's own two-master decks
    // put every layout in one flat folder and only the rels say who owns what.
    const store = twoMasterPackage();
    for (const n of [1, 2]) {
      const other = n === 1 ? 2 : 1;
      const master = store.relationships(`/ppt/slideMasters/slideMaster${String(n)}.xml`);
      master.remove('rId1');
      master.addWithId(
        'rId1',
        REL_TYPE.slideLayout,
        `../slideLayouts/slideLayout${String(other)}.xml`,
      );
      const layout = store.relationships(`/ppt/slideLayouts/slideLayout${String(n)}.xml`);
      layout.remove('rId1');
      layout.addWithId(
        'rId1',
        REL_TYPE.slideMaster,
        `../slideMasters/slideMaster${String(other)}.xml`,
      );
      const slide = store.relationships(`/ppt/slides/slide${String(n)}.xml`);
      slide.remove('rId1');
      slide.addWithId(
        'rId1',
        REL_TYPE.slideLayout,
        `../slideLayouts/slideLayout${String(other)}.xml`,
      );
    }

    const document = loadDocument(store);
    expect(document.problems).toEqual([]);
    expect(sheetChain(document.slides[0]!).map((sheet) => sheet.partName)).toEqual([
      '/ppt/slides/slide1.xml',
      '/ppt/slideLayouts/slideLayout2.xml',
      '/ppt/slideMasters/slideMaster1.xml',
    ]);
    // And so the palette follows the master, not the layout's number.
    expect(themeOf(document.slides[0]!)?.partName).toBe('/ppt/theme/theme1.xml');
    expect(themeOf(document.slides[1]!)?.partName).toBe('/ppt/theme/theme2.xml');
    expect(hex(resolveSolidFill(onlyShape(document.slides[1]!), document.slides[1]!))).toBe(
      fixture.findings.themes.perMaster.twoExpected,
    );
  });

  it('records a missing binding as a problem rather than throwing', () => {
    const store = twoMasterPackage();
    store.relationships('/ppt/slides/slide1.xml').remove('rId1');
    const document = loadDocument(store);
    expect(document.problems.map((p) => p.code)).toContain('MODEL_NO_PARENT');
    expect(document.slides[0]!.parent).toBeNull();
    // And a slide with no layout still resolves everything it stated itself.
    expect(resolveXfrm(onlyShape(document.slides[0]!), document.slides[0]!)).toBeDefined();
  });

  it('records a dangling layout target rather than throwing', () => {
    const store = twoMasterPackage();
    const rels = store.relationships('/ppt/slides/slide1.xml');
    rels.remove('rId1');
    rels.addWithId('rId1', REL_TYPE.slideLayout, '../slideLayouts/slideLayout9.xml');
    const document = loadDocument(store);
    expect(document.problems.map((p) => p.code)).toContain('MODEL_PART_MISSING');
    expect(document.slides[0]!.parent).toBeNull();
  });

  it('records a master with no theme rather than throwing', () => {
    const store = twoMasterPackage();
    store.relationships('/ppt/slideMasters/slideMaster1.xml').remove('rId2');
    const document = loadDocument(store);
    expect(document.problems.map((p) => p.code)).toContain('MODEL_NO_THEME');
    expect(themeOf(document.slides[0]!)).toBeNull();
  });

  it('refuses to follow a layout whose master relationship names a layout', () => {
    const store = twoMasterPackage();
    const rels = store.relationships('/ppt/slideLayouts/slideLayout1.xml');
    rels.remove('rId1');
    rels.addWithId('rId1', REL_TYPE.slideMaster, '../slideLayouts/slideLayout2.xml');
    const document = loadDocument(store);
    expect(document.problems.map((p) => p.code)).toContain('MODEL_PART_KIND');
    expect(document.slides[0]!.parent?.parent).toBeNull();
  });

  it('throws only when the package names no presentation at all', () => {
    const store = PartStore.create();
    expect(() => loadDocument(store)).toThrow(ModelError);
  });
});

describe('the master placeholder vocabulary', () => {
  it('is the six PowerPoint accepts', () => {
    const vocabulary = fixture.findings.masterPlaceholderVocabulary;
    expect(vocabulary.accepted).toEqual(['title', 'body', 'dt', 'ftr', 'sldNum', 'hdr']);
    expect(vocabulary.hdrAccepted).toBe(true);
    for (const entry of vocabulary.refused) expect(entry.repaired).toBe(true);
  });

  it('is why the second hop folds', () => {
    // Every type a master may not carry has to fold to one it may, or the stock
    // layouts inherit nothing.
    for (const entry of fixture.findings.masterPlaceholderVocabulary.refused) {
      const folded = masterPlaceholderType(entry.type as PlaceholderType);
      expect(fixture.findings.masterPlaceholderVocabulary.accepted).toContain(folded);
    }
  });

  it('finds nothing on a master that carries only what it may not', () => {
    const master = parse('master', { shapes: [phXml({ type: 'title', box: 'M0' })] });
    const masterSheet: Sheet = { ...master, parent: null, theme: null };
    const layout = parse('layout', { shapes: [phXml({ type: 'body', idx: 1 })] });
    expect(matchInMaster(layout.shapes[0]!.placeholder!, masterSheet)).toBeNull();
  });
});

describe('parsing', () => {
  it('keeps an absent xfrm absent', () => {
    const sheet = parse('slide', { shapes: [phXml({ type: 'title' })] });
    expect(sheet.shapes[0]!.xfrm).toBeUndefined();
  });

  /**
   * Both halves of `a:path` are optional and both absences carry a meaning that
   * a default would destroy: no `@path` paints the box ramp, and no
   * `a:fillToRect` puts the focus at the centre - which four zero insets do not.
   * Measured in ADR 0022; the parser is the only place that can keep them apart.
   */
  it.each([
    ['<a:path><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>', 'rect', 'rect'],
    ['<a:path path="circle"/>', 'circle', null],
    ['<a:path path="rect"/>', 'rect', null],
    ['<a:path/>', 'rect', null],
    ['<a:path path="circle"><a:fillToRect/></a:path>', 'circle', 'rect'],
  ])('parses %s as path=%s', (shade, kind, rect) => {
    const sheet = parse('slide', {
      shapes: [
        '<p:sp><p:nvSpPr><p:cNvPr id="9" name="g"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr><a:gradFill><a:gsLst>' +
          '<a:gs pos="0"><a:srgbClr val="000000"/></a:gs>' +
          '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>' +
          `</a:gsLst>${shade}</a:gradFill></p:spPr></p:sp>`,
      ],
    });
    const fill = sheet.shapes[0]!.fill;
    if (fill?.type !== 'gradient') throw new Error('expected a gradient fill');
    const found = fill.shade;
    if (found?.kind !== 'path') throw new Error('expected a path shade');
    expect(found.path).toBe(kind);
    if (rect === null) expect(found.fillToRect).toBeNull();
    else expect(found.fillToRect).not.toBeNull();
  });

  it('reads a graphicFrame geometry from p:xfrm, not from p:spPr', () => {
    const sheet = parse('slide', {
      shapes: [
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="frame"/>' +
          '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
          '<p:xfrm><a:off x="127000" y="254000"/><a:ext cx="381000" cy="508000"/></p:xfrm>' +
          '<a:graphic><a:graphicData uri="x"/></a:graphic></p:graphicFrame>',
      ],
    });
    expect(sheet.shapes[0]!.kind).toBe('graphicFrame');
    expect(sheet.shapes[0]!.xfrm).toMatchObject({ x: 127000, y: 254000, cx: 381000, cy: 508000 });
  });

  it('reads a group and its child coordinate space', () => {
    const sheet = parse('slide', {
      shapes: [
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
          '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/>' +
          '<a:chOff x="10" y="20"/><a:chExt cx="50" cy="60"/></a:xfrm></p:grpSpPr>' +
          '<p:sp><p:nvSpPr><p:cNvPr id="7" name="kid"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
          '<p:spPr/></p:sp></p:grpSp>',
      ],
    });
    const group = sheet.shapes[0]!;
    expect(group.kind).toBe('grpSp');
    expect(group.xfrm?.child).toEqual({ x: 10, y: 20, cx: 50, cy: 60 });
    expect(group.children).toHaveLength(1);
    expect(group.children[0]!.name).toBe('kid');
  });

  it('refuses a clrMap that is missing one of its twelve', () => {
    const xml = sheetXml('master').replace(' folHlink="folHlink"', '');
    expect(() => parseSheet(parseXmlString(xml).root, '/ppt/master.xml')).toThrow(ModelError);
  });

  it('refuses a clrMap naming something that is not a slot', () => {
    const xml = sheetXml('master').replace('accent1="accent1"', 'accent1="bg1"');
    expect(() => parseSheet(parseXmlString(xml).root, '/ppt/master.xml')).toThrow(ModelError);
  });

  it('reads showMasterSp off the sheet element, not off p:cSld', () => {
    expect(parse('slide', { showMasterSp: false }).showMasterShapes).toBe(false);
    expect(parse('slide').showMasterShapes).toBe(true);
  });

  it('distinguishes masterClrMapping from an override', () => {
    expect(parse('slide').clrMapOvr).toEqual({ kind: 'inherit' });
    expect(parse('master').clrMapOvr).toBeUndefined();
    expect(parse('master').clrMap?.bg1).toBe('lt1');
  });

  it('parses an effect list into paint effects', () => {
    const sheet = parse('slide', {
      shapes: [
        '<p:sp><p:nvSpPr><p:cNvPr id="8" name="e"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
          '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000">' +
          '<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>' +
          '<a:glow rad="63500"><a:schemeClr val="accent1"/></a:glow></a:effectLst>' +
          '</p:spPr></p:sp>',
      ],
    });
    const effects = sheet.shapes[0]!.effects;
    expect(effects?.map((e) => e.kind)).toEqual(['outerShdw', 'glow']);
    expect(effects?.[0]).toMatchObject({ blurRad: 50800, dist: 38100, dir: 2700000, algn: 'b' });
  });

  it('resolves a theme colour that is a sysClr with a lastClr', () => {
    const t = parseTheme(
      parseXmlString(
        themeXml(SCHEME).replace(
          '<a:dk1><a:srgbClr val="000000"/></a:dk1>',
          '<a:dk1><a:sysClr val="windowText" lastClr="112233"/></a:dk1>',
        ),
      ).root,
      '/ppt/theme/theme1.xml',
    );
    expect(toHexColor(resolveColor(t.scheme.dk1, {}))).toBe('112233');
  });
});

describe('the resolver reports where it looked', () => {
  it('names the level that declared the value', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 1 })] },
      { shapes: [phXml({ type: 'body', idx: 1 })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'M1' })] },
    );
    const found = resolveXfrm(onlyShape(slide), slide);
    expect(found?.origin).toBe('masterPh');
    expect(found?.explicit).toBe(false);
    expect(found?.sheet?.kind).toBe('master');
  });

  it('marks a value the shape stated itself as explicit', () => {
    const { slide } = chain(
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A3' })] },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A1' })] },
    );
    const found = resolveXfrm(onlyShape(slide), slide);
    expect(found?.origin).toBe('shape');
    expect(found?.explicit).toBe(true);
  });

  it('returns undefined rather than a default when nobody said', () => {
    const { slide } = chain({ shapes: [phXml({ type: 'body', idx: 1 })] }, {});
    expect(resolve(onlyShape(slide), slide, (s) => s.fill)).toBeUndefined();
  });

  it('does not walk the chain for a shape that is not a placeholder', () => {
    const { slide } = chain(
      {
        shapes: [
          '<p:sp><p:nvSpPr><p:cNvPr id="9" name="free"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
            '<p:spPr/></p:sp>',
        ],
      },
      { shapes: [phXml({ type: 'body', idx: 1, box: 'A1' })] },
    );
    expect(inheritanceChain(onlyShape(slide), slide)).toHaveLength(1);
    expect(resolveXfrm(onlyShape(slide), slide)).toBeUndefined();
  });
});

/* -------------------------------------------------------------------------- */
/* picture shapes                                                             */
/* -------------------------------------------------------------------------- */

/**
 * `CT_Picture` is `nvPicPr, blipFill, spPr`, so a picture's image is a sibling of
 * its shape properties and in PresentationML. Every rule here is re-derived from
 * `corpus/ground-truth/pictures.json`, where 23 of 24 readings were refuted at a
 * full channel. ADR 0037.
 */
describe('a picture shape', () => {
  const STRETCH = '<a:stretch><a:fillRect/></a:stretch>';

  function pic(blipFill: string, spPrExtra = ''): string {
    return (
      '<p:pic><p:nvPicPr><p:cNvPr id="7" name="picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      blipFill +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrExtra}</p:spPr></p:pic>`
    );
  }

  function probe(id: string) {
    const found = pictures.samples.find((sample) => sample.id === id);
    if (found === undefined) throw new Error(`no probe called ${id} in the fixture`);
    return found;
  }

  it('takes its fill from p:blipFill, which p:spPr does not contain', () => {
    const sheet = parse('slide', {
      shapes: [pic(`<p:blipFill><a:blip r:embed="rId9"/>${STRETCH}</p:blipFill>`)],
    });
    const fill = sheet.shapes[0]!.fill;
    if (fill?.type !== 'blip') throw new Error('expected a blip fill');
    expect(fill.embed).toBe('rId9');
    expect(fill.svgEmbed).toBeNull();
  });

  /**
   * PowerPoint painted the image on `precedence-solid`, not the cyan `a:solidFill`
   * beside it - so a reading that lets `p:spPr` win is 255 out on every quadrant.
   */
  it('paints the image rather than a competing p:spPr fill', () => {
    const measured = probe('precedence-solid');
    expect(measured.at.tl).toBe('FF0000');
    expect(measured.at.tl).not.toBe('00FFFF');

    const sheet = parse('slide', {
      shapes: [
        pic(
          `<p:blipFill><a:blip r:embed="rId9"/>${STRETCH}</p:blipFill>`,
          '<a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill>',
        ),
      ],
    });
    expect(sheet.shapes[0]!.fill?.type).toBe('blip');
  });

  /** An SVG-only picture: PowerPoint writes an `a:blip` with no `@r:embed` at all. */
  it('reads an SVG-only blip as naming no raster', () => {
    const svg =
      '<a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" ' +
      'r:embed="rId4"/></a:ext></a:extLst></a:blip>';
    const sheet = parse('slide', { shapes: [pic(`<p:blipFill>${svg}${STRETCH}</p:blipFill>`)] });
    const fill = sheet.shapes[0]!.fill;
    if (fill?.type !== 'blip') throw new Error('expected a blip fill');
    expect(fill.embed).toBeNull();
    expect(fill.svgEmbed).toBe('rId4');
  });

  it('keeps both images when a blip names a raster and an SVG', () => {
    const both =
      '<a:blip r:embed="rId3"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" ' +
      'r:embed="rId2"/></a:ext></a:extLst></a:blip>';
    const sheet = parse('slide', { shapes: [pic(`<p:blipFill>${both}${STRETCH}</p:blipFill>`)] });
    const fill = sheet.shapes[0]!.fill;
    if (fill?.type !== 'blip') throw new Error('expected a blip fill');
    expect(fill.embed).toBe('rId3');
    expect(fill.svgEmbed).toBe('rId2');
  });

  /**
   * An extension is identified by its `@uri`, never by what it contains: a blip
   * routinely carries `{28A0092B-...}` useLocalDpi beside the SVG one.
   */
  it('ignores an svgBlip under an extension that is not the SVG one', () => {
    const decoy =
      '<a:blip r:embed="rId3"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' +
      '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" ' +
      'r:embed="rId7"/></a:ext></a:extLst></a:blip>';
    const sheet = parse('slide', { shapes: [pic(`<p:blipFill>${decoy}${STRETCH}</p:blipFill>`)] });
    const fill = sheet.shapes[0]!.fill;
    if (fill?.type !== 'blip') throw new Error('expected a blip fill');
    expect(fill.embed).toBe('rId3');
    expect(fill.svgEmbed).toBeNull();
  });

  it('refuses a blip that names no image at all', () => {
    expect(() =>
      parse('slide', { shapes: [pic(`<p:blipFill><a:blip/>${STRETCH}</p:blipFill>`)] }),
    ).toThrow(ModelError);
  });

  /**
   * The rule the fixture settled, so a later edit that quietly re-reads one of
   * the four questions fails here rather than only in a rendered pixel.
   */
  it('is the reading the experiment left standing', () => {
    expect(pictures.rule).toEqual({
      clip: 'geometry',
      precedence: 'blip',
      outline: 'outside',
      mirror: 'image',
    });
    expect(pictures.outlineBands.picture).toEqual({ outsidePt: 12, insidePt: 0 });
    expect(pictures.outlineBands.shape).toEqual({ outsidePt: 6, insidePt: 6 });
  });

  /** A flip mirrors the image itself: the quadrants swap, they do not stay put. */
  it.each([
    ['mirror-h', 'tl', '00FF00'],
    ['mirror-h', 'tr', 'FF0000'],
    ['mirror-v', 'tl', '0000FF'],
    ['mirror-hv', 'tl', 'FFFF00'],
  ])('%s puts %s at %s', (id, sample, expected) => {
    expect(probe(id).at[sample as 'tl']).toBe(expected);
    expect(probe('control-rect').at[sample as 'tl']).not.toBe(expected);
  });

  /** A non-rect geometry clips the image; the bounding box does not survive. */
  it.each([
    ['clip-ellipse', 'corner', '000000'],
    ['clip-triangle', 'tl', '000000'],
  ])('%s paints the background at %s', (id, sample, expected) => {
    expect(probe(id).at[sample as 'tl']).toBe(expected);
    expect(probe('control-rect').at[sample as 'tl']).not.toBe(expected);
  });
});

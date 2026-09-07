/**
 * The text cascade, against what PowerPoint actually did.
 *
 * The fixture is imported and the rules are **re-derived from it**, never
 * compared against a summary of it. Each case carries the situation its package
 * put PowerPoint in - which level declared what, and which placeholder types the
 * slide, the layout and the master had - and the test rebuilds that situation
 * out of XML, runs it through this package's own resolver, and checks the number
 * PowerPoint reported. So a resolver that drifts fails on 55 measurements rather
 * than on one sentence somebody wrote about them.
 *
 * The refutations are here too, and they are the point. It is worth little that
 * the measured bucket rule fits; what is worth something is that the rule the
 * plan specified - and that every implementation this project has read uses -
 * fits 13 of the same 37 cases, and that the test says so by scoring it.
 */

import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/text-cascade.json' with { type: 'json' };

import { BUILTIN_TEXT_STYLES, TEXT_FLOOR } from './builtin-text-styles.js';
import { ModelError } from './errors.js';
import { parseSheet, parseTheme } from './parse/sheet.js';
import { parseListStyle } from './parse/text.js';
import {
  bucketOf,
  floorOf,
  resolveIndent,
  resolveMarginLeft,
  resolveParagraph,
  resolveRun,
  resolveSize,
  textLevels,
  type TextContext,
} from './resolve/text.js';
import { themeFontRef, type Typeface } from './text.js';
import { requestedTypefaces, resolveTypeface } from './resolve/typeface.js';
import type { FontScheme } from './types.js';
import type { Paragraph, Sheet, TextContent } from './index.js';

/* -------------------------------------------------------------------------- */
/* building packages by hand                                                  */
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

const MAJOR = fixture.themeFonts.major;
const MINOR = fixture.themeFonts.minor;

let nextId = 100;

/** `a:lvlNpPr` holding whatever the case declares, or nothing at all. */
function lvlPr(level: number, inner: string, attrs = ''): string {
  const tag = `a:lvl${String(level)}pPr`;
  const open = `<${tag}${attrs === '' ? '' : ` ${attrs}`}`;
  return inner === '' ? `${open}/>` : `${open}>${inner}</${tag}>`;
}

const sizeRPr = (hundredths: number): string =>
  hundredths === 0 ? '<a:defRPr/>' : `<a:defRPr sz="${String(hundredths)}"/>`;

/** An `a:lstStyle` declaring one size at level one, as the ladder decks do. */
const sizeStyle = (hundredths: number): string =>
  `<a:lstStyle>${lvlPr(1, sizeRPr(hundredths))}</a:lstStyle>`;

interface ShapeSpec {
  readonly name: string;
  readonly ph?: string | undefined;
  readonly lstStyle?: string | undefined;
  readonly paragraphs?: readonly string[] | undefined;
}

function spXml(spec: ShapeSpec): string {
  const ph = spec.ph === undefined ? '' : `<p:ph ${spec.ph}/>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${String(nextId++)}" name="${spec.name}"/><p:cNvSpPr/>` +
    `<p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    '<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/>' +
    (spec.lstStyle ?? '<a:lstStyle/>') +
    (spec.paragraphs ?? ['<a:p><a:r><a:rPr lang="en-US"/><a:t>Ag</a:t></a:r></a:p>']).join('') +
    '</p:txBody></p:sp>'
  );
}

interface SheetSpec {
  readonly shapes?: readonly string[] | undefined;
  readonly txStyles?: string | undefined;
}

function sheetXml(kind: 'slide' | 'layout' | 'master', spec: SheetSpec = {}): string {
  const tag = kind === 'slide' ? 'p:sld' : kind === 'layout' ? 'p:sldLayout' : 'p:sldMaster';
  const attrs = kind === 'layout' ? ' type="obj"' : '';
  // `p:txStyles` follows `p:clrMap` in `CT_SlideMaster`, and out of order it is
  // a repair prompt rather than a fixture.
  const tail =
    kind === 'master'
      ? CLR_MAP + (spec.txStyles ?? '')
      : '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
  return (
    `<${tag} ${NS}${attrs}><p:cSld name="${kind}">${SP_TREE_HEAD}` +
    `${(spec.shapes ?? []).join('')}</p:spTree></p:cSld>${tail}</${tag}>`
  );
}

function themeXml(objectDefaults = ''): string {
  const font = (kind: 'major' | 'minor', latin: string): string =>
    `<a:${kind}Font><a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/></a:${kind}Font>`;
  return (
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
    '<a:themeElements><a:clrScheme name="T">' +
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
    ]
      .map((k) => `<a:${k}><a:srgbClr val="123456"/></a:${k}>`)
      .join('') +
    '</a:clrScheme>' +
    `<a:fontScheme name="T">${font('major', MAJOR)}${font('minor', MINOR)}</a:fontScheme>` +
    '<a:fmtScheme name="T"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/>' +
    '</a:fmtScheme></a:themeElements>' +
    objectDefaults +
    '</a:theme>'
  );
}

function parseSheetSpec(kind: 'slide' | 'layout' | 'master', spec: SheetSpec = {}) {
  return parseSheet(parseXmlString(sheetXml(kind, spec)).root, `/ppt/${kind}.xml`);
}

/** A three-link chain, wired the way `document.ts` wires one. */
function chain(
  slide: SheetSpec,
  layout: SheetSpec,
  master: SheetSpec,
  objectDefaults = '',
): { slide: Sheet; layout: Sheet; master: Sheet } {
  const theme = parseTheme(parseXmlString(themeXml(objectDefaults)).root, '/ppt/theme/theme1.xml');
  const masterSheet: Sheet = { ...parseSheetSpec('master', master), parent: null, theme };
  const layoutSheet: Sheet = {
    ...parseSheetSpec('layout', layout),
    parent: masterSheet,
    theme: null,
  };
  const slideSheet: Sheet = { ...parseSheetSpec('slide', slide), parent: layoutSheet, theme: null };
  return { slide: slideSheet, layout: layoutSheet, master: masterSheet };
}

function listStyle(xml: string) {
  return parseListStyle(parseXmlString(`<a:lstStyle ${NS}>${xml}</a:lstStyle>`).root, '/x.xml');
}

/** The named shape on a sheet, or a failure that says which name was missing. */
function shapeNamed(sheet: Sheet, name: string) {
  const shape = sheet.shapes.find((s) => s.name === name);
  if (shape === undefined) throw new Error(`no shape named ${name}`);
  return shape;
}

function firstParagraph(sheet: Sheet, name: string): Paragraph {
  const paragraph = shapeNamed(sheet, name).text?.paragraphs[0];
  if (paragraph === undefined) throw new Error(`${name} has no paragraph`);
  return paragraph;
}

const firstRun = (paragraph: Paragraph): TextContent | undefined => paragraph.content[0];

/* -------------------------------------------------------------------------- */
/* the situations the fixture records                                         */
/* -------------------------------------------------------------------------- */

interface BucketSizes {
  readonly title: number;
  readonly body: number;
  readonly other: number;
}

interface LadderSituation {
  readonly kind: 'ladder';
  readonly shape: 'ph' | 'plain';
  readonly declares: Readonly<Record<string, number>>;
  readonly txStyles: BucketSizes | null;
  readonly defaultTextStyle: number | null;
  readonly objectDefaults: number | null;
}

interface BucketSituation {
  readonly kind: 'bucket';
  readonly slideType: string | null;
  readonly layoutType: string | null;
  readonly masterType: string | null;
  readonly txStyles: BucketSizes | null;
  readonly defaultTextStyle: number | null;
}

interface Case {
  readonly id: string;
  readonly question: string;
  readonly situation?: LadderSituation | BucketSituation | undefined;
  readonly measured: {
    readonly sizePt: number | null;
    readonly font: string | null;
    readonly bold: number | null;
    readonly italic: number | null;
    readonly underline: number | null;
    readonly strike: number | null;
    readonly indentLevel: number | null;
    readonly leftIndentPt: number | null;
    readonly firstLineIndentPt: number | null;
    readonly spaceWithin: number | null;
    readonly lineRuleWithin: number | null;
    readonly spaceBefore: number | null;
    readonly lineRuleBefore: number | null;
  };
}

/**
 * The fixture's case lists, named rather than indexed.
 *
 * An index signature would be shorter and would make every access
 * `Case[] | undefined` under `noUncheckedIndexedAccess`, which turns a missing
 * list - a real failure, and one worth failing on - into forty optional chains
 * that quietly test nothing.
 */
interface Findings {
  readonly ladderCases: readonly Case[];
  readonly ladderPlainCases: readonly Case[];
  readonly bucketCases: readonly Case[];
  readonly levelCases: readonly Case[];
  readonly builtinCases: readonly Case[];
  readonly hopCases: readonly Case[];
  readonly mergeCases: readonly Case[];
  readonly marginCases: readonly Case[];
  readonly percentCases: readonly Case[];
  readonly fontCases: readonly Case[];
}

const cases = fixture.findings as unknown as Findings;

function txStylesXml(sizes: BucketSizes | null): string | undefined {
  if (sizes === null) return undefined;
  const bucket = (tag: string, hundredths: number): string =>
    `<p:${tag}>${lvlPr(1, sizeRPr(hundredths))}</p:${tag}>`;
  return (
    '<p:txStyles>' +
    bucket('titleStyle', sizes.title) +
    bucket('bodyStyle', sizes.body) +
    bucket('otherStyle', sizes.other) +
    '</p:txStyles>'
  );
}

/** The six types a master may carry, so a folded second hop always lands. */
const MASTER_PH = ['title', 'body', 'dt', 'ftr', 'sldNum', 'hdr'] as const;

function masterShapes(lstStyleFor?: string): readonly string[] {
  return MASTER_PH.map((type, i) =>
    spXml({
      name: `m-${type}`,
      ph: type === 'title' ? 'type="title"' : `type="${type}" idx="${String(i)}"`,
      ...(type === 'body' && lstStyleFor !== undefined ? { lstStyle: lstStyleFor } : {}),
    }),
  );
}

/** Rebuild a ladder case's package and return everything the resolver needs. */
function buildLadder(situation: LadderSituation): {
  context: TextContext;
  paragraph: Paragraph;
} {
  const declared = (name: string): string | undefined => {
    const value = situation.declares[name];
    return value === undefined ? undefined : sizeStyle(value);
  };
  const runSz = situation.declares['run'];
  const paraSz = situation.declares['para'];
  const paragraphXml =
    '<a:p>' +
    (paraSz === undefined ? '' : `<a:pPr><a:defRPr sz="${String(paraSz)}"/></a:pPr>`) +
    `<a:r><a:rPr lang="en-US"${runSz === undefined ? '' : ` sz="${String(runSz)}"`}/>` +
    '<a:t>Ag</a:t></a:r></a:p>';

  const objectDefaults =
    situation.objectDefaults === null
      ? ''
      : '<a:objectDefaults><a:spDef><a:spPr/><a:bodyPr/>' +
        sizeStyle(situation.objectDefaults) +
        '</a:spDef></a:objectDefaults>';

  const sheets = chain(
    {
      shapes: [
        spXml({
          name: 'probe',
          ...(situation.shape === 'ph' ? { ph: 'type="body" idx="1"' } : {}),
          ...(declared('shapeLst') === undefined ? {} : { lstStyle: declared('shapeLst') }),
          paragraphs: [paragraphXml],
        }),
      ],
    },
    {
      shapes: [
        spXml({
          name: 'l-body',
          ph: 'type="body" idx="1"',
          ...(declared('layoutPh') === undefined ? {} : { lstStyle: declared('layoutPh') }),
        }),
      ],
    },
    {
      shapes: masterShapes(declared('masterPh')),
      ...(txStylesXml(situation.txStyles) === undefined
        ? {}
        : { txStyles: txStylesXml(situation.txStyles) }),
    },
    objectDefaults,
  );

  return {
    context: {
      sheet: sheets.slide,
      shape: shapeNamed(sheets.slide, 'probe'),
      defaultTextStyle:
        situation.defaultTextStyle === null
          ? undefined
          : listStyle(lvlPr(1, sizeRPr(situation.defaultTextStyle))),
    },
    paragraph: firstParagraph(sheets.slide, 'probe'),
  };
}

/** Rebuild a bucket case's package. */
function buildBucket(situation: BucketSituation): {
  context: TextContext;
  paragraph: Paragraph;
} {
  const sheets = chain(
    {
      shapes: [
        spXml({
          name: 'probe',
          ...(situation.slideType === null ? {} : { ph: `type="${situation.slideType}" idx="0"` }),
        }),
      ],
    },
    {
      shapes:
        situation.layoutType === null
          ? []
          : [spXml({ name: 'l-ph', ph: `type="${situation.layoutType}" idx="0"` })],
    },
    {
      shapes: masterShapes(),
      ...(txStylesXml(situation.txStyles) === undefined
        ? {}
        : { txStyles: txStylesXml(situation.txStyles) }),
    },
  );
  return {
    context: {
      sheet: sheets.slide,
      shape: shapeNamed(sheets.slide, 'probe'),
      defaultTextStyle:
        situation.defaultTextStyle === null
          ? undefined
          : listStyle(lvlPr(1, sizeRPr(situation.defaultTextStyle))),
    },
    paragraph: firstParagraph(sheets.slide, 'probe'),
  };
}

function resolvePt(built: { context: TextContext; paragraph: Paragraph }): number {
  return resolveSize(built.context, built.paragraph, firstRun(built.paragraph)).value / 100;
}

/* -------------------------------------------------------------------------- */

describe('the ladder: which source wins', () => {
  const ladder = [...cases.ladderCases, ...cases.ladderPlainCases];

  it('has every rung, for a placeholder and for a shape that is not one', () => {
    expect(cases.ladderCases.length).toBe(9);
    expect(cases.ladderPlainCases.length).toBe(9);
  });

  it.each(ladder.map((c) => [c.id, c] as const))('reproduces %s', (_id: string, probe: Case) => {
    const situation = probe.situation;
    if (situation?.kind !== 'ladder') throw new Error(`${probe.id} has no ladder situation`);
    expect(resolvePt(buildLadder(situation))).toBe(probe.measured.sizePt);
  });

  it('rules out every other order of the sources, as the analysis did', () => {
    // The scoring is re-stated here rather than trusted: none of the 40320
    // permutations fits while a placeholder is assumed to read all eight
    // sources, and exactly one read order fits once the two it ignores go.
    expect(fixture.scores.ladder.permutationsOverAllSources).toBe(40320);
    expect(fixture.scores.ladder.fittingWithEverySourceRead).toBe(0);
    expect(fixture.scores.ladder.distinctReadOrdersFitting).toBe(1);
    expect(fixture.scores.ladder.sourcesIgnored).toEqual(['defaultText', 'objDefaults']);
    expect(fixture.model.placeholderOrder).toEqual([
      'run',
      'para',
      'shapeLst',
      'layoutPh',
      'masterPh',
      'txStyles',
    ]);
  });

  it('names the level that answered', () => {
    const situation = cases.ladderCases.find((c) => c.id === 'order-4-layoutPh-ph')?.situation;
    if (situation?.kind !== 'ladder') throw new Error('order-4-layoutPh-ph is missing');
    const built = buildLadder(situation);
    const size = resolveSize(built.context, built.paragraph, firstRun(built.paragraph));
    expect(size.origin).toBe('layoutPh');
    expect(size.explicit).toBe(false);
    expect(size.sheet?.kind).toBe('layout');
  });

  it('marks a value the run stated itself as explicit', () => {
    const situation = cases.ladderCases.find((c) => c.id === 'order-1-run-ph')?.situation;
    if (situation?.kind !== 'ladder') throw new Error('order-1-run-ph is missing');
    const built = buildLadder(situation);
    const size = resolveSize(built.context, built.paragraph, firstRun(built.paragraph));
    expect(size.origin).toBe('run');
    expect(size.explicit).toBe(true);
  });
});

describe('the bucket: whose @type selects it', () => {
  const bucketCases = cases.bucketCases;

  it('has the cases the analysis scored', () => {
    expect(bucketCases.length).toBe(37);
  });

  it.each(bucketCases.map((c) => [c.id, c] as const))(
    'reproduces %s',
    (_id: string, probe: Case) => {
      const situation = probe.situation;
      if (situation?.kind !== 'bucket') throw new Error(`${probe.id} has no bucket situation`);
      expect(resolvePt(buildBucket(situation))).toBe(probe.measured.sizePt);
    },
  );

  it('refutes the rule every implementation uses, scored over the same cases', () => {
    // The plan's rule: the shape's own @type picks the bucket, the
    // header-and-footer trio reads p:otherStyle, and so does a shape that is not
    // a placeholder. Scored here rather than asserted, so the claim is a number.
    const planned = (situation: BucketSituation): number => {
      const type = situation.slideType;
      const sizes = situation.txStyles;
      if (sizes === null) return TEXT_FLOOR.sz;
      // A bucket that exists and declares nothing ends the walk at the floor, in
      // this reading as in the measured one - the two disagree about *which*
      // bucket, and giving the plan's reading the benefit of the shared rule is
      // what makes the score a comparison rather than a straw man.
      const declared = (value: number): number => (value === 0 ? TEXT_FLOOR.sz : value);
      if (type === null) return declared(sizes.other);
      if (type === 'title' || type === 'ctrTitle') return declared(sizes.title);
      if (type === 'dt' || type === 'ftr' || type === 'sldNum' || type === 'hdr') {
        return declared(sizes.other);
      }
      return declared(sizes.body);
    };
    let hits = 0;
    for (const probe of bucketCases) {
      const situation = probe.situation;
      if (situation?.kind !== 'bucket') continue;
      const measured = probe.measured.sizePt;
      if (measured !== null && planned(situation) / 100 === measured) hits += 1;
    }
    expect(hits).toBeLessThan(bucketCases.length);
    expect(hits).toBe(13);
    const winner = fixture.scores.bucket.find((s) => s.hits === s.of);
    expect(winner?.model).toBe('chainEnd/hf-none/plain-none');
  });

  it('sends the four header-and-footer types to no bucket at all', () => {
    expect(bucketOf('title')).toBe('title');
    expect(bucketOf('ctrTitle')).toBe('title');
    expect(bucketOf('body')).toBe('body');
    expect(bucketOf('subTitle')).toBe('body');
    expect(bucketOf('pic')).toBe('body');
    for (const type of ['dt', 'ftr', 'sldNum', 'hdr']) expect(bucketOf(type)).toBeNull();
  });

  it('never walks p:otherStyle, whatever the shape', () => {
    // Every level the resolver would consult, for all sixteen placeholder types
    // and for a shape that is not one. `p:otherStyle` is in the file and in the
    // model and is reached by nothing - measured on the slide, the layout and
    // the master.
    let seen = 0;
    for (const probe of bucketCases) {
      const situation = probe.situation;
      if (situation?.kind !== 'bucket') continue;
      const built = buildBucket(situation);
      for (const level of textLevels(built.context)) {
        seen += 1;
        expect(level.style).not.toBe(built.context.sheet.parent?.parent?.txStyles?.other);
      }
    }
    expect(seen).toBeGreaterThan(bucketCases.length);
  });
});

describe('levels are independent, all the way up', () => {
  const LEVEL_PT = [8, 12, 16, 20, 24, 28, 32, 36, 40];

  function levelChain(masterBody: string, layoutStyle?: string) {
    const paragraphs = LEVEL_PT.map((_, i) => `<a:p><a:pPr lvl="${String(i)}"/></a:p>`);
    return chain(
      { shapes: [spXml({ name: 'probe', ph: 'type="body" idx="1"', paragraphs })] },
      {
        shapes: [
          spXml({
            name: 'l-body',
            ph: 'type="body" idx="1"',
            ...(layoutStyle === undefined ? {} : { lstStyle: layoutStyle }),
          }),
        ],
      },
      {
        shapes: masterShapes(),
        txStyles:
          '<p:txStyles><p:titleStyle/>' +
          `<p:bodyStyle>${masterBody}</p:bodyStyle>` +
          '<p:otherStyle/></p:txStyles>',
      },
    );
  }

  it('maps @lvl n to a:lvl(n+1)pPr, over all nine', () => {
    const body = LEVEL_PT.map((pt, i) => lvlPr(i + 1, sizeRPr(pt * 100))).join('');
    const sheets = levelChain(body);
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const measured = cases.levelCases.filter((c) => c.id.startsWith('level-body-'));
    expect(measured.length).toBe(9);
    for (const [i, probe] of measured.entries()) {
      const paragraph = shape.text?.paragraphs[i];
      if (paragraph === undefined) throw new Error('missing paragraph');
      expect(paragraph.level).toBe(i);
      expect(resolveSize(context, paragraph, undefined).value / 100).toBe(probe.measured.sizePt);
    }
  });

  it('lets a layout declare one level and leave its neighbours to the master', () => {
    const body = [8, 12, 16].map((pt, i) => lvlPr(i + 1, sizeRPr(pt * 100))).join('');
    const sheets = levelChain(body, `<a:lstStyle>${lvlPr(2, sizeRPr(5000))}</a:lstStyle>`);
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const measured = cases.levelCases.filter((c) => c.id.startsWith('level-cross-'));
    expect(measured.length).toBe(3);
    for (const [i, probe] of measured.entries()) {
      const paragraph = shape.text?.paragraphs[i];
      if (paragraph === undefined) throw new Error('missing paragraph');
      expect(resolveSize(context, paragraph, undefined).value / 100).toBe(probe.measured.sizePt);
    }
  });

  it('does not let a nearer style that declares level one shadow the rest', () => {
    // The discriminating case, and the one `level-cross` does *not* cover: the
    // layout placeholder declares a:lvl1pPr and nothing else, so a paragraph at
    // lvl="1" has a nearer style that speaks - but not about the level being
    // asked for. Measured 50/12/16: the levels are independent. A resolver that
    // fell back to a:lvl1pPr would say 50/50/50 and make every second-level
    // bullet the size of a first-level one.
    const masterBody =
      '<a:lvl1pPr marL="228600"><a:defRPr sz="800"/></a:lvl1pPr>' +
      '<a:lvl2pPr marL="457200"><a:defRPr sz="1200"/></a:lvl2pPr>' +
      '<a:lvl3pPr marL="685800"><a:defRPr sz="1600"/></a:lvl3pPr>';
    const sheets = levelChain(
      masterBody,
      '<a:lstStyle><a:lvl1pPr marL="914400"><a:defRPr sz="5000"/></a:lvl1pPr></a:lstStyle>',
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const measured = cases.levelCases.filter((c) => c.id.startsWith('level-shadow-'));
    expect(measured.length).toBe(3);
    for (const [i, probe] of measured.entries()) {
      const paragraph = shape.text?.paragraphs[i];
      if (paragraph === undefined) throw new Error('missing paragraph');
      expect(resolveSize(context, paragraph, undefined).value / 100).toBe(probe.measured.sizePt);
      expect(resolveMarginLeft(context, paragraph).value / 12700).toBe(probe.measured.leftIndentPt);
    }
    // And the refutation, stated as a number: the shadowing reading gives 50 at
    // every level, and the measurement agrees with it at exactly one.
    expect(measured.filter((c) => c.measured.sizePt === 50).length).toBe(1);
  });

  it('does not fall back to a:lvl1pPr for a level nobody declared', () => {
    // The refutation: a shape lstStyle declaring only a:lvl3pPr must leave
    // lvl="0" to the master. A cascade that quietly used level one when the
    // asked-for level is absent gives both paragraphs 64.
    const body = [8, 12, 16].map((pt, i) => lvlPr(i + 1, sizeRPr(pt * 100))).join('');
    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 'probe',
            ph: 'type="body" idx="1"',
            lstStyle: `<a:lstStyle>${lvlPr(3, sizeRPr(6400))}</a:lstStyle>`,
            paragraphs: ['<a:p><a:pPr lvl="0"/></a:p>', '<a:p><a:pPr lvl="2"/></a:p>'],
          }),
        ],
      },
      { shapes: [spXml({ name: 'l-body', ph: 'type="body" idx="1"' })] },
      {
        shapes: masterShapes(),
        txStyles: `<p:txStyles><p:titleStyle/><p:bodyStyle>${body}</p:bodyStyle><p:otherStyle/></p:txStyles>`,
      },
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const sparse = cases.levelCases.filter((c) => c.id.startsWith('level-sparse-'));
    expect(sparse.length).toBe(2);
    for (const [i, probe] of sparse.entries()) {
      const paragraph = shape.text?.paragraphs[i];
      if (paragraph === undefined) throw new Error('missing paragraph');
      expect(resolveSize(context, paragraph, undefined).value / 100).toBe(probe.measured.sizePt);
    }
    expect(sparse[0]?.measured.sizePt).not.toBe(sparse[1]?.measured.sizePt);
  });

  it('clamps a level outside ST_TextIndentLevelType rather than refusing the file', () => {
    const sheets = chain(
      { shapes: [spXml({ name: 'probe', paragraphs: ['<a:p><a:pPr lvl="42"/></a:p>'] })] },
      {},
      { shapes: masterShapes() },
    );
    expect(firstParagraph(sheets.slide, 'probe').level).toBe(8);
  });

  it('refuses a level that is not an integer', () => {
    expect(() =>
      chain(
        { shapes: [spXml({ name: 'probe', paragraphs: ['<a:p><a:pPr lvl="two"/></a:p>'] })] },
        {},
        {},
      ),
    ).toThrow(ModelError);
  });
});

describe('the cascade merges per property', () => {
  it('takes one property from each of five levels', () => {
    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 'probe',
            ph: 'type="body" idx="1"',
            lstStyle: `<a:lstStyle>${lvlPr(1, '<a:defRPr i="1"/>')}</a:lstStyle>`,
            paragraphs: [
              '<a:p><a:pPr><a:defRPr u="sng"/></a:pPr>' +
                '<a:r><a:rPr lang="en-US" strike="sngStrike"/><a:t>Ag</a:t></a:r></a:p>',
            ],
          }),
        ],
      },
      {
        shapes: [
          spXml({
            name: 'l-body',
            ph: 'type="body" idx="1"',
            lstStyle: `<a:lstStyle>${lvlPr(1, '<a:defRPr b="1"/>')}</a:lstStyle>`,
          }),
        ],
      },
      {
        shapes: masterShapes(),
        txStyles: `<p:txStyles><p:titleStyle/><p:bodyStyle>${lvlPr(1, sizeRPr(2400))}</p:bodyStyle><p:otherStyle/></p:txStyles>`,
      },
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const paragraph = firstParagraph(sheets.slide, 'probe');
    const run = firstRun(paragraph);
    const probe = cases.mergeCases.find((c) => c.id === 'merge-per-property');
    if (probe === undefined) throw new Error('merge-per-property is missing');

    expect(resolveSize(context, paragraph, run).value / 100).toBe(probe.measured.sizePt);
    expect(resolveRun(context, paragraph, run, (p) => p.b)?.value).toBe(probe.measured.bold === -1);
    expect(resolveRun(context, paragraph, run, (p) => p.i)?.value).toBe(
      probe.measured.italic === -1,
    );
    expect(resolveRun(context, paragraph, run, (p) => p.u)?.value).toBe('sng');
    expect(resolveRun(context, paragraph, run, (p) => p.strike)?.value).toBe('sngStrike');

    // And each came from a different level, which is what "per property" means.
    expect(resolveRun(context, paragraph, run, (p) => p.strike)?.origin).toBe('run');
    expect(resolveRun(context, paragraph, run, (p) => p.u)?.origin).toBe('paragraph');
    expect(resolveRun(context, paragraph, run, (p) => p.i)?.origin).toBe('shape');
    expect(resolveRun(context, paragraph, run, (p) => p.b)?.origin).toBe('layoutPh');
    expect(resolveSize(context, paragraph, run).origin).toBe('txStyles');
  });
});

describe('marL and indent', () => {
  const marginCase = (id: string): Case => {
    const probe = cases.marginCases.find((c) => c.id === id);
    if (probe === undefined) throw new Error(`${id} is missing`);
    return probe;
  };

  function marginChain(masterMarL: number | null, layoutIndent: number | null, ph: boolean) {
    return chain(
      {
        shapes: [spXml({ name: 'probe', ...(ph ? { ph: 'type="body" idx="1"' } : {}) })],
      },
      {
        shapes: [
          spXml({
            name: 'l-body',
            ph: 'type="body" idx="1"',
            ...(layoutIndent === null
              ? {}
              : {
                  lstStyle: `<a:lstStyle>${lvlPr(1, '', `indent="${String(layoutIndent)}"`)}</a:lstStyle>`,
                }),
          }),
        ],
      },
      {
        shapes: masterShapes(),
        txStyles:
          '<p:txStyles><p:titleStyle/><p:bodyStyle>' +
          lvlPr(1, '', masterMarL === null ? '' : `marL="${String(masterMarL)}"`) +
          '</p:bodyStyle><p:otherStyle/></p:txStyles>',
      },
    );
  }

  it('takes marL and indent from different sheets and merges them', () => {
    const probe = marginCase('margin-split');
    const sheets = marginChain(742950, -285750, true);
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const paragraph = firstParagraph(sheets.slide, 'probe');
    expect(resolveMarginLeft(context, paragraph).value / 12700).toBe(probe.measured.leftIndentPt);
    expect(resolveIndent(context, paragraph).value / 12700).toBe(probe.measured.firstLineIndentPt);
    expect(resolveMarginLeft(context, paragraph).origin).toBe('txStyles');
    expect(resolveIndent(context, paragraph).origin).toBe('layoutPh');
  });

  it('resolves an undeclared indent to zero, not to the schema default', () => {
    const probe = marginCase('margin-bare');
    const sheets = marginChain(742950, null, true);
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const paragraph = firstParagraph(sheets.slide, 'probe');
    expect(resolveIndent(context, paragraph).value / 12700).toBe(probe.measured.firstLineIndentPt);
    expect(resolveIndent(context, paragraph).value).toBe(0);
    // The refutation. ECMA-376 gives @indent a default of -342900, and a parser
    // that materialises it puts a 27-point hanging indent on every paragraph.
    expect(resolveIndent(context, paragraph).value).not.toBe(
      fixture.findings.margins.schemaDefaults.indent,
    );
  });

  it('resolves an undeclared marL to zero for a shape that is not a placeholder', () => {
    const probe = marginCase('margin-plain');
    const sheets = marginChain(742950, null, false);
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const paragraph = firstParagraph(sheets.slide, 'probe');
    expect(resolveMarginLeft(context, paragraph).value / 12700).toBe(probe.measured.leftIndentPt);
    expect(resolveMarginLeft(context, paragraph).value).not.toBe(
      fixture.findings.margins.schemaDefaults.marL,
    );
  });
});

describe('ST_Percentage has two spellings and ST_TextSpacingPoint has one', () => {
  const spacing = (lnSpc: string) => {
    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 'probe',
            paragraphs: [`<a:p><a:pPr><a:lnSpc>${lnSpc}</a:lnSpc></a:pPr></a:p>`],
          }),
        ],
      },
      {},
      { shapes: masterShapes() },
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    return resolveParagraph(context, firstParagraph(sheets.slide, 'probe'), (p) => p.lnSpc)?.value;
  };

  it('reads both spellings of a percentage as the same number', () => {
    const numeric = cases.percentCases.find((c) => c.id === 'percent-numeric');
    const strict = cases.percentCases.find((c) => c.id === 'percent-strict');
    expect(numeric?.measured.spaceWithin).toBe(strict?.measured.spaceWithin);
    expect(spacing('<a:spcPct val="150000"/>')).toEqual({ kind: 'percent', value: 150000 });
    expect(spacing('<a:spcPct val="150%"/>')).toEqual({ kind: 'percent', value: 150000 });
  });

  it('does not read a percentage the way parseInt would', () => {
    // `parseInt("150%")` is 150, which is 0.15% line spacing, and the slide
    // collapses to a line. The refutation is worth a test of its own.
    expect(spacing('<a:spcPct val="150%"/>')).not.toEqual({ kind: 'percent', value: 150 });
  });

  it('keeps a:spcPts in hundredths of a point and says so', () => {
    const points = cases.percentCases.find((c) => c.id === 'percent-points');
    expect(points?.measured.lineRuleWithin).toBe(0);
    expect(spacing('<a:spcPts val="3000"/>')).toEqual({ kind: 'points', value: 3000 });
  });
});

describe('the theme font scheme', () => {
  it('reads +mj-lt and +mn-lt as references rather than typefaces', () => {
    expect(themeFontRef('+mj-lt')).toEqual({ collection: 'major', script: 'latin' });
    expect(themeFontRef('+mn-ea')).toEqual({ collection: 'minor', script: 'ea' });
    expect(themeFontRef('Georgia')).toBeNull();
    expect(themeFontRef('+mn')).toBeNull();
  });

  it('resolves the reference a p:txStyles bucket declares to the theme face', () => {
    const major = cases.fontCases.find((c) => c.id === 'font-major');
    const minor = cases.fontCases.find((c) => c.id === 'font-minor');
    const sheets = chain(
      {
        shapes: [
          spXml({ name: 's-title', ph: 'type="title" idx="0"' }),
          spXml({ name: 's-body', ph: 'type="body" idx="1"' }),
        ],
      },
      {
        shapes: [
          spXml({ name: 'l-title', ph: 'type="title" idx="0"' }),
          spXml({ name: 'l-body', ph: 'type="body" idx="1"' }),
        ],
      },
      {
        shapes: masterShapes(),
        txStyles:
          '<p:txStyles>' +
          `<p:titleStyle>${lvlPr(1, '<a:defRPr sz="4000"><a:latin typeface="+mj-lt"/></a:defRPr>')}</p:titleStyle>` +
          `<p:bodyStyle>${lvlPr(1, '<a:defRPr sz="2400"><a:latin typeface="+mn-lt"/></a:defRPr>')}</p:bodyStyle>` +
          '<p:otherStyle/></p:txStyles>',
      },
    );
    const faceOf = (name: string): string => {
      const shape = shapeNamed(sheets.slide, name);
      const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
      const paragraph = firstParagraph(sheets.slide, name);
      const latin = resolveRun(context, paragraph, undefined, (p) => p.latin)?.value;
      if (latin === undefined) throw new Error(`${name} resolved no typeface`);
      const ref = themeFontRef(latin.typeface);
      if (ref === null) return latin.typeface;
      const fonts = sheets.master.theme?.fonts;
      if (fonts === undefined) throw new Error('the master has no theme');
      return fonts[ref.collection][ref.script] ?? '';
    };
    expect(faceOf('s-title')).toBe(major?.measured.font);
    expect(faceOf('s-body')).toBe(minor?.measured.font);
    expect(faceOf('s-title')).toBe(MAJOR);
    expect(faceOf('s-body')).toBe(MINOR);
  });
});

describe('the sources that are not sources', () => {
  it('ignores a:objectDefaults, which the plan named as a level', () => {
    const probe = cases.ladderPlainCases.find((c) => c.id === 'order-8-objDefaults-plain');
    if (probe?.situation?.kind !== 'ladder')
      throw new Error('order-8-objDefaults-plain is missing');
    expect(probe.situation.objectDefaults).not.toBeNull();
    // The theme really does carry the declaration - `buildLadder` writes it -
    // and the resolved size is the floor regardless.
    expect(resolvePt(buildLadder(probe.situation))).toBe(probe.measured.sizePt);
    expect(resolvePt(buildLadder(probe.situation))).toBe(TEXT_FLOOR.sz / 100);
    expect(fixture.model.notSources).toContain('objDefaults');
  });

  it('parses a:defPPr and never resolves against it', () => {
    const style = listStyle('<a:defPPr><a:defRPr sz="3000"/></a:defPPr>');
    expect(style.defPPr?.defRPr?.sz).toBe(3000);
    expect(style.levels.every((level) => level === undefined)).toBe(true);

    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 'probe',
            lstStyle: '<a:lstStyle><a:defPPr><a:defRPr sz="3000"/></a:defPPr></a:lstStyle>',
          }),
        ],
      },
      {},
      { shapes: masterShapes() },
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const probe = cases.ladderCases.length > 0 ? fixture.findings.negatives.defPPrOnShape : null;
    expect(probe?.measuredPt).toBe(TEXT_FLOOR.sz / 100);
    expect(resolveSize(context, firstParagraph(sheets.slide, 'probe'), undefined).value / 100).toBe(
      probe?.measuredPt,
    );
  });

  it('records that nothing reads p:otherStyle, on any of the three sheets', () => {
    const n = fixture.findings.negatives;
    expect(n.otherStyleOnSlide.measuredPt).toBe(TEXT_FLOOR.sz / 100);
    expect(n.otherStyleOnLayout.measuredPt).toBe(TEXT_FLOOR.sz / 100);
    expect(n.otherStyleOnMaster.measuredPt).toBe(TEXT_FLOOR.sz / 100);
    expect(fixture.model.notSources).toContain('otherStyle');
  });
});

describe("PowerPoint's own text styles, for a master that declares none", () => {
  it('matches the committed table, level for level', () => {
    for (const bucket of ['title', 'body', 'other'] as const) {
      const measured = fixture.model.builtinTextStyles[bucket];
      expect(BUILTIN_TEXT_STYLES[bucket].length).toBe(9);
      for (const [i, level] of measured.entries()) {
        expect(BUILTIN_TEXT_STYLES[bucket][i]).toEqual({
          sz: level.szHundredths,
          marL: level.marLEmu,
          indent: level.indentEmu,
          bullet: level.bullet,
        });
      }
    }
  });

  it('reproduces every built-in probe', () => {
    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 's-title',
            ph: 'type="title" idx="0"',
            paragraphs: [0, 1, 2, 3, 4, 5, 6, 7, 8].map(
              (l) => `<a:p><a:pPr lvl="${String(l)}"/></a:p>`,
            ),
          }),
          spXml({
            name: 's-body',
            ph: 'type="body" idx="1"',
            paragraphs: [0, 1, 2, 3, 4, 5, 6, 7, 8].map(
              (l) => `<a:p><a:pPr lvl="${String(l)}"/></a:p>`,
            ),
          }),
          spXml({
            name: 's-plain',
            paragraphs: [0, 1, 2, 3, 4, 5, 6, 7, 8].map(
              (l) => `<a:p><a:pPr lvl="${String(l)}"/></a:p>`,
            ),
          }),
        ],
      },
      {
        shapes: [
          spXml({ name: 'l-title', ph: 'type="title" idx="0"' }),
          spXml({ name: 'l-body', ph: 'type="body" idx="1"' }),
        ],
      },
      { shapes: masterShapes() },
    );
    const byShape: Record<string, string> = {
      'builtin-title': 's-title',
      'builtin-body': 's-body',
      'builtin-plain': 's-plain',
    };
    expect(cases.builtinCases.length).toBe(27);
    for (const probe of cases.builtinCases) {
      const [prefix, index] = [probe.id.slice(0, probe.id.lastIndexOf('-')), probe.id.slice(-1)];
      const name = byShape[prefix];
      if (name === undefined) throw new Error(`unexpected builtin probe ${probe.id}`);
      const shape = shapeNamed(sheets.slide, name);
      const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
      const paragraph = shape.text?.paragraphs[Number(index)];
      if (paragraph === undefined) throw new Error(`${probe.id} has no paragraph`);
      expect(resolveSize(context, paragraph, undefined).value / 100).toBe(probe.measured.sizePt);
      expect(resolveMarginLeft(context, paragraph).value / 12700).toBe(probe.measured.leftIndentPt);
      expect(resolveIndent(context, paragraph).value / 12700).toBe(
        probe.measured.firstLineIndentPt,
      );
    }
  });

  it('is substituted wholesale, and never merged under a p:txStyles that exists', () => {
    // The distinction the `bucket-partial` probe was built for: a master that
    // declares p:txStyles gets no per-property backstop from the built-in
    // styles, so a bodyStyle declaring only marL leaves the size at the floor
    // and not at the built-in body's 28 points.
    const probe = cases.bucketCases.find((c) => c.id === 'bucket-partial-ph');
    if (probe?.situation?.kind !== 'bucket') throw new Error('bucket-partial-ph is missing');
    const built = buildBucket(probe.situation);
    expect(resolvePt(built)).toBe(probe.measured.sizePt);
    expect(resolvePt(built)).toBe(TEXT_FLOOR.sz / 100);
    expect(resolvePt(built)).not.toBe((BUILTIN_TEXT_STYLES.body[0]?.sz ?? 0) / 100);
    expect(floorOf(built.context, 0)).toBe(TEXT_FLOOR);
  });

  it('names the origin builtin rather than schemaDefault when it applies', () => {
    const sheets = chain(
      { shapes: [spXml({ name: 'probe', ph: 'type="body" idx="1"' })] },
      { shapes: [spXml({ name: 'l-body', ph: 'type="body" idx="1"' })] },
      { shapes: masterShapes() },
    );
    const shape = shapeNamed(sheets.slide, 'probe');
    const context: TextContext = { sheet: sheets.slide, shape, defaultTextStyle: undefined };
    const size = resolveSize(context, firstParagraph(sheets.slide, 'probe'), undefined);
    expect(size.origin).toBe('builtin');
    expect(size.value).toBe(BUILTIN_TEXT_STYLES.body[0]?.sz);
  });
});

describe('parsing', () => {
  it('keeps an absent size absent rather than defaulting it', () => {
    const style = listStyle(lvlPr(1, '<a:defRPr b="1"/>'));
    expect(style.levels[0]?.defRPr?.sz).toBeUndefined();
    expect(style.levels[0]?.defRPr?.b).toBe(true);
    expect(style.levels[0]?.marL).toBeUndefined();
    expect(style.levels[0]?.indent).toBeUndefined();
  });

  it('reads all four lexical forms of ST_OnOff and refuses a fifth', () => {
    for (const [written, value] of [
      ['1', true],
      ['true', true],
      ['0', false],
      ['false', false],
    ] as const) {
      expect(listStyle(lvlPr(1, `<a:defRPr b="${written}"/>`)).levels[0]?.defRPr?.b).toBe(value);
    }
    expect(() => listStyle(lvlPr(1, '<a:defRPr b="yes"/>'))).toThrow(ModelError);
  });

  it('keeps @charset signed, because Shift-JIS is -128', () => {
    const style = listStyle(
      lvlPr(1, '<a:defRPr><a:ea typeface="MS Gothic" charset="-128" pitchFamily="49"/></a:defRPr>'),
    );
    expect(style.levels[0]?.defRPr?.ea?.charset).toBe(-128);
    expect(style.levels[0]?.defRPr?.ea?.pitchFamily).toBe(49);
  });

  it('refuses an a:fld with no @id, rather than inventing a ST_Guid', () => {
    expect(() =>
      chain(
        {
          shapes: [
            spXml({ name: 'probe', paragraphs: ['<a:p><a:fld><a:t>3</a:t></a:fld></a:p>'] }),
          ],
        },
        {},
        {},
      ),
    ).toThrow(ModelError);
  });

  it('reads a run, a break and a field as three different things', () => {
    const sheets = chain(
      {
        shapes: [
          spXml({
            name: 'probe',
            paragraphs: [
              '<a:p><a:r><a:rPr lang="en-US"/><a:t>one</a:t></a:r><a:br/>' +
                '<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum">' +
                '<a:t>3</a:t></a:fld></a:p>',
            ],
          }),
        ],
      },
      {},
      {},
    );
    const content = firstParagraph(sheets.slide, 'probe').content;
    expect(content.map((c) => c.kind)).toEqual(['run', 'br', 'field']);
    expect(content[0]?.kind === 'run' ? content[0].text : '').toBe('one');
    expect(content[2]?.kind === 'field' ? content[2].fieldType : '').toBe('slidenum');
  });

  it('records that a master has no p:txStyles rather than inventing an empty one', () => {
    const withNone = parseSheetSpec('master', { shapes: masterShapes() });
    expect(withNone.txStyles).toBeUndefined();
    const withEmpty = parseSheetSpec('master', {
      shapes: masterShapes(),
      txStyles: '<p:txStyles/>',
    });
    expect(withEmpty.txStyles).toBeDefined();
    expect(withEmpty.txStyles?.body).toBeUndefined();
  });
});

/* -------------------------------------------------------------------------- */
/* the theme reference on a typeface (3.7)                                    */
/* -------------------------------------------------------------------------- */

describe('resolveTypeface', () => {
  const scheme: FontScheme = {
    name: 'T7',
    major: { latin: 'Georgia', ea: 'MS Gothic', cs: null },
    minor: { latin: 'Verdana', ea: '', cs: 'Leelawadee UI' },
  };

  it('sends +mj- to the major collection and +mn- to the minor one', () => {
    expect(resolveTypeface('+mj-lt', scheme)).toBe('Georgia');
    expect(resolveTypeface('+mn-lt', scheme)).toBe('Verdana');
    expect(resolveTypeface('+mj-ea', scheme)).toBe('MS Gothic');
    expect(resolveTypeface('+mn-cs', scheme)).toBe('Leelawadee UI');
  });

  it('leaves a real typeface name alone, including one that starts with a plus', () => {
    expect(resolveTypeface('Calibri', scheme)).toBe('Calibri');
    expect(resolveTypeface('+mn', scheme)).toBe('+mn');
  });

  it('is undefined when the collection entry is empty or absent', () => {
    expect(resolveTypeface('+mn-ea', scheme)).toBeUndefined();
    expect(resolveTypeface('+mj-cs', scheme)).toBeUndefined();
  });

  it('collects what a set of runs asks for, dropping what the theme leaves empty', () => {
    const font = (typeface: string): Typeface => ({
      typeface,
      panose: undefined,
      pitchFamily: undefined,
      charset: undefined,
    });
    expect(
      requestedTypefaces([font('+mj-lt'), font('Arial'), font('+mn-ea'), undefined], scheme),
    ).toEqual(['Georgia', 'Arial']);
  });
});

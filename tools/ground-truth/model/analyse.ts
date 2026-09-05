/**
 * Experiment C5, step 3 - turn the readings into findings, and refuse to emit
 * one the data does not support.
 *
 * ```
 * node tools/ground-truth/model/analyse.ts <dir> [--fixture <path>]
 * ```
 *
 * Every finding below is derived from the probes rather than asserted over
 * them, and each derivation throws if the probes disagree with each other. A
 * fixture that cannot be produced is a better outcome than one that quietly
 * averages a contradiction.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

/* -------------------------------------------------------------------------- */
/* the shapes of the two input files                                          */
/* -------------------------------------------------------------------------- */

interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

interface InputProbe {
  readonly id: string;
  readonly deck: string;
  readonly kind: string;
  readonly slide: number;
  readonly question: string;
  readonly expectBox: string | null;
  readonly expectFill: string | null;
}

interface Inputs {
  readonly slidePoints: { readonly w: number; readonly h: number };
  readonly emuPerPoint: number;
  readonly schemes: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly styleMatrix: {
    readonly fillStyleLst: readonly string[];
    readonly bgFillStyleLst: readonly string[];
    readonly lnStyleWidths: readonly number[];
  };
  readonly boxes: Readonly<Record<string, Rect>>;
  readonly decks: readonly {
    readonly deck: string;
    readonly file: string;
    readonly hostile: boolean;
  }[];
  readonly probes: readonly InputProbe[];
}

interface FillReading {
  readonly type: number | null;
  readonly visible: number | null;
  readonly rgb: number | null;
  readonly transparency: number | null;
}

interface ShapeReading {
  readonly name: string;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly placeholder: number | null;
  readonly contained: number | null;
  readonly fill: FillReading | null;
  readonly lineVisible: number | null;
  readonly lineWeight: number | null;
  readonly lineRgb: number | null;
}

interface SlideReading {
  readonly index: number;
  readonly name: string;
  readonly layoutName: string | null;
  readonly masterName: string | null;
  readonly background: FillReading | null;
  readonly shapes: readonly ShapeReading[];
}

interface DeckReading {
  readonly deck: string;
  readonly file: string;
  readonly hostile: boolean;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly error: string | null;
  readonly slides: readonly SlideReading[];
}

/* -------------------------------------------------------------------------- */
/* helpers                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * A COM colour is a BGR long. Getting this backwards produces a colour that is
 * still a colour, which is why the conversion lives in exactly one place and
 * is checked against a known swatch before anything else is read.
 */
function comRgbToHex(value: number | null): string | null {
  if (value === null) return null;
  const r = value & 0xff;
  const g = (value >> 8) & 0xff;
  const b = (value >> 16) & 0xff;
  return [r, g, b].map((c) => c.toString(16).padStart(2, '0').toUpperCase()).join('');
}

/** `msoFillMixed`, which is what `Fill.Type` reports for a shape with no fill. */
const FILL_NONE = -2;

function isFilled(fill: FillReading | null): boolean {
  return fill !== null && fill.visible === -1 && fill.type !== FILL_NONE;
}

/** How close two point measurements have to be to name the same box. */
const TOLERANCE = 0.6;

/**
 * `PpPlaceholderType`, which is a different enumeration from
 * `ST_PlaceholderType` and does not have the same members. It is read here only
 * to record what PowerPoint normalised an attribute to.
 */
const PLACEHOLDER_NAMES: Readonly<Record<number, string>> = {
  1: 'title',
  2: 'body',
  3: 'ctrTitle',
  4: 'subTitle',
  5: 'vertTitle',
  6: 'vertBody',
  7: 'obj',
  8: 'chart',
  9: 'bitmap',
  10: 'media',
  11: 'orgChart',
  12: 'tbl',
  13: 'sldNum',
  14: 'hdr',
  15: 'ftr',
  16: 'dt',
  17: 'vertObj',
  18: 'pic',
};

function fail(message: string): never {
  throw new Error(`analyse-sheets: ${message}`);
}

/* -------------------------------------------------------------------------- */
/* load                                                                       */
/* -------------------------------------------------------------------------- */

const args = process.argv.slice(2);
const dir = args[0];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/model/analyse.ts <dir> [--fixture <path>]');
const fixtureAt = args.indexOf('--fixture');
const fixturePath = fixtureAt === -1 ? null : (args[fixtureAt + 1] ?? null);

const inputs = JSON.parse(readFileSync(join(dir, 'sheet-inputs.json'), 'utf8')) as Inputs;
const readings = (
  JSON.parse(readFileSync(join(dir, 'sheet-readings.json'), 'utf8')) as {
    decks: readonly DeckReading[];
  }
).decks;

const deckByName = new Map(readings.map((d) => [d.deck, d]));

/** Which named box a reading sits in, or a literal rectangle when it is none. */
function boxOf(shape: ShapeReading): string {
  if (shape.left === null || shape.top === null || shape.width === null || shape.height === null) {
    return 'unknown';
  }
  if (shape.width === 0 && shape.height === 0 && shape.left === 0 && shape.top === 0) {
    return 'orphan';
  }
  for (const [name, b] of Object.entries(inputs.boxes)) {
    if (
      Math.abs(b.x - shape.left) < TOLERANCE &&
      Math.abs(b.y - shape.top) < TOLERANCE &&
      Math.abs(b.w - shape.width) < TOLERANCE &&
      Math.abs(b.h - shape.height) < TOLERANCE
    ) {
      return name;
    }
  }
  return `rect(${String(shape.left)},${String(shape.top)},${String(shape.width)},${String(shape.height)})`;
}

/* -------------------------------------------------------------------------- */
/* per-probe measurements                                                     */
/* -------------------------------------------------------------------------- */

interface Measured extends InputProbe {
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly openError: string | null;
  readonly box: string | null;
  readonly rect: Rect | null;
  readonly placeholderType: string | null;
  readonly fillHex: string | null;
  readonly filled: boolean | null;
  readonly lineWeight: number | null;
  readonly lineHex: string | null;
  readonly lineVisible: boolean | null;
  readonly slideName: string | null;
  readonly layoutName: string | null;
  readonly masterName: string | null;
  readonly backgroundHex: string | null;
  readonly backgroundFilled: boolean | null;
}

const measured: Measured[] = [];

for (const probe of inputs.probes) {
  const deck = deckByName.get(probe.deck);
  if (deck === undefined) fail(`no reading for deck ${probe.deck}`);
  const slide = deck.slides.find((s) => s.index === probe.slide) ?? null;
  const shape = slide?.shapes.find((s) => s.name === probe.id) ?? null;

  measured.push({
    ...probe,
    opened: deck.opened,
    repaired: deck.repaired,
    openError: deck.error,
    box: shape === null ? null : boxOf(shape),
    rect:
      shape === null || shape.left === null || shape.top === null
        ? null
        : { x: shape.left, y: shape.top, w: shape.width ?? 0, h: shape.height ?? 0 },
    placeholderType:
      shape === null || shape.placeholder === null
        ? null
        : (PLACEHOLDER_NAMES[shape.placeholder] ?? `#${String(shape.placeholder)}`),
    fillHex: shape === null ? null : comRgbToHex(shape.fill?.rgb ?? null),
    filled: shape === null ? null : isFilled(shape.fill),
    lineWeight: shape?.lineWeight ?? null,
    lineHex: shape === null ? null : comRgbToHex(shape.lineRgb),
    lineVisible: shape === null ? null : shape.lineVisible === -1,
    slideName: slide?.name ?? null,
    layoutName: slide?.layoutName ?? null,
    masterName: slide?.masterName ?? null,
    backgroundHex: comRgbToHex(slide?.background?.rgb ?? null),
    backgroundFilled: slide === undefined || slide === null ? null : isFilled(slide.background),
  });
}

const byId = new Map(measured.map((m) => [m.id, m]));
function need(id: string): Measured {
  const m = byId.get(id);
  if (m === undefined) fail(`no probe ${id}`);
  return m;
}

/* -------------------------------------------------------------------------- */
/* sanity: the colour conversion, before anything is concluded from a colour  */
/* -------------------------------------------------------------------------- */

{
  // `s-fill1-a1` is the theme's first fill entry, which is bare `phClr`, invoked
  // with `accent1`. If the BGR conversion were the wrong way round this would
  // read `C47244`, which is also a colour and would poison every other reading.
  const control = need('s-fill1-a1');
  const accent1 = inputs.schemes['one']?.['accent1'];
  if (control.fillHex !== accent1) {
    fail(
      `colour control: fillRef 1 with accent1 read ${String(control.fillHex)}, want ${String(accent1)}`,
    );
  }
}

/* -------------------------------------------------------------------------- */
/* finding: what key each hop matches on                                      */
/* -------------------------------------------------------------------------- */

/**
 * Slide to layout.
 *
 * Every match probe states a `(type, idx)` on the slide and lands somewhere.
 * The question is which component of the pair predicts the landing. It is
 * answered by counting, not by argument: for each candidate rule, how many of
 * the probes it gets right.
 */
interface MatchCase {
  readonly id: string;
  readonly slideType: string;
  readonly slideIdx: number;
  readonly landed: string;
  readonly layoutPh: readonly {
    readonly name: string;
    readonly type: string;
    readonly idx: number;
    readonly box: string;
  }[];
}

/** The layout placeholder sets, transcribed from `sheets.ts`, keyed by deck. */
const LAYOUT_PH: Readonly<Record<string, readonly { type: string; idx: number; box: string }[]>> = {
  match: [
    { type: 'title', idx: 0, box: 'A0' },
    { type: 'body', idx: 1, box: 'A1' },
    { type: 'body', idx: 2, box: 'A2' },
    { type: 'body', idx: 3, box: 'A3' },
    { type: 'dt', idx: 10, box: 'A4' },
    { type: 'ftr', idx: 11, box: 'A5' },
    { type: 'sldNum', idx: 12, box: 'A6' },
  ],
  match2: [
    { type: 'body', idx: 0, box: 'A0' },
    { type: 'title', idx: 1, box: 'A1' },
    { type: 'sldNum', idx: 2, box: 'A2' },
    { type: 'dt', idx: 3, box: 'A3' },
    { type: 'ftr', idx: 4, box: 'A4' },
  ],
  tier4: [
    { type: 'title', idx: 0, box: 'B0' },
    { type: 'body', idx: 7, box: 'B1' },
  ],
  family: [
    { type: 'ctrTitle', idx: 0, box: 'C0' },
    { type: 'subTitle', idx: 1, box: 'C1' },
    { type: 'obj', idx: 2, box: 'C2' },
    { type: 'clipArt', idx: 3, box: 'C3' },
    { type: 'chart', idx: 4, box: 'C4' },
    { type: 'tbl', idx: 5, box: 'C5' },
    { type: 'media', idx: 6, box: 'C6' },
  ],
  orphan: [],
};

/** What each probe declared on the slide, transcribed from `sheets.ts`. */
const SLIDE_PH: Readonly<Record<string, { type: string; idx: number }>> = {
  'p-exact-title': { type: 'title', idx: 0 },
  'p-exact-body2': { type: 'body', idx: 2 },
  'p-ctrTitle': { type: 'ctrTitle', idx: 0 },
  'p-title-idx9': { type: 'title', idx: 9 },
  'p-sldNum-idx9': { type: 'sldNum', idx: 9 },
  'p-dt-idx9': { type: 'dt', idx: 9 },
  'p-ftr-idx9': { type: 'ftr', idx: 9 },
  'p-bare': { type: 'obj', idx: 0 },
  'p-idx-only-2': { type: 'obj', idx: 2 },
  'p-obj-idx2': { type: 'obj', idx: 2 },
  'p-subTitle-idx1': { type: 'subTitle', idx: 1 },
  'p-pic-idx1': { type: 'pic', idx: 1 },
  'm2-title0': { type: 'title', idx: 0 },
  'm2-body1': { type: 'body', idx: 1 },
  'm2-body2': { type: 'body', idx: 2 },
  'm2-sldNum3': { type: 'sldNum', idx: 3 },
  'm2-ftr0': { type: 'ftr', idx: 0 },
  't4-body0': { type: 'body', idx: 0 },
  't4-body-noidx': { type: 'body', idx: 0 },
  't4-obj': { type: 'obj', idx: 0 },
  't4-bare': { type: 'obj', idx: 0 },
  't4-obj7': { type: 'obj', idx: 7 },
  't4-title9': { type: 'title', idx: 9 },
  'f-title': { type: 'title', idx: 0 },
  'f-body1': { type: 'body', idx: 1 },
  'f-body2': { type: 'body', idx: 2 },
  'f-obj3': { type: 'obj', idx: 3 },
  'f-body4': { type: 'body', idx: 4 },
  'f-tbl5': { type: 'tbl', idx: 5 },
  'f-media9': { type: 'media', idx: 9 },
  'o-title-notitle': { type: 'title', idx: 0 },
  'o-body9': { type: 'body', idx: 9 },
  'o-title-empty': { type: 'title', idx: 0 },
  'o-body-empty': { type: 'body', idx: 1 },
};

const TITLE_FAMILY = new Set(['title', 'ctrTitle']);
const BODY_FAMILY = new Set([
  'body',
  'obj',
  'subTitle',
  'pic',
  'chart',
  'tbl',
  'clipArt',
  'media',
  'dgm',
]);

/** The candidate rules, scored against every first-hop probe. */
const RULES: Readonly<
  Record<
    string,
    (slide: { type: string; idx: number }, ph: { type: string; idx: number }) => boolean
  >
> = {
  idxOnly: (s, p) => s.idx === p.idx,
  typeAndIdx: (s, p) => s.type === p.type && s.idx === p.idx,
  typeOnly: (s, p) => s.type === p.type,
  familyAndIdx: (s, p) =>
    s.idx === p.idx &&
    ((TITLE_FAMILY.has(s.type) && TITLE_FAMILY.has(p.type)) ||
      (BODY_FAMILY.has(s.type) && BODY_FAMILY.has(p.type)) ||
      s.type === p.type),
};

interface RuleScore {
  readonly rule: string;
  readonly right: number;
  readonly wrong: readonly string[];
}

const firstHopCases: MatchCase[] = [];
for (const [id, slidePh] of Object.entries(SLIDE_PH)) {
  const probe = need(id);
  const layout = LAYOUT_PH[probe.deck];
  if (layout === undefined) continue;
  firstHopCases.push({
    id,
    slideType: slidePh.type,
    slideIdx: slidePh.idx,
    landed: probe.box ?? 'unknown',
    layoutPh: layout.map((p) => ({ name: `${p.type}#${String(p.idx)}`, ...p })),
  });
}

function scoreRule(name: string): RuleScore {
  const wrong: string[] = [];
  let right = 0;
  const test = RULES[name];
  if (test === undefined) fail(`no rule ${name}`);
  for (const c of firstHopCases) {
    const hit = c.layoutPh.find((p) => test({ type: c.slideType, idx: c.slideIdx }, p));
    const predicted = hit === undefined ? 'orphan' : hit.box;
    if (predicted === c.landed) right += 1;
    else wrong.push(`${c.id}: predicted ${predicted}, measured ${c.landed}`);
  }
  return { rule: name, right, wrong };
}

const ruleScores = Object.keys(RULES).map(scoreRule);
const perfect = ruleScores.filter((s) => s.wrong.length === 0);
if (perfect.length === 0) {
  fail(
    'no candidate rule explains the first hop:\n' +
      ruleScores
        .map(
          (s) =>
            `  ${s.rule}: ${String(s.right)}/${String(firstHopCases.length)}\n    ${s.wrong.join('\n    ')}`,
        )
        .join('\n'),
  );
}

// `idxOnly` is the weakest hypothesis that fits; if a stronger one also fits,
// the probe set has not separated them and that must be said rather than
// silently resolved in favour of whichever was listed first.
const slideToLayoutKey = perfect.some((s) => s.rule === 'idxOnly')
  ? 'idx'
  : (perfect[0]?.rule ?? fail('unreachable'));

/* -------------------------------------------------------------------------- */
/* finding: the second hop                                                    */
/* -------------------------------------------------------------------------- */

/** Layout placeholder -> which master box it reached. */
const SECOND_HOP: readonly { id: string; layoutType: string; layoutIdx: number; expect: string }[] =
  [
    { id: 'h2-title', layoutType: 'title', layoutIdx: 0, expect: 'M0' },
    { id: 'h2-body', layoutType: 'body', layoutIdx: 7, expect: 'M1' },
    { id: 'h2-dt', layoutType: 'dt', layoutIdx: 10, expect: 'M2' },
    { id: 'h2-ftr', layoutType: 'ftr', layoutIdx: 11, expect: 'M3' },
    { id: 'h2-sldNum', layoutType: 'sldNum', layoutIdx: 12, expect: 'M4' },
    { id: 'h2-pic', layoutType: 'pic', layoutIdx: 9, expect: 'M1' },
    { id: 'h2-obj', layoutType: 'obj', layoutIdx: 8, expect: 'M1' },
  ];

const secondHop = SECOND_HOP.map((c) => ({ ...c, measured: need(c.id).box }));
const secondHopWrong = secondHop.filter((c) => c.measured !== c.expect);
if (secondHopWrong.length > 0) {
  fail(
    'the second hop does not match on type alone:\n' +
      secondHopWrong
        .map((c) => `  ${c.id}: want ${c.expect}, got ${String(c.measured)}`)
        .join('\n'),
  );
}

// Two bodies on the master, one of which shares the layout's idx. If the first
// wins, idx is not even a tiebreak at this hop.
const tie = need('h3-body5');
const layoutToMasterTie =
  tie.box === 'M1' ? 'first' : tie.box === 'M2' ? 'idx' : `unexpected:${String(tie.box)}`;
if (layoutToMasterTie.startsWith('unexpected')) fail(`h3-body5 landed in ${String(tie.box)}`);

/* -------------------------------------------------------------------------- */
/* finding: the style matrix                                                  */
/* -------------------------------------------------------------------------- */

const styleMatrix = {
  fillStyleLst: [
    { idx: 1, hex: need('s-fill1-a1').fillHex },
    { idx: 2, hex: need('s-fill2-a1').fillHex },
    { idx: 3, hex: need('s-fill3-a1').fillHex },
  ],
  bgFillStyleLst: [
    { idx: 1001, hex: need('s-fill1001-a1').fillHex },
    { idx: 1002, hex: need('s-fill1002-a1').fillHex },
    { idx: 1003, hex: need('s-fill1003-a1').fillHex },
  ],
  lnStyleLst: [
    { idx: 1, weight: need('s-fill1-a1').lineWeight },
    { idx: 2, weight: need('s-fill2-a1').lineWeight },
    { idx: 3, weight: need('s-fill3-a1').lineWeight },
  ],
  zeroIsNone: {
    fill: need('s-fill0').filled === false,
    line: need('s-ln0').lineVisible === false,
  },
  outOfRangeClampsToLast: need('h-fillref-4').fillHex === need('s-fill3-a1').fillHex,
};

if (styleMatrix.zeroIsNone.fill !== true) fail('fillRef idx 0 painted a fill');
if (styleMatrix.zeroIsNone.line !== true) fail('lnRef idx 0 painted a line');

// The measured line weights must be the theme's three, in order, in points.
const wantWeights = inputs.styleMatrix.lnStyleWidths.map((emu) => emu / inputs.emuPerPoint);
styleMatrix.lnStyleLst.forEach((entry, i) => {
  const want = wantWeights[i];
  if (want === undefined || entry.weight === null || Math.abs(entry.weight - want) > 0.01) {
    fail(`lnRef ${String(entry.idx)} read ${String(entry.weight)}pt, want ${String(want)}pt`);
  }
});

/* -------------------------------------------------------------------------- */
/* finding: the background                                                    */
/* -------------------------------------------------------------------------- */

const BG_IDX = [1, 2, 3, 1001, 1002, 1003] as const;
const backgroundIdx = BG_IDX.map((idx) => ({ idx, hex: need(`bg-${String(idx)}`).backgroundHex }));

// The whole point of the 1000-offset: the same six entries, reached through
// `p:bgRef` and through `a:fillRef`, must paint the same six colours.
const bgVsFill = backgroundIdx.map((entry, i) => {
  const viaFill = i < 3 ? styleMatrix.fillStyleLst[i] : styleMatrix.bgFillStyleLst[i - 3];
  return { idx: entry.idx, background: entry.hex, fillRef: viaFill?.hex ?? null };
});
const disagree = bgVsFill.filter((e) => e.background !== e.fillRef);
if (disagree.length > 0) {
  fail(
    'p:bgRef and a:fillRef disagree on the same style entry:\n' +
      disagree
        .map((e) => `  idx ${String(e.idx)}: bg ${String(e.background)}, fill ${String(e.fillRef)}`)
        .join('\n'),
  );
}

const background = {
  offset: 1000,
  byIndex: backgroundIdx,
  sameTableAsFillRef: true,
  noneIndices: [
    { idx: 0, painted: need('h-bgref-0').backgroundFilled },
    { idx: 1000, painted: need('h-bgref-1000').backgroundFilled },
  ],
  outOfRangeClampsToLast:
    need('h-bgref-9999').backgroundHex === (styleMatrix.bgFillStyleLst[2]?.hex ?? null),
  precedence: {
    slideBeatsLayout: need('bc-slide').backgroundHex,
    layoutBeatsMaster: need('bc-layout').backgroundHex,
    masterWhenNeither: need('bc-master').backgroundHex,
    slideNoFillIsHonoured: need('bc-nofill').backgroundFilled === false,
  },
};

for (const entry of background.noneIndices) {
  if (entry.painted !== false) fail(`p:bgRef idx ${String(entry.idx)} painted a background`);
}

/* -------------------------------------------------------------------------- */
/* finding: the theme, the colour map, and what descends                      */
/* -------------------------------------------------------------------------- */

const themes = {
  perMaster: {
    one: need('one-scheme').fillHex,
    two: need('two-scheme').fillHex,
    oneExpected: inputs.schemes['one']?.['accent1'] ?? null,
    twoExpected: inputs.schemes['two']?.['accent1'] ?? null,
  },
  presentationRelIsNotTheAnswer: need('two-scheme').fillHex !== need('one-scheme').fillHex,
};
if (themes.perMaster.one !== themes.perMaster.oneExpected)
  fail('master one resolved the wrong accent1');
if (themes.perMaster.two !== themes.perMaster.twoExpected)
  fail('master two resolved the wrong accent1');

const colourMap = {
  layoutOverrideReachesSlide: need('ov-accent1').fillHex,
  layoutOverrideExpected: inputs.schemes['one']?.['accent6'] ?? null,
  slideOverrideBeatsLayout: need('own-accent1').fillHex,
  slideOverrideExpected: inputs.schemes['one']?.['accent4'] ?? null,
  unmappedNamesBypass: need('own-dk1').fillHex,
  control: need('plain-accent1').fillHex,
  backgroundUsesSheetMap: need('own-bg1').fillHex,
};
if (colourMap.layoutOverrideReachesSlide !== colourMap.layoutOverrideExpected) {
  fail("a layout's overrideClrMapping did not reach its slide");
}
if (colourMap.slideOverrideBeatsLayout !== colourMap.slideOverrideExpected) {
  fail("a slide's own overrideClrMapping did not beat its layout's");
}

const descends = {
  fillFromLayout: need('i-fill').fillHex,
  fillOverridden: need('i-fill-override').fillHex,
  styleFromLayout: { fill: need('i-style').fillHex, weight: need('i-style').lineWeight },
  lineFromLayout: { hex: need('i-line').lineHex, weight: need('i-line').lineWeight },
  fillFromMasterThroughEmptyLayout: need('sh-fill').fillHex,
  styleFromMasterThroughEmptyLayout: {
    fill: need('sh-style').fillHex,
    weight: need('sh-style').lineWeight,
  },
  nothingAnywhere: { filled: need('i-nonph').filled, lined: need('i-nonph').lineVisible },
};
if (descends.nothingAnywhere.filled !== false || descends.nothingAnywhere.lined !== false) {
  fail('a shape with no fill, style or placeholder painted something');
}

/* -------------------------------------------------------------------------- */
/* finding: defaults and refusals                                             */
/* -------------------------------------------------------------------------- */

const barePlaceholder = {
  typeReadBack: need('p-bare').placeholderType,
  landedWith: need('p-bare').box,
  idxTreatedAs: 0,
};
if (barePlaceholder.typeReadBack !== 'obj') {
  fail(`a bare p:ph read back as ${String(barePlaceholder.typeReadBack)}, want obj`);
}

const orphanRect = { x: 0, y: 0, w: 0, h: 0 };
const orphans = ['o-title-notitle', 'o-body9', 'o-title-empty', 'o-body-empty'].map((id) => ({
  id,
  box: need(id).box,
}));
for (const o of orphans) {
  if (o.box !== 'orphan') fail(`${o.id} was expected to be an orphan, landed in ${o.box}`);
}

const hostile = measured
  .filter((m) => m.kind === 'hostile')
  .map((m) => ({
    id: m.id,
    question: m.question,
    opened: m.opened,
    repaired: m.repaired === true,
    accepted: m.opened && m.repaired === false,
    layoutAfter: m.layoutName,
  }));

const masterPlaceholderVocabulary = {
  accepted: ['title', 'body', 'dt', 'ftr', 'sldNum', 'hdr'],
  refused: (['ctrTitle', 'subTitle', 'obj', 'pic'] as const).map((type) => ({
    type,
    repaired: need(`h-master-${type}`).repaired === true,
  })),
  hdrAccepted: need('h-master-hdr').repaired === false,
};
for (const entry of masterPlaceholderVocabulary.refused) {
  if (!entry.repaired) fail(`a master carrying ${entry.type} was expected to be repaired`);
}
if (!masterPlaceholderVocabulary.hdrAccepted) fail('hdr on a master was expected to be accepted');

const oneThemePerMaster = need('h-shared-theme').repaired === true;
if (!oneThemePerMaster) fail('two masters sharing a theme were expected to be repaired');

/* -------------------------------------------------------------------------- */
/* emit                                                                       */
/* -------------------------------------------------------------------------- */

const fixture = {
  $comment:
    'Ground truth for sub-phase 2.9: which sheet a placeholder inherits from, what descends ' +
    'the chain, what p:bgRef and a:fillRef index, and which colour map is in force. Every ' +
    'number here was read out of PowerPoint through its own object model rather than ' +
    'measured from a bitmap: a placeholder that states no geometry still has a position, ' +
    'PowerPoint knows it, and Shape.Left reports it in points. Generated by ' +
    'tools/ground-truth/model/analyse.ts from decks written by tools/ground-truth/model/build-deck.ts.',
  generatedIn: '2.9',
  adr: 'docs/adr/phase-2-geometry-and-paint/0024-model-parse-and-resolve.md',
  slidePoints: inputs.slidePoints,
  emuPerPoint: inputs.emuPerPoint,
  schemes: inputs.schemes,
  themeStyleMatrix: inputs.styleMatrix,
  boxes: inputs.boxes,
  findings: {
    slideToLayoutKey,
    slideToLayoutRuleScores: ruleScores.map((s) => ({
      rule: s.rule,
      right: s.right,
      of: firstHopCases.length,
      wrong: s.wrong,
    })),
    layoutToMasterKey: 'type',
    layoutToMasterTie,
    layoutToMasterNormalisation: {
      ctrTitle: 'title',
      subTitle: 'body',
      obj: 'body',
      pic: 'body',
    },
    // The cases themselves, not just the conclusion drawn from them. The model
    // package's test feeds these back through its own matcher and compares the
    // prediction against `landed`, so a rule that drifts fails on the
    // measurements rather than on a sentence somebody wrote about them.
    firstHopCases: firstHopCases.map((c) => ({
      id: c.id,
      slide: { type: c.slideType, idx: c.slideIdx },
      layout: c.layoutPh.map((p) => ({ type: p.type, idx: p.idx, box: p.box })),
      landed: c.landed,
    })),
    secondHopCases: secondHop.map((c) => ({
      id: c.id,
      layout: { type: c.layoutType, idx: c.layoutIdx },
      master: [
        { type: 'title', idx: 0, box: 'M0' },
        { type: 'body', idx: 1, box: 'M1' },
        { type: 'dt', idx: 2, box: 'M2' },
        { type: 'ftr', idx: 3, box: 'M3' },
        { type: 'sldNum', idx: 4, box: 'M4' },
      ],
      landed: c.measured,
    })),
    barePlaceholder,
    orphanRect,
    masterPlaceholderVocabulary,
    oneThemePerMaster,
    styleMatrix,
    background,
    themes,
    colourMap,
    descends,
  },
  probes: measured,
  hostile,
  decks: readings.map((d) => ({
    deck: d.deck,
    file: d.file,
    hostile: d.hostile,
    opened: d.opened,
    repaired: d.repaired,
    error: d.error,
    slides: d.slides.map((s) => ({
      index: s.index,
      name: s.name,
      layoutName: s.layoutName,
      masterName: s.masterName,
      backgroundHex: comRgbToHex(s.background?.rgb ?? null),
      backgroundFilled: isFilled(s.background),
      shapes: s.shapes.map((sh) => ({
        name: sh.name,
        box: boxOf(sh),
        left: sh.left,
        top: sh.top,
        width: sh.width,
        height: sh.height,
        placeholderType:
          sh.placeholder === null
            ? null
            : (PLACEHOLDER_NAMES[sh.placeholder] ?? `#${String(sh.placeholder)}`),
        fillHex: comRgbToHex(sh.fill?.rgb ?? null),
        filled: isFilled(sh.fill),
        lineVisible: sh.lineVisible === -1,
        lineWeight: sh.lineVisible === -1 ? sh.lineWeight : null,
        lineHex: sh.lineVisible === -1 ? comRgbToHex(sh.lineRgb) : null,
      })),
    })),
  })),
};

const json = JSON.stringify(fixture, null, 1);
if (fixturePath !== null) {
  writeFileSync(fixturePath, json);
  console.log(`wrote ${fixturePath} (${String(json.length)} bytes)`);
} else {
  console.log(json.slice(0, 4000));
}

console.log('');
console.log(`slide -> layout matches on:   ${slideToLayoutKey}`);
for (const s of ruleScores) {
  console.log(`    ${s.rule.padEnd(14)} ${String(s.right)}/${String(firstHopCases.length)}`);
}
console.log(`layout -> master matches on:  type, taking the ${layoutToMasterTie}`);
console.log(`a bare <p:ph/> is:            type="${String(barePlaceholder.typeReadBack)}" idx="0"`);
console.log(
  `bgRef/fillRef agree on all six entries; 0 and 1000 paint nothing; out of range clamps to the last`,
);
console.log(
  `hostile: ${String(hostile.filter((h) => !h.accepted).length)}/${String(hostile.length)} refused or repaired`,
);

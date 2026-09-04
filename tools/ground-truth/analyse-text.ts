/**
 * Experiment T1, step 3 - score every reading of the cascade against what
 * PowerPoint did, and emit `corpus/ground-truth/text-cascade.json`.
 *
 * ```
 * node tools/ground-truth/analyse-text.ts <work-dir>
 * ```
 *
 * ## The scoring, not the confirming
 *
 * The plan names nine sources and an order for them. Confirming that order is
 * worth almost nothing - the ladder was built so that a *wrong* order predicts a
 * different number at some rung, and the useful output is which wrong orders it
 * refutes and by how much. So:
 *
 * - **The order** is scored over all 40320 permutations of the eight named
 *   sources, against the nine placeholder rungs. The report says how many fit,
 *   which sources have to be dropped before any permutation fits at all, and
 *   which single permutation survives once they are.
 * - **The bucket rule** is scored over the eight combinations of the three
 *   independent choices anyone could make: whose `@type` selects it, whether the
 *   header-and-footer trio reads `p:otherStyle`, and whether a shape with no
 *   bucket falls through to `p:defaultTextStyle`.
 *
 * Nothing is written unless one model fits every probe of its kind exactly.
 * A near-miss is a finding to report, not a fixture to commit.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { Situation, Source, TextProbe } from './text-cascade.ts';

/* -------------------------------------------------------------------------- */
/* what the reader wrote                                                      */
/* -------------------------------------------------------------------------- */

interface ParagraphReading {
  readonly text: string | null;
  readonly size: number | null;
  readonly font: string | null;
  readonly bold: number | null;
  readonly italic: number | null;
  readonly underline: number | null;
  readonly strike: number | null;
  readonly alignment: number | null;
  readonly indentLevel: number | null;
  readonly leftIndent: number | null;
  readonly firstLineIndent: number | null;
  readonly spaceBefore: number | null;
  readonly spaceAfter: number | null;
  readonly spaceWithin: number | null;
  readonly lineRuleBefore: number | null;
  readonly lineRuleAfter: number | null;
  readonly lineRuleWithin: number | null;
  readonly bulletVisible: number | null;
  readonly bulletCharacter: number | null;
  readonly bulletFont: string | null;
  readonly bulletType: number | null;
}

interface ShapeReading {
  readonly name: string;
  readonly placeholder: number | null;
  readonly paragraphs: readonly ParagraphReading[];
  readonly legacyFonts: readonly (string | null)[];
}

interface SlideReading {
  readonly index: number;
  readonly shapes: readonly ShapeReading[];
}

interface DeckReading {
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly slides: readonly SlideReading[];
}

interface Inputs {
  readonly ladderPt: Readonly<Record<Source, number>>;
  readonly bucketPt: { readonly title: number; readonly body: number; readonly other: number };
  readonly themeFonts: { readonly major: string; readonly minor: string };
  readonly sources: readonly Source[];
  readonly probes: readonly TextProbe[];
}

const workDir = process.argv[2];
if (workDir === undefined) throw new Error('usage: analyse-text.ts <work-dir>');

const inputs = JSON.parse(readFileSync(join(workDir, 'text-inputs.json'), 'utf8')) as Inputs;
const readings = JSON.parse(readFileSync(join(workDir, 'text-readings.json'), 'utf8')) as {
  readonly decks: readonly DeckReading[];
};

const byDeck = new Map(readings.decks.map((d) => [d.deck, d]));

/* -------------------------------------------------------------------------- */
/* joining a probe to its reading                                             */
/* -------------------------------------------------------------------------- */

interface Measured {
  readonly probe: TextProbe;
  readonly repaired: boolean;
  readonly reading: ParagraphReading | null;
  readonly legacyFont: string | null;
}

function measure(probe: TextProbe): Measured {
  const deck = byDeck.get(probe.deck);
  if (deck === undefined) throw new Error(`${probe.id}: no reading for deck ${probe.deck}`);
  if (deck.repaired === true || !deck.opened) {
    return { probe, repaired: true, reading: null, legacyFont: null };
  }
  const slide = deck.slides.find((s) => s.index === probe.slide);
  if (slide === undefined) throw new Error(`${probe.id}: deck ${probe.deck} has no slide`);
  const shape = slide.shapes.find((s) => s.name === probe.shape);
  if (shape === undefined) {
    throw new Error(
      `${probe.id}: no shape "${probe.shape}" on ${probe.deck} slide ${String(probe.slide)}`,
    );
  }
  const reading = shape.paragraphs[probe.paragraph - 1];
  if (reading === undefined) {
    throw new Error(
      `${probe.id}: shape "${probe.shape}" has no paragraph ${String(probe.paragraph)}`,
    );
  }
  return {
    probe,
    repaired: false,
    reading,
    legacyFont: shape.legacyFonts[probe.paragraph - 1] ?? null,
  };
}

const measured = inputs.probes.map(measure);
const usable = measured.filter((m) => !m.repaired);
const excluded = measured.filter((m) => m.repaired);

const sizeOf = (m: Measured): number => {
  const size = m.reading?.size;
  if (size === null || size === undefined) throw new Error(`${m.probe.id}: no size was read`);
  return size;
};

/* -------------------------------------------------------------------------- */
/* the built-in text styles                                                   */
/* -------------------------------------------------------------------------- */

/**
 * What PowerPoint substitutes when the master declares no `p:txStyles`.
 *
 * Read out of the `builtin` deck rather than typed in, so the fixture cannot
 * drift from the measurement it claims to record.
 */
interface BuiltinLevel {
  readonly level: number;
  readonly szHundredths: number;
  readonly marLEmu: number;
  readonly indentEmu: number;
  readonly bullet: boolean;
}

const EMU_PER_POINT = 12700;

function builtinLevels(shape: string, idPrefix: string): readonly BuiltinLevel[] {
  return usable
    .filter((m) => m.probe.deck === 'builtin' && m.probe.shape === shape)
    .sort((a, b) => a.probe.paragraph - b.probe.paragraph)
    .map((m) => {
      const r = m.reading;
      if (r === null) throw new Error(`${idPrefix}: no reading`);
      return {
        level: m.probe.paragraph,
        szHundredths: Math.round(sizeOf(m) * 100),
        marLEmu: Math.round((r.leftIndent ?? 0) * EMU_PER_POINT),
        indentEmu: Math.round((r.firstLineIndent ?? 0) * EMU_PER_POINT),
        bullet: r.bulletVisible === -1,
      };
    });
}

const builtin = {
  title: builtinLevels('s-title', 'builtin-title'),
  body: builtinLevels('s-body', 'builtin-body'),
  other: builtinLevels('s-plain', 'builtin-plain'),
};

const builtinBodyFirst = builtin.body[0];
if (builtinBodyFirst === undefined) throw new Error('the builtin deck produced no body levels');
const BUILTIN_BODY_L1: number = builtinBodyFirst.szHundredths;

/**
 * The size a shape lands on when every source is silent.
 *
 * Distinct from the built-in `p:txStyles` above: that is substituted wholesale
 * when the master declares none, whereas this is what remains when the master
 * *does* declare `p:txStyles` and the bucket it selects says nothing. Measured
 * on `bucket-partial`, where a body placeholder whose `p:bodyStyle` declares
 * only `marL` reported neither the built-in body size nor the package default.
 */
const floorProbe = usable.find((m) => m.probe.id === 'bucket-partial-ph');
if (floorProbe === undefined) throw new Error('bucket-partial-ph is missing');
const FLOOR_HUNDREDTHS = Math.round(sizeOf(floorProbe) * 100);

/* -------------------------------------------------------------------------- */
/* scoring the order                                                          */
/* -------------------------------------------------------------------------- */

const SOURCES = inputs.sources;

function permutations<T>(items: readonly T[]): readonly (readonly T[])[] {
  if (items.length <= 1) return [items];
  const out: T[][] = [];
  for (let i = 0; i < items.length; i++) {
    const rest = [...items.slice(0, i), ...items.slice(i + 1)];
    const head = items[i];
    if (head === undefined) continue;
    for (const tail of permutations(rest)) out.push([head, ...tail]);
  }
  return out;
}

/**
 * What a rung reports under a candidate order, for a placeholder.
 *
 * The rung declares a size at every source in its situation; a candidate order
 * picks the first of them it can see. Anything it cannot see is skipped, and a
 * rung where it can see nothing lands on the built-in body style - which is what
 * "no source spoke" looks like from the object model.
 */
function predictLadder(
  situation: Extract<Situation, { kind: 'ladder' }>,
  order: readonly Source[],
  reads: ReadonlySet<Source>,
): number {
  for (const source of order) {
    if (!reads.has(source)) continue;
    if (source === 'txStyles') {
      if (situation.txStyles !== null) return situation.txStyles.body;
      continue;
    }
    if (source === 'defaultText') {
      if (situation.defaultTextStyle !== null) return situation.defaultTextStyle;
      continue;
    }
    if (source === 'objDefaults') {
      if (situation.objectDefaults !== null) return situation.objectDefaults;
      continue;
    }
    const declared = situation.declares[source];
    if (declared !== undefined) return declared;
  }
  return BUILTIN_BODY_L1;
}

const ladderPh = usable.filter(
  (m) => m.probe.situation?.kind === 'ladder' && m.probe.situation.shape === 'ph',
);
if (ladderPh.length !== SOURCES.length + 1) {
  throw new Error(
    `expected ${String(SOURCES.length + 1)} placeholder rungs, got ${String(ladderPh.length)}`,
  );
}

function fitsLadder(order: readonly Source[], reads: ReadonlySet<Source>): number {
  let hits = 0;
  for (const m of ladderPh) {
    const situation = m.probe.situation;
    if (situation?.kind !== 'ladder') continue;
    if (predictLadder(situation, order, reads) === Math.round(sizeOf(m) * 100)) hits += 1;
  }
  return hits;
}

const ALL_ORDERS = permutations(SOURCES);

const allSources = new Set(SOURCES);
const ordersFittingWithAll = ALL_ORDERS.filter(
  (order) => fitsLadder(order, allSources) === ladderPh.length,
);

// Which sources a placeholder actually reads: the largest subset for which some
// permutation fits every rung. Enumerated rather than assumed, so "a placeholder
// ignores `p:defaultTextStyle`" is a result and not a premise.
let bestReads: ReadonlySet<Source> | null = null;
let bestOrders: readonly (readonly Source[])[] = [];
for (let mask = (1 << SOURCES.length) - 1; mask >= 0; mask--) {
  const reads = new Set(SOURCES.filter((_, i) => (mask & (1 << i)) !== 0));
  const fitting = ALL_ORDERS.filter((order) => fitsLadder(order, reads) === ladderPh.length);
  if (fitting.length === 0) continue;
  if (bestReads === null || reads.size > bestReads.size) {
    bestReads = reads;
    bestOrders = fitting;
  }
}
if (bestReads === null) throw new Error('no order over any subset of the sources fits the ladder');

// Every permutation that fits differs only in where it puts a source the
// placeholder never reads, so the *read* prefix has to be unique or the
// experiment has not answered its own question.
const readOrders = new Set(
  bestOrders.map((order) => order.filter((s) => bestReads.has(s)).join(' > ')),
);
if (readOrders.size !== 1) {
  throw new Error(`the ladder does not pin a single order: ${[...readOrders].join(' | ')}`);
}
const MEASURED_ORDER = [...bestReads].sort(
  (a, b) => (bestOrders[0]?.indexOf(a) ?? 0) - (bestOrders[0]?.indexOf(b) ?? 0),
);
const NOT_SOURCES = SOURCES.filter((s) => !bestReads.has(s));

/* -------------------------------------------------------------------------- */
/* scoring the bucket rule                                                    */
/* -------------------------------------------------------------------------- */

type BucketName = 'title' | 'body' | 'other' | null;

interface BucketModel {
  /** Whose `@type` selects the bucket. */
  readonly from: 'slide' | 'chainEnd';
  /** Whether `dt`, `ftr`, `sldNum` and `hdr` read `p:otherStyle`. */
  readonly hfReadsOther: boolean;
  /** Whether a shape that is not a placeholder reads `p:otherStyle`. */
  readonly plainReadsOther: boolean;
}

const HF = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
const TITLE = new Set(['title', 'ctrTitle']);

function bucketFor(
  situation: Extract<Situation, { kind: 'bucket' }>,
  model: BucketModel,
): BucketName {
  const type =
    model.from === 'slide'
      ? situation.slideType
      : (situation.masterType ?? situation.layoutType ?? situation.slideType);
  if (type === null) return model.plainReadsOther ? 'other' : null;
  if (TITLE.has(type)) return 'title';
  if (HF.has(type)) return model.hfReadsOther ? 'other' : null;
  return 'body';
}

function predictBucket(
  situation: Extract<Situation, { kind: 'bucket' }>,
  model: BucketModel,
): number {
  const bucket = bucketFor(situation, model);
  if (bucket !== null && situation.txStyles !== null) {
    const declared = situation.txStyles[bucket];
    // A bucket that exists and declares nothing ends the walk at the floor
    // rather than handing on - measured on `bucket-partial`.
    return declared === 0 ? FLOOR_HUNDREDTHS : declared;
  }
  if (situation.defaultTextStyle !== null) return situation.defaultTextStyle;
  return FLOOR_HUNDREDTHS;
}

const bucketProbes = usable.filter((m) => m.probe.situation?.kind === 'bucket');

const BUCKET_MODELS: readonly (BucketModel & { readonly name: string })[] = [
  { name: 'slide/hf-other/plain-other', from: 'slide', hfReadsOther: true, plainReadsOther: true },
  { name: 'slide/hf-other/plain-none', from: 'slide', hfReadsOther: true, plainReadsOther: false },
  { name: 'slide/hf-none/plain-other', from: 'slide', hfReadsOther: false, plainReadsOther: true },
  { name: 'slide/hf-none/plain-none', from: 'slide', hfReadsOther: false, plainReadsOther: false },
  {
    name: 'chainEnd/hf-other/plain-other',
    from: 'chainEnd',
    hfReadsOther: true,
    plainReadsOther: true,
  },
  {
    name: 'chainEnd/hf-other/plain-none',
    from: 'chainEnd',
    hfReadsOther: true,
    plainReadsOther: false,
  },
  {
    name: 'chainEnd/hf-none/plain-other',
    from: 'chainEnd',
    hfReadsOther: false,
    plainReadsOther: true,
  },
  {
    name: 'chainEnd/hf-none/plain-none',
    from: 'chainEnd',
    hfReadsOther: false,
    plainReadsOther: false,
  },
];

const bucketScores = BUCKET_MODELS.map((model) => {
  let hits = 0;
  const misses: string[] = [];
  for (const m of bucketProbes) {
    const situation = m.probe.situation;
    if (situation?.kind !== 'bucket') continue;
    const predicted = predictBucket(situation, model);
    if (predicted === Math.round(sizeOf(m) * 100)) hits += 1;
    else misses.push(m.probe.id);
  }
  return { name: model.name, model, hits, of: bucketProbes.length, misses };
});

const bucketWinners = bucketScores.filter((s) => s.hits === s.of);
if (bucketWinners.length !== 1) {
  const table = bucketScores.map((s) => `${s.name} ${String(s.hits)}/${String(s.of)}`).join('\n  ');
  throw new Error(
    `${String(bucketWinners.length)} bucket models fit every probe; exactly one must:\n  ${table}`,
  );
}
const BUCKET_WINNER = bucketWinners[0];
if (BUCKET_WINNER === undefined) throw new Error('unreachable: a winner with no entry');
const BUCKET_MODEL: BucketModel = BUCKET_WINNER.model;

/* -------------------------------------------------------------------------- */
/* one model, scored against every situation at once                          */
/* -------------------------------------------------------------------------- */

/**
 * The whole cascade, as one function, in the terms the two scorings settled.
 *
 * The two halves above each answered their own question with the other held
 * fixed, which is how they were made decisive; this puts them back together and
 * checks the result against every probe that carries a situation. A model that
 * fits each half separately and not the two together is a model that has been
 * fitted twice rather than measured once.
 */
function predict(situation: Situation): number {
  if (situation.kind === 'ladder') {
    for (const source of ['run', 'para', 'shapeLst'] as const) {
      const declared = situation.declares[source];
      if (declared !== undefined) return declared;
    }
    if (situation.shape === 'ph') {
      for (const source of ['layoutPh', 'masterPh'] as const) {
        const declared = situation.declares[source];
        if (declared !== undefined) return declared;
      }
      // A body placeholder ends at p:bodyStyle whatever it says, and at the
      // built-in body style when the master declares no p:txStyles at all.
      if (situation.txStyles === null) return BUILTIN_BODY_L1;
      return situation.txStyles.body === 0 ? FLOOR_HUNDREDTHS : situation.txStyles.body;
    }
    if (situation.defaultTextStyle !== null) return situation.defaultTextStyle;
    return FLOOR_HUNDREDTHS;
  }
  return predictBucket(situation, BUCKET_MODEL);
}

const situationProbes = usable.filter((m) => m.probe.situation !== undefined);
const unifiedMisses = situationProbes.filter((m) => {
  const situation = m.probe.situation;
  if (situation === undefined) return false;
  return predict(situation) !== Math.round(sizeOf(m) * 100);
});
if (unifiedMisses.length > 0) {
  const lines = unifiedMisses
    .slice(0, 10)
    .map((m) => {
      const situation = m.probe.situation;
      const predicted = situation === undefined ? NaN : predict(situation) / 100;
      return `  ${m.probe.id}: predicted ${String(predicted)}pt, measured ${String(sizeOf(m))}pt`;
    })
    .join('\n');
  throw new Error(
    `the two halves do not compose: ${String(unifiedMisses.length)} of ` +
      `${String(situationProbes.length)} situations miss\n${lines}`,
  );
}

/* -------------------------------------------------------------------------- */
/* the remaining findings, each one probe deep                                */
/* -------------------------------------------------------------------------- */

const find = (id: string): Measured => {
  const m = usable.find((x) => x.probe.id === id);
  if (m === undefined) throw new Error(`probe ${id} is missing or came from a repaired deck`);
  return m;
};

const readingOf = (id: string): ParagraphReading => {
  const r = find(id).reading;
  if (r === null) throw new Error(`probe ${id} has no reading`);
  return r;
};

/** Per-property merge, or does the nearest level that speaks take everything. */
const merge = readingOf('merge-per-property');
const MERGES_PER_PROPERTY =
  Math.round((merge.size ?? 0) * 100) === inputs.bucketPt.body * 100 &&
  merge.bold === -1 &&
  merge.italic === -1 &&
  (merge.underline ?? 0) !== 0 &&
  (merge.strike ?? 0) !== 0;

/** `@marL` and `@indent` when nothing in the cascade declares them. */
const marginPlain = readingOf('margin-plain');
const marginBare = readingOf('margin-bare');
const marginSplit = readingOf('margin-split');

/** The two spellings of `ST_Percentage`, and the one that is not a percentage. */
const pctNumeric = readingOf('percent-numeric');
const pctStrict = readingOf('percent-strict');
const pctPoints = readingOf('percent-points');

/** The negatives: three elements that turn out not to be sources at all. */
const negatives = {
  objectDefaults: {
    probe: 'order-8-objDefaults-plain',
    measuredPt: sizeOf(find('order-8-objDefaults-plain')),
    declaredPt: inputs.ladderPt.objDefaults,
  },
  defPPrOnShape: {
    probe: 'defppr-shape',
    measuredPt: sizeOf(find('defppr-shape')),
    declaredPt: 30,
  },
  defPPrOnPackage: {
    probe: 'defppr-package',
    measuredPt: sizeOf(find('defppr-package')),
    declaredPt: 26,
  },
  otherStyleOnSlide: { probe: 'otherstyle-slide', measuredPt: sizeOf(find('otherstyle-slide')) },
  otherStyleOnLayout: { probe: 'otherstyle-layout', measuredPt: sizeOf(find('otherstyle-layout')) },
  otherStyleOnMaster: { probe: 'otherstyle-master', measuredPt: sizeOf(find('otherstyle-master')) },
};

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const root = fileURLToPath(new URL('../..', import.meta.url));

const caseOf = (m: Measured): Record<string, unknown> => ({
  id: m.probe.id,
  kind: m.probe.kind,
  question: m.probe.question,
  ...(m.probe.situation === undefined ? {} : { situation: m.probe.situation }),
  measured: {
    sizePt: m.reading?.size ?? null,
    font: m.reading?.font ?? null,
    legacyFont: m.legacyFont,
    bold: m.reading?.bold ?? null,
    italic: m.reading?.italic ?? null,
    underline: m.reading?.underline ?? null,
    strike: m.reading?.strike ?? null,
    indentLevel: m.reading?.indentLevel ?? null,
    leftIndentPt: m.reading?.leftIndent ?? null,
    firstLineIndentPt: m.reading?.firstLineIndent ?? null,
    spaceWithin: m.reading?.spaceWithin ?? null,
    lineRuleWithin: m.reading?.lineRuleWithin ?? null,
    spaceBefore: m.reading?.spaceBefore ?? null,
    lineRuleBefore: m.reading?.lineRuleBefore ?? null,
    bulletVisible: m.reading?.bulletVisible ?? null,
  },
});

const fixture = {
  $comment:
    'Ground truth for sub-phase 3.1: which of the nine claimed sources of a text property PowerPoint actually reads, in what order, and which of them it ignores entirely. Every number here was read out of PowerPoint through its own object model rather than measured from a bitmap: a run whose a:rPr declares nothing still has a resolved size, PowerPoint knows it, and Font.Size reports it in points. Generated by tools/ground-truth/analyse-text.ts from decks written by build-text-deck.ts.',
  generatedIn: '3.1',
  adr: 'docs/adr/0027-the-text-cascade.md',
  emuPerPoint: EMU_PER_POINT,
  ladderPt: inputs.ladderPt,
  bucketPt: inputs.bucketPt,
  themeFonts: inputs.themeFonts,

  model: {
    /**
     * The walk, nearest first. Every level contributes only the properties it
     * declares, and the walk ends at whichever terminus the shape reaches.
     */
    walk: ['runRPr', 'paraDefRPr', 'shapeLstStyle', 'ancestorPlaceholderLstStyle'],
    /**
     * Where the walk ends, and it is one of two places rather than a rung of
     * one ladder: a shape whose chain ends at a type with a bucket reads that
     * bucket and stops, and one with no bucket reads p:defaultTextStyle and
     * stops. Measured: a body placeholder ignores a p:defaultTextStyle that is
     * right there, and a sldNum placeholder reads it.
     */
    terminus: { withBucket: 'txStyles', withoutBucket: 'defaultTextStyle' },
    /** The order the ladder pinned for a placeholder, as source names. */
    placeholderOrder: MEASURED_ORDER,
    /** Named as sources in the plan, and refuted: PowerPoint reads none of them. */
    notSources: [...NOT_SOURCES.filter((s) => s !== 'defaultText'), 'defPPr', 'otherStyle'],
    bucket: {
      /** Whose `@type` selects the bucket: the shape's own, or the chain's end. */
      from: BUCKET_WINNER.model.from,
      hfReadsOtherStyle: BUCKET_WINNER.model.hfReadsOther,
      nonPlaceholderReadsOtherStyle: BUCKET_WINNER.model.plainReadsOther,
      titleTypes: ['title', 'ctrTitle'],
      noBucketTypes: ['dt', 'ftr', 'sldNum', 'hdr'],
    },
    /** The size, margin and indent left when every source is silent. */
    floor: { szHundredths: FLOOR_HUNDREDTHS, marLEmu: 0, indentEmu: 0 },
    /** Substituted wholesale when the master declares no p:txStyles at all. */
    builtinTextStyles: builtin,
    mergesPerProperty: MERGES_PER_PROPERTY,
    percentageSpellings: ['150000', '150%'],
  },

  scores: {
    ladder: {
      rungs: ladderPh.length,
      permutationsOverAllSources: ALL_ORDERS.length,
      fittingWithEverySourceRead: ordersFittingWithAll.length,
      sourcesRead: MEASURED_ORDER,
      sourcesIgnored: NOT_SOURCES,
      fittingOnceIgnoredSourcesAreDropped: bestOrders.length,
      distinctReadOrdersFitting: readOrders.size,
    },
    bucket: bucketScores.map((s) => ({
      model: s.name,
      hits: s.hits,
      of: s.of,
      misses: s.misses.slice(0, 6),
    })),
    /** The two halves put back together and re-scored as one function. */
    unified: { hits: situationProbes.length - unifiedMisses.length, of: situationProbes.length },
  },

  findings: {
    ladderCases: ladderPh.map(caseOf),
    ladderPlainCases: usable
      .filter((m) => m.probe.situation?.kind === 'ladder' && m.probe.situation.shape === 'plain')
      .map(caseOf),
    bucketCases: bucketProbes.map(caseOf),
    levelCases: usable.filter((m) => m.probe.kind === 'level').map(caseOf),
    builtinCases: usable.filter((m) => m.probe.kind === 'builtin').map(caseOf),
    hopCases: usable.filter((m) => m.probe.kind === 'hop').map(caseOf),
    mergeCases: usable.filter((m) => m.probe.kind === 'merge').map(caseOf),
    marginCases: usable.filter((m) => m.probe.kind === 'margin').map(caseOf),
    percentCases: usable.filter((m) => m.probe.kind === 'percent').map(caseOf),
    fontCases: usable.filter((m) => m.probe.kind === 'font').map(caseOf),
    negatives,
    margins: {
      nothingDeclaresEither: {
        leftIndentPt: marginPlain.leftIndent,
        firstLineIndentPt: marginPlain.firstLineIndent,
      },
      marLDeclaredIndentNot: {
        leftIndentPt: marginBare.leftIndent,
        firstLineIndentPt: marginBare.firstLineIndent,
      },
      marLAndIndentFromDifferentSheets: {
        leftIndentPt: marginSplit.leftIndent,
        firstLineIndentPt: marginSplit.firstLineIndent,
      },
      schemaDefaults: { marL: 347663, indent: -342900 },
    },
    percentage: {
      numeric: { spaceWithin: pctNumeric.spaceWithin, lineRuleWithin: pctNumeric.lineRuleWithin },
      strict: { spaceWithin: pctStrict.spaceWithin, lineRuleWithin: pctStrict.lineRuleWithin },
      points: { spaceWithin: pctPoints.spaceWithin, lineRuleWithin: pctPoints.lineRuleWithin },
    },
  },

  /** Packages PowerPoint rewrote rather than opened, excluded from every score. */
  repaired: [...new Set(excluded.map((m) => m.probe.deck))],
  excludedProbes: excluded.map((m) => m.probe.id),

  probes: usable.map(caseOf),
};

const out = join(root, 'corpus', 'ground-truth', 'text-cascade.json');
writeFileSync(out, `${JSON.stringify(fixture, null, 2)}\n`);

/* -------------------------------------------------------------------------- */

console.log(
  `probes            ${String(usable.length)} usable, ${String(excluded.length)} excluded`,
);
console.log(`repaired decks    ${fixture.repaired.join(', ') || 'none'}`);
console.log('');
console.log(`ladder            ${String(ladderPh.length)} rungs`);
console.log(
  `  permutations    ${String(ALL_ORDERS.length)} over ${String(SOURCES.length)} sources; ` +
    `${String(ordersFittingWithAll.length)} fit with every source read`,
);
console.log(`  sources read    ${MEASURED_ORDER.join(' > ')}`);
console.log(`  sources ignored ${NOT_SOURCES.join(', ') || 'none'}`);
console.log(
  `  once dropped    ${String(bestOrders.length)} permutation(s) fit, ` +
    `${String(readOrders.size)} distinct read order(s)`,
);
console.log('');
console.log('bucket rule');
for (const s of bucketScores) {
  const mark = s.hits === s.of ? '*' : ' ';
  console.log(`  ${mark} ${s.name.padEnd(30)} ${String(s.hits)}/${String(s.of)}`);
}
console.log('');
console.log(
  `unified model       ${String(situationProbes.length - unifiedMisses.length)}/${String(situationProbes.length)} situations`,
);
console.log(`merge per property  ${String(MERGES_PER_PROPERTY)}`);
console.log(
  `floor               ${String(FLOOR_HUNDREDTHS / 100)}pt, marL 0, indent 0 ` +
    `(the schema says ${String(347663 / EMU_PER_POINT)}pt and ${String(-342900 / EMU_PER_POINT)}pt)`,
);
console.log(
  `builtin body        ${builtin.body.map((l) => String(l.szHundredths / 100)).join(', ')}`,
);
console.log(`wrote               ${out}`);

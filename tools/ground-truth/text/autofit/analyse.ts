/**
 * Experiment T4, step 4 - score the candidate models and emit the fixture.
 *
 * ```
 * node tools/ground-truth/text/autofit/analyse.ts <work-dir>
 * ```
 *
 * Reads `autofit-inputs.json`, `autofit-readings.json` and the decks PowerPoint
 * saved, and writes `corpus/ground-truth/autofit.json` only if exactly one
 * candidate fits every case of each question. A model that fits all but one is
 * not the rule; the one is the finding, and in this experiment the one turned
 * out to be the whole of section C.
 *
 * ## Two instruments, pointed at the same rule
 *
 * The emitted `@fontScale` says which rung PowerPoint *chose*, and a rung is
 * chosen by a fit test, so every recompute probe is one observation of the
 * text-height model read through a decision. The view decks say what a scale
 * *does* - the distance between two baselines, and the height of the whole
 * block - with no decision in the way at all.
 *
 * The order matters. Sections C and D derive the height model from the view
 * decks, where it is measured by subtraction. Section E then predicts 3000-odd
 * rungs PowerPoint chose, using constants fitted to data that contains no rungs
 * at all. That is an out-of-sample test rather than a curve fit.
 *
 * ## What is deliberately not imported
 *
 * Nothing from `@pptx-studio/text`. Every model below is written from the
 * measurement; the package is written from the fixture this emits. An analysis
 * that imported the implementation would let a wrong implementation score
 * itself correct, and the tests derived from the fixture would agree with it
 * forever.
 *
 * The one thing this cannot score is the `wrap` deck, whose line count is a
 * consequence of the rung rather than a constant - predicting it needs the line
 * breaker, and the line breaker lives in the package this file must not import.
 * So `wrap` is checked here only for the half of the rule that needs no
 * breaker, and its readings are carried into the fixture for the package test
 * to replay in full.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import type { Insets, Probe, Spacing } from './probes.ts';
import { readZip } from '../../lib/zip.ts';

const dir = process.argv[2];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/text/autofit/analyse.ts <work-dir>');
const work = resolve(dir);

/* ------------------------------------------------------------------ inputs */

interface DeckInput {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
  readonly probes: number;
  readonly recompute: boolean;
}

interface LineReading {
  readonly index: number;
  readonly text: string | null;
  readonly top: number | null;
  readonly height: number | null;
  readonly length: number | null;
}

interface ShapeReading {
  readonly id: string;
  readonly shapeWidth: number | null;
  readonly shapeHeight: number | null;
  readonly autoSize: number | null;
  readonly wordWrap: number | null;
  readonly lineCount: number | null;
  readonly paraCount: number | null;
  readonly fontSize: number | null;
  readonly fontName: string | null;
  readonly boundTop: number | null;
  readonly boundHeight: number | null;
  readonly boundWidth: number | null;
  readonly lines: readonly LineReading[];
}

interface DeckReading {
  readonly deck: string;
  readonly savedAs: string | null;
  readonly recompute: boolean;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly error: string | null;
  readonly shapes: readonly ShapeReading[];
}

type FullProbe = Probe & { readonly paragraphCount: number };

const inputs = JSON.parse(readFileSync(join(work, 'autofit-inputs.json'), 'utf8')) as {
  readonly decks: readonly DeckInput[];
  readonly probes: readonly FullProbe[];
};
const readings = JSON.parse(readFileSync(join(work, 'autofit-readings.json'), 'utf8')) as {
  readonly decks: readonly DeckReading[];
};

const probeById = new Map(inputs.probes.map((p) => [p.id, p]));

/* ------------------------------------------------ what PowerPoint wrote */

interface Rung {
  /** Thousandths of a percent, as `@fontScale` states it. */
  readonly fontScale: number;
  /** Thousandths of a percent, as `@lnSpcReduction` states it. */
  readonly lnSpcReduction: number;
}

const rungKey = (r: Rung): string => `${String(r.fontScale)}/${String(r.lnSpcReduction)}`;

const SLIDE_PART = /^ppt\/slides\/slide\d+\.xml$/;

/** Split a saved slide into one chunk per `p:sp`, with the shape's name. */
function* shapesOf(file: string): Generator<{ name: string; xml: string }> {
  const decoder = new TextDecoder();
  for (const part of readZip(readFileSync(file))) {
    if (!SLIDE_PART.test(part.name)) continue;
    const xml = decoder.decode(part.bytes);
    for (const sp of xml.split('<p:sp>').slice(1)) {
      const name = /<p:cNvPr[^>]*\sname="([^"]*)"/.exec(sp);
      if (name?.[1] === undefined) continue;
      yield { name: name[1], xml: sp };
    }
  }
}

/**
 * Every `a:normAutofit` in a saved deck, by shape name.
 *
 * A missing `@fontScale` is 100% and a missing `@lnSpcReduction` is 0 - the
 * schema default, and also what PowerPoint writes when it decides the text
 * already fits.
 */
function emittedRungs(file: string): Map<string, Rung> {
  const out = new Map<string, Rung>();
  for (const { name, xml } of shapesOf(file)) {
    const auto = /<a:normAutofit([^/>]*)\/>/.exec(xml);
    if (auto?.[1] === undefined) continue;
    const scale = /fontScale="(\d+)"/.exec(auto[1]);
    const reduce = /lnSpcReduction="(\d+)"/.exec(auto[1]);
    out.set(name, {
      fontScale: scale?.[1] === undefined ? 100000 : Number(scale[1]),
      lnSpcReduction: reduce?.[1] === undefined ? 0 : Number(reduce[1]),
    });
  }
  return out;
}

/** Every shape's height in EMU, for the `spAutoFit` decks. */
function emittedHeights(file: string): Map<string, number> {
  const out = new Map<string, number>();
  for (const { name, xml } of shapesOf(file)) {
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
    if (ext?.[2] === undefined) continue;
    out.set(name, Number(ext[2]));
  }
  return out;
}

/* ------------------------------------------------------------------- cases */

interface Case {
  readonly probe: FullProbe;
  readonly deck: string;
  readonly observed: Rung;
  readonly lines: number;
  /** True when the line count is a constant of the probe rather than of the rung. */
  readonly fixedLines: boolean;
}

/** A view deck reading: what a stated scale and spacing actually drew. */
interface BlockCase {
  readonly probe: FullProbe;
  readonly advance: number;
  readonly blockHeight: number;
  readonly lines: number;
}

interface SpCase {
  readonly probe: FullProbe;
  readonly emu: number;
  readonly boxPt: number;
  readonly lines: number;
}

const cases: Case[] = [];
const blocks: BlockCase[] = [];
const spCases: SpCase[] = [];
const refused: string[] = [];
const repaired: string[] = [];
const lineCountMismatch: string[] = [];

const EMU_PER_POINT = 12700;

for (const deck of readings.decks) {
  if (!deck.opened) {
    refused.push(deck.deck);
    continue;
  }
  if (deck.repaired === true) repaired.push(deck.deck);

  const rungs = deck.savedAs === null ? null : emittedRungs(join(work, deck.savedAs));
  const heights = deck.savedAs === null ? null : emittedHeights(join(work, deck.savedAs));

  for (const shape of deck.shapes) {
    const probe = probeById.get(shape.id);
    if (probe === undefined) throw new Error(`${deck.deck}: shape ${shape.id} is not a probe`);
    const lines = shape.lineCount ?? 0;

    if (probe.expectLines !== undefined && lines !== probe.expectLines) {
      // A probe built so that nothing can wrap, that wrapped anyway, measures a
      // box too narrow rather than a ladder.
      lineCountMismatch.push(
        `${probe.id}: expected ${String(probe.expectLines)} lines, laid out ${String(lines)}`,
      );
    }

    if (probe.autofit === 'sp') {
      const emu = heights?.get(shape.id);
      if (emu === undefined) throw new Error(`${probe.id}: no saved height`);
      spCases.push({ probe, emu, boxPt: probe.rect.h, lines });
      continue;
    }

    if (probe.recompute) {
      if (rungs === null) throw new Error(`${deck.deck}: recompute deck was not saved`);
      const observed = rungs.get(shape.id);
      if (observed === undefined) throw new Error(`${probe.id}: no emitted a:normAutofit`);
      cases.push({
        probe,
        deck: deck.deck,
        observed,
        lines,
        fixedLines: probe.expectLines !== undefined,
      });
      continue;
    }

    const tops = shape.lines.map((l) => l.top).filter((t): t is number => t !== null);
    if (tops.length < 2 || shape.boundHeight === null) continue;
    let sum = 0;
    for (let i = 1; i < tops.length; i += 1) {
      const a = tops[i];
      const b = tops[i - 1];
      if (a === undefined || b === undefined) throw new Error('unreachable');
      sum += a - b;
    }
    blocks.push({
      probe,
      advance: sum / (tops.length - 1),
      blockHeight: shape.boundHeight,
      lines,
    });
  }
}

/* -------------------------------------------------------------------------- */
/* report plumbing                                                            */
/* -------------------------------------------------------------------------- */

const report: string[] = [];
const findings: Record<string, unknown> = {};

function say(line = ''): void {
  report.push(line);
  console.log(line);
}

interface Score {
  readonly name: string;
  readonly fits: number;
  readonly total: number;
  readonly misses: readonly string[];
}

function table(rows: readonly Score[]): void {
  const width = Math.max(...rows.map((r) => r.name.length));
  for (const r of rows) {
    say(`    ${r.name.padEnd(width)}  ${String(r.fits).padStart(5)} / ${String(r.total)}`);
  }
}

/** Refuse to emit a fixture unless exactly one candidate is perfect. */
function chooseUnique(what: string, rows: readonly Score[]): Score {
  const perfect = rows.filter((r) => r.fits === r.total);
  if (perfect.length !== 1) {
    say();
    for (const r of rows) {
      if (r.misses.length > 0) {
        say(`  ${r.name} misses:`);
        for (const m of r.misses) say(`    ${m}`);
      }
    }
    throw new Error(
      `${what}: expected exactly one candidate to fit every case, got ${String(perfect.length)}` +
        ` (${perfect.map((r) => r.name).join(', ') || 'none'})`,
    );
  }
  const winner = perfect[0];
  if (winner === undefined) throw new Error('unreachable');
  return winner;
}

say('T4 - autofit, measured against PowerPoint');
say(
  `  ${String(cases.length)} recompute probes, ${String(blocks.length)} block readings, ` +
    `${String(spCases.length)} spAutoFit probes`,
);
say();

/* ------------------------------------------------------- A: the packages */

say('A. Packages');
if (refused.length > 0) throw new Error(`refused on open: ${refused.join(', ')}`);
if (repaired.length > 0) {
  throw new Error(
    `repaired on open, so the measurement is of a repaired file: ${repaired.join(', ')}`,
  );
}
if (lineCountMismatch.length > 0) {
  for (const m of lineCountMismatch.slice(0, 10)) say(`    ${m}`);
  throw new Error(
    `${String(lineCountMismatch.length)} probes did not lay out the line count they were built for`,
  );
}
say(`    ${String(readings.decks.length)} decks opened, none repaired`);
say('    every probe laid out the line count it was built for');
say();

/* -------------------------------------------- B: the effective font size */

say('B. The effective font size');

/** T2's finding: the line box is a font-independent 1.2 x the size. */
const NATURAL = 1.2;

type SizeRule = (nominalPt: number, scale: number) => number;

const SIZE_RULES: Readonly<Record<string, SizeRule>> = {
  /**
   * Round to a whole point, halves up - but only when a scale actually
   * applies. A `@fontScale` of 100% leaves the size alone, fraction and all.
   */
  roundWholeWhenScaled: (sz, s) => (s === 1 ? sz : Math.round(sz * s)),
  /** The same, applied unconditionally, so 10.5pt becomes 11pt for nothing. */
  roundWholeAlways: (sz, s) => Math.round(sz * s),
  roundHalfEven: (sz, s) => {
    if (s === 1) return sz;
    const x = sz * s;
    const f = Math.floor(x);
    if (Math.abs(x - f - 0.5) > 1e-9) return Math.round(x);
    return f % 2 === 0 ? f : f + 1;
  },
  exact: (sz, s) => sz * s,
  floorWhole: (sz, s) => (s === 1 ? sz : Math.floor(sz * s)),
  ceilWhole: (sz, s) => (s === 1 ? sz : Math.ceil(sz * s)),
  roundHalfPoint: (sz, s) => (s === 1 ? sz : Math.round(sz * s * 2) / 2),
  roundQuarterPoint: (sz, s) => (s === 1 ? sz : Math.round(sz * s * 4) / 4),
};

/** Round half up to a whole percent, which T2 measured on `a:spcPct`. */
function quantisePercent(thousandths: number): number {
  return Math.round(thousandths / 1000) / 100;
}

/**
 * The smallest line spacing PowerPoint will use, as a fraction of the line box.
 *
 * Measured on five sizes: a 15% spacing reduced by 20% is -5%, and PowerPoint
 * lays it out at one per cent of the line box rather than at zero or at
 * something negative. 8pt gives 0.095pt a line where 1% of the box is 0.096,
 * and 40pt gives 0.480 where 1% is 0.480.
 */
const MIN_SPACING_PERCENT = 0.01;

/**
 * The advance between two baselines under a stated scale.
 *
 * Measured on the view decks: `a:lnSpc` in exact points is used as given and
 * ignores `@lnSpcReduction` entirely, and a percentage has the reduction
 * *subtracted from it* rather than multiplied into the result - which is the
 * same number at 100% and a 4% difference at 150%.
 */
type AdvanceRule = (effPt: number, reduction: number, lnSpc: Spacing | undefined) => number;

const ADVANCE_RULES: Readonly<Record<string, AdvanceRule>> = {
  /** Subtract first, quantise the difference, then clamp at one per cent. */
  subtractThenQuantise: (eff, red, lnSpc) => {
    if (lnSpc?.kind === 'points') return lnSpc.value / 100;
    const pct = quantisePercent((lnSpc?.value ?? 100000) - red);
    return Math.max(MIN_SPACING_PERCENT, pct) * NATURAL * eff;
  },
  /** Quantise the stated percentage first, then subtract - which differs
   *  whenever either value is a fraction of a percent. */
  quantiseThenSubtract: (eff, red, lnSpc) => {
    if (lnSpc?.kind === 'points') return lnSpc.value / 100;
    const pct = quantisePercent(lnSpc?.value ?? 100000) - red / 100000;
    return Math.max(MIN_SPACING_PERCENT, pct) * NATURAL * eff;
  },
  /** No clamp: a reduction larger than the spacing stacks the lines. */
  subtractUnclamped: (eff, red, lnSpc) => {
    if (lnSpc?.kind === 'points') return lnSpc.value / 100;
    return quantisePercent((lnSpc?.value ?? 100000) - red) * NATURAL * eff;
  },
  /** The reading anyone writes first: the reduction scales the advance. */
  multiply: (eff, red, lnSpc) => {
    const k = 1 - red / 100000;
    if (lnSpc?.kind === 'points') return (lnSpc.value / 100) * k;
    return (
      Math.max(MIN_SPACING_PERCENT, quantisePercent(lnSpc?.value ?? 100000) * k) * NATURAL * eff
    );
  },
  /** Subtracted, but exact point spacing is reduced too. */
  subtractAndReducePoints: (eff, red, lnSpc) => {
    if (lnSpc?.kind === 'points') return (lnSpc.value / 100) * (1 - red / 100000);
    const pct = quantisePercent((lnSpc?.value ?? 100000) - red);
    return Math.max(MIN_SPACING_PERCENT, pct) * NATURAL * eff;
  },
};

const sizeScores: Score[] = [];
for (const [name, rule] of Object.entries(SIZE_RULES)) {
  let fits = 0;
  const misses: string[] = [];
  for (const b of blocks) {
    if (b.probe.autofit !== 'norm') continue;
    const nominal = b.probe.sz / 100;
    const eff = rule(nominal, (b.probe.stored.fontScale ?? 100000) / 100000);
    const predicted = ADVANCE_RULES['subtractThenQuantise']?.(
      eff,
      b.probe.stored.lnSpcReduction ?? 0,
      b.probe.lnSpc,
    );
    if (predicted === undefined) throw new Error('unreachable');
    if (Math.abs(predicted - b.advance) < 0.005) fits += 1;
    else if (misses.length < 8) {
      misses.push(
        `${b.probe.id} (${String(nominal)}pt x ${String(b.probe.stored.fontScale ?? 100000)}) ` +
          `predicted ${predicted.toFixed(4)} measured ${b.advance.toFixed(4)}`,
      );
    }
  }
  const total = blocks.filter((b) => b.probe.autofit === 'norm').length;
  sizeScores.push({ name, fits, total, misses });
}
table(sizeScores);
const sizeWinner = chooseUnique('the effective font size', sizeScores);
say();
say(`    ${sizeWinner.name}: ${String(sizeWinner.fits)} / ${String(sizeWinner.total)}`);
say();

const advanceScores: Score[] = [];
for (const [name, rule] of Object.entries(ADVANCE_RULES)) {
  let fits = 0;
  const misses: string[] = [];
  const subset = blocks.filter((b) => b.probe.autofit === 'norm');
  for (const b of subset) {
    const nominal = b.probe.sz / 100;
    const eff = SIZE_RULES['roundWholeWhenScaled']?.(
      nominal,
      (b.probe.stored.fontScale ?? 100000) / 100000,
    );
    if (eff === undefined) throw new Error('unreachable');
    const predicted = rule(eff, b.probe.stored.lnSpcReduction ?? 0, b.probe.lnSpc);
    if (Math.abs(predicted - b.advance) < 0.005) fits += 1;
    else if (misses.length < 8) {
      misses.push(
        `${b.probe.id} predicted ${predicted.toFixed(4)} measured ${b.advance.toFixed(4)}`,
      );
    }
  }
  advanceScores.push({ name, fits, total: subset.length, misses });
}
say('   the advance:');
table(advanceScores);
const advanceWinner = chooseUnique('the advance', advanceScores);
say();

/**
 * A rule out of one of the tables above, by the name a scorer chose.
 *
 * A function rather than an index-and-check, because narrowing does not reach
 * into the bodies of the scoring functions below - and a `!` there would be a
 * non-null assertion this repository does not allow.
 */
function ruleNamed<T>(table: Readonly<Record<string, T>>, name: string): T {
  const rule = table[name];
  if (rule === undefined) throw new Error(`no rule named ${name}`);
  return rule;
}

const effSize = ruleNamed(SIZE_RULES, sizeWinner.name);
const advanceOf = ruleNamed(ADVANCE_RULES, advanceWinner.name);

/* ---------------------------------------------------- C: the last line */

say('C. The height of the last line');

/**
 * The term sub-phase 3.2 could not measure and 3.6 was going to.
 *
 * A block of `n` lines is `(n-1)` advances plus the height of the last line,
 * and at single spacing those are the same number - which is why 3.2 could
 * fit its whole table without ever separating them. Under a stated `a:lnSpc`
 * they come apart, and the `last` deck reads the difference off directly:
 * `BoundHeight` is the block, the line tops give the advance, and one
 * subtraction gives the last line.
 */
interface LastLineSample {
  readonly face: string;
  readonly eff: number;
  readonly advance: number;
  readonly lastLine: number;
  readonly id: string;
  /**
   * Whether this sample can go into the fit.
   *
   * The rule has two branches and only one of them is a straight line, so a
   * sample sitting on the floor would drag the fit off the line it is supposed
   * to be measuring. The percentage-spacing rows are provably on the linear
   * branch for any coefficient above 0.18 - which every face turns out to be -
   * so those are fitted and the exact-point rows are held back to check the
   * result. One of the three exact-point rows per face does sit on the floor:
   * 32pt text with a 30pt line spacing draws its last line at exactly 30pt.
   */
  readonly fittable: boolean;
}

const lastSamples: LastLineSample[] = [];
for (const b of blocks) {
  if (b.probe.deck !== 'last') continue;
  const eff = b.probe.sz / 100;
  lastSamples.push({
    face: b.probe.face,
    eff,
    advance: b.advance,
    lastLine: b.blockHeight - (b.lines - 1) * b.advance,
    id: b.probe.id,
    fittable: b.probe.lnSpc?.kind === 'percent',
  });
}

/** Least squares for `lastLine = a x advance + b x eff`, two unknowns. */
function fitTwo(samples: readonly LastLineSample[]): { a: number; b: number; residual: number } {
  let sxx = 0;
  let sxy = 0;
  let syy = 0;
  let sxz = 0;
  let syz = 0;
  for (const s of samples) {
    sxx += s.advance * s.advance;
    sxy += s.advance * s.eff;
    syy += s.eff * s.eff;
    sxz += s.advance * s.lastLine;
    syz += s.eff * s.lastLine;
  }
  const det = sxx * syy - sxy * sxy;
  if (Math.abs(det) < 1e-12) throw new Error('degenerate fit');
  const a = (sxz * syy - syz * sxy) / det;
  const b = (syz * sxx - sxz * sxy) / det;
  let residual = 0;
  for (const s of samples) {
    residual = Math.max(residual, Math.abs(a * s.advance + b * s.eff - s.lastLine));
  }
  return { a, b, residual };
}

const faces = [...new Set(lastSamples.map((s) => s.face))];
const perFace: Record<string, number> = {};
/** The per-face floor threshold C, filled in from the  sweep below. */
const floorRatio: Record<string, number> = {};
say('    fitted lastLine = a x advance + b x effectiveSize, per face:');
let worstResidual = 0;
const coefficients: { face: string; a: number; b: number; residual: number }[] = [];
for (const face of faces) {
  const subset = lastSamples.filter((s) => s.face === face && s.fittable);
  const fit = fitTwo(subset);
  coefficients.push({ face, ...fit });
  worstResidual = Math.max(worstResidual, fit.residual);
  say(
    `      ${face.padEnd(16)} a = ${fit.a.toFixed(6)}   b = ${fit.b.toFixed(6)}` +
      `   worst residual ${fit.residual.toFixed(4)}pt  (${String(subset.length)} samples)`,
  );
}

// `a` is expected to be one number for every face; `b` is expected not to be.
const aValues = coefficients.map((c) => c.a);
const aSpread = Math.max(...aValues) - Math.min(...aValues);
const aMean = aValues.reduce((x, y) => x + y, 0) / aValues.length;
say();
say(
  `    a spans ${aSpread.toFixed(6)} across ${String(faces.length)} faces, mean ${aMean.toFixed(6)}`,
);
if (aSpread > 0.002) {
  throw new Error(
    `the advance coefficient is not one number across faces (spread ${String(aSpread)})`,
  );
}
const ADVANCE_COEFFICIENT = Math.round(aMean * 100) / 100;
say(`    taken as ${String(ADVANCE_COEFFICIENT)}`);

// Re-fit b with a pinned, so the per-face number is not absorbing error in a.
say();
say(`    with a pinned at ${String(ADVANCE_COEFFICIENT)}:`);
const perFaceResidual: Record<string, number> = {};
for (const face of faces) {
  const subset = lastSamples.filter((s) => s.face === face && s.fittable);
  let num = 0;
  let den = 0;
  for (const s of subset) {
    num += s.eff * (s.lastLine - ADVANCE_COEFFICIENT * s.advance);
    den += s.eff * s.eff;
  }
  // Nine decimals. The coefficient multiplies one line's size, so at the
  // largest size measured a rounding at the ninth place moves a block height by
  // 4e-8pt - three orders below the tolerance the fit test compares with, and
  // still short enough to read.
  const b = Math.round((num / den) * 1e9) / 1e9;
  perFace[face] = b;
  let residual = 0;
  for (const s of subset) {
    residual = Math.max(
      residual,
      Math.abs(ADVANCE_COEFFICIENT * s.advance + b * s.eff - s.lastLine),
    );
  }
  perFaceResidual[face] = residual;
  say(`      ${face.padEnd(16)} b = ${b.toFixed(6)}   worst residual ${residual.toFixed(4)}pt`);
}
const bSpread = Math.max(...Object.values(perFace)) - Math.min(...Object.values(perFace));
say();
say(`    b spans ${bSpread.toFixed(6)} across the six faces - it is a property of the typeface`);
say();

/*
 * `C`, from the sweep that holds the line spacing at an exact length and moves
 * the font size through it. Every size at which the last line still measures
 * the whole advance is a lower bound; the first size at which it drops to the
 * fitted value is an upper one.
 */
say('    the floor threshold C, bracketed by the sweep:');
const floorBracket: Record<string, { lower: number; upper: number }> = {};
for (const face of faces) {
  let lower = 0;
  let upper = Infinity;
  for (const b of blocks) {
    if (b.probe.deck !== 'floor' || b.probe.face !== face) continue;
    const eff = b.probe.sz / 100;
    const ratio = b.advance / eff;
    const lastLine = b.blockHeight - (b.lines - 1) * b.advance;
    const coefficient = perFace[face];
    if (coefficient === undefined) throw new Error('unreachable');
    const fitted = ADVANCE_COEFFICIENT * b.advance + coefficient * eff;
    if (Math.abs(lastLine - b.advance) < 0.006) lower = Math.max(lower, ratio);
    else if (fitted < b.advance - 0.006) upper = Math.min(upper, ratio);
  }
  floorBracket[face] = { lower, upper };
  // The lower bound is the value, rounded *up* to six decimals so that the
  // probe which established it still satisfies `advance <= C x size` - rounding
  // down by a millionth would put that one probe on the wrong branch. Five of
  // the six brackets start at exactly 1.2 and are unchanged by this.
  floorRatio[face] = Math.ceil(lower * 1e6 - 1e-6) / 1e6;
  say(
    `      ${face.padEnd(16)} C in [${lower.toFixed(6)}, ${upper.toFixed(6)})` +
      `${Math.abs(lower - NATURAL) < 0.02 ? '   indistinguishable from the 1.2 line box' : '   NOT the line box'}`,
  );
}
say();

/**
 * The floor, and where it stops.
 *
 * At single spacing the fitted line is *smaller* than the advance and the last
 * line measures the advance instead - which is why 3.2's whole table fitted
 * `n x advance` and why every ladder deck here does too. So there is a floor.
 * But it does not hold for ever: a 24pt Arial run with a 30pt exact line
 * spacing draws its last line at 27.96pt, which is below both the 30pt advance
 * and the 28.8pt natural box.
 *
 * So the floor holds while `advance <= C x size` for a per-face `C`, and the
 * `floor` deck reads `C` off directly. The natural guess is that `C` is 1.2 -
 * the line box - and for five of the six faces it is indistinguishable from
 * it. Courier New's is 1.27, which is why the guess had to be measured.
 */
function lastLineOf(advance: number, eff: number, face: string): number {
  const b = perFace[face];
  const c = floorRatio[face];
  if (b === undefined || c === undefined) throw new Error(`no line metrics for ${face}`);
  const fitted = ADVANCE_COEFFICIENT * advance + b * eff;
  return advance <= c * eff + 1e-9 ? Math.max(advance, fitted) : fitted;
}

interface BlockModel {
  readonly name: string;
  readonly height: (advance: number, eff: number, face: string, lines: number) => number;
}

const BLOCK_MODELS: readonly BlockModel[] = [
  {
    name: 'measured',
    height: (adv, eff, face, n) => (n - 1) * adv + lastLineOf(adv, eff, face),
  },
  { name: 'n x advance', height: (adv, _e, _f, n) => n * adv },
  { name: 'natural box', height: (adv, eff, _f, n) => (n - 1) * adv + NATURAL * eff },
  {
    name: 'min floor',
    height: (adv, eff, face, n) => {
      const b = perFace[face];
      if (b === undefined) throw new Error('unreachable');
      return (
        (n - 1) * adv + Math.max(Math.min(adv, NATURAL * eff), ADVANCE_COEFFICIENT * adv + b * eff)
      );
    },
  },
  {
    // The reading five of the six faces cannot tell apart from the measured
    // one, and Courier New refutes.
    name: 'floor 1.2',
    height: (adv, eff, face, n) => {
      const b = perFace[face];
      if (b === undefined) throw new Error('unreachable');
      const fitted = ADVANCE_COEFFICIENT * adv + b * eff;
      return (n - 1) * adv + (adv <= NATURAL * eff + 1e-9 ? Math.max(adv, fitted) : fitted);
    },
  },
  {
    name: 'no floor',
    height: (adv, eff, face, n) => {
      const b = perFace[face];
      if (b === undefined) throw new Error('unreachable');
      return (n - 1) * adv + ADVANCE_COEFFICIENT * adv + b * eff;
    },
  },
];

/**
 * The view-deck readings a block model can be scored against: every probe with
 * measured face metrics and no paragraph spacing. Spacing is a separate term
 * with its own section, and a block model scored against a spacing probe would
 * be scoring two rules at once and could not say which one was wrong.
 */
const blockSubset = blocks.filter(
  (b) =>
    perFace[b.probe.face] !== undefined &&
    b.probe.spcBef === undefined &&
    b.probe.spcAft === undefined,
);

const blockScores: Score[] = [];
for (const model of BLOCK_MODELS) {
  let fits = 0;
  const misses: string[] = [];
  // Paragraph spacing is a separate term with its own section; a block model
  // scored against `spacing-ctl` would be scoring two rules at once and could
  // not say which one was wrong.
  const subset = blockSubset;
  for (const b of subset) {
    const nominal = b.probe.sz / 100;
    const eff = effSize(nominal, (b.probe.stored.fontScale ?? 100000) / 100000);
    const predicted = model.height(b.advance, eff, b.probe.face, b.lines);
    if (Math.abs(predicted - b.blockHeight) < 0.02) fits += 1;
    else if (misses.length < 8) {
      misses.push(
        `${b.probe.id} predicted ${predicted.toFixed(3)} measured ${b.blockHeight.toFixed(3)}`,
      );
    }
  }
  blockScores.push({ name: model.name, fits, total: subset.length, misses });
}
table(blockScores);
chooseUnique('the block height', blockScores);
say();

/* ---------------------------------- C2: can a browser supply the constants */

say('C2. Whether a browser can supply b and C');

interface FaceMetricReading {
  readonly face: string;
  readonly sizePt: number;
  readonly metrics: {
    readonly fontBoundingBoxAscent: number;
    readonly fontBoundingBoxDescent: number;
  };
}
interface BrowserMetrics {
  readonly chromium: string;
  readonly results: readonly FaceMetricReading[];
}
/** Absent when `measure-face-metrics.ts` has not been run. The section says so
 *  rather than failing: the browser comparison is a side question, and the
 *  fixture it feeds is a negative result rather than a rule. */
function readBrowserMetrics(): BrowserMetrics | null {
  try {
    return JSON.parse(readFileSync(join(work, 'face-metrics.json'), 'utf8')) as BrowserMetrics;
  } catch {
    return null;
  }
}
const browserMetrics = readBrowserMetrics();

const browserComparison: {
  face: string;
  descentRatio: number;
  boxRatio: number;
  b: number;
  c: number;
}[] = [];
if (browserMetrics === null) {
  say('    face-metrics.json not present - run measure-face-metrics.ts');
} else {
  // The largest size measured, where Chromium's whole-pixel quantisation of
  // these two numbers is four decimal places below the ratio being read.
  const at = Math.max(...browserMetrics.results.map((r) => r.sizePt));
  say(`    chromium ${browserMetrics.chromium}, read at ${String(at)}px:`);
  for (const face of faces) {
    const row = browserMetrics.results.find((r) => r.face === face && r.sizePt === at);
    const b = perFace[face];
    const c = floorRatio[face];
    if (row === undefined || b === undefined || c === undefined) continue;
    const descentRatio = row.metrics.fontBoundingBoxDescent / at;
    const boxRatio = (row.metrics.fontBoundingBoxAscent + row.metrics.fontBoundingBoxDescent) / at;
    browserComparison.push({ face, descentRatio, boxRatio, b, c });
    say(
      `      ${face.padEnd(16)} descent ${descentRatio.toFixed(6)} vs b ${b.toFixed(6)}` +
        ` (${(b - descentRatio >= 0 ? '+' : '') + (b - descentRatio).toFixed(6)})` +
        `   fontBox ${boxRatio.toFixed(6)} vs C ${c.toFixed(6)}`,
    );
  }
  const worstB = Math.max(...browserComparison.map((r) => Math.abs(r.b - r.descentRatio)));
  // Two faces reporting the same descent while their measured b differs is the
  // decisive part: no function of the descent alone can produce both.
  const collisions = browserComparison.filter((r) =>
    browserComparison.some(
      (o) =>
        o.face !== r.face &&
        Math.abs(o.descentRatio - r.descentRatio) < 0.003 &&
        Math.abs(o.b - r.b) > 0.01,
    ),
  );
  say();
  say(`    worst |b - descent| is ${worstB.toFixed(6)}, and the descent is not a function of b:`);
  say(
    `    ${String(collisions.length)} faces share a descent to within 0.003 while their b differs by more than 0.01` +
      ` (${collisions.map((r) => r.face).join(', ')})`,
  );
  if (collisions.length === 0) {
    throw new Error('the browser descent may explain b after all - re-open the question');
  }
}
say();

/* ------------------------------------------------------------ D: the ladder */

say('D. The ladder');

const observedRungs = new Map<string, Rung>();
for (const c of cases) observedRungs.set(rungKey(c.observed), c.observed);

/**
 * The order the rungs are tried in: font scale descending and, within one font
 * scale, the smaller reduction first. That this reproduces every observation is
 * the measurement; the alternative - take whichever fitting rung leaves the
 * text tallest - is scored beside it in section E and is refuted by the sizes
 * where rounding makes a smaller font scale the taller rung.
 */
const LADDER: readonly Rung[] = [...observedRungs.values()].sort(
  (a, b) => b.fontScale - a.fontScale || a.lnSpcReduction - b.lnSpcReduction,
);

say(`    ${String(LADDER.length)} distinct rungs across every recompute deck:`);
for (const [i, r] of LADDER.entries()) {
  say(
    `      ${String(i + 1).padStart(2)}. fontScale ${String(r.fontScale).padStart(6)}` +
      `  lnSpcReduction ${String(r.lnSpcReduction).padStart(5)}`,
  );
}
say();

/* ------------------------------------------------------- E: the fit test */

say('E. The fit test');

type ParaSpaceRule = (v: Spacing | undefined, effPt: number, nominalPt: number) => number;

const PARA_RULES: Readonly<Record<string, ParaSpaceRule>> = {
  /** Percent of the scaled line box; exact points used raw. */
  lineBoxRawPoints: (v, eff) =>
    v === undefined ? 0 : v.kind === 'points' ? v.value / 100 : (v.value / 100000) * NATURAL * eff,
  lineBoxScaledPoints: (v, eff, nominal) =>
    v === undefined
      ? 0
      : v.kind === 'points'
        ? (v.value / 100) * (eff / nominal)
        : (v.value / 100000) * NATURAL * eff,
  sizeRawPoints: (v, eff) =>
    v === undefined ? 0 : v.kind === 'points' ? v.value / 100 : (v.value / 100000) * eff,
  nominalBoxRawPoints: (v, _eff, nominal) =>
    v === undefined
      ? 0
      : v.kind === 'points'
        ? v.value / 100
        : (v.value / 100000) * NATURAL * nominal,
};

interface Model {
  readonly name: string;
  readonly size: string;
  readonly advance: string;
  readonly para: string;
  /** A key of BLOCK_MODELS, so a model added in the middle cannot silently
   *  renumber the rivals - which it did once, and turned a refuted reading
   *  into a tie. */
  readonly block: string;
  readonly spcFirstLastParaDefault: boolean;
  readonly useInsets: boolean;
  readonly tallestFitting: boolean;
  readonly ladder: readonly Rung[];
}

/** Points below which two lengths are the same length - four orders of
 *  magnitude below the quarter point the box heights are quantised to. */
const EPS = 1e-7;

function textHeight(m: Model, p: Probe, lines: number, rung: Rung): number {
  const nominal = p.sz / 100;
  const sizeRule = SIZE_RULES[m.size];
  const advanceRule = ADVANCE_RULES[m.advance];
  const paraRule = PARA_RULES[m.para];
  const blockModel = BLOCK_MODELS.find((x) => x.name === m.block);
  if (
    sizeRule === undefined ||
    advanceRule === undefined ||
    paraRule === undefined ||
    blockModel === undefined
  ) {
    throw new Error(`model ${m.name} names a rule that does not exist`);
  }
  const eff = sizeRule(nominal, rung.fontScale / 100000);
  const advance = advanceRule(eff, rung.lnSpcReduction, p.lnSpc);
  const paragraphs = p.paragraphs.length;
  const counted = p.spcFirstLastPara ?? m.spcFirstLastParaDefault;
  const outer = counted ? paragraphs : paragraphs - 1;
  return (
    blockModel.height(advance, eff, p.face, lines) +
    paraRule(p.spcBef, eff, nominal) * outer +
    paraRule(p.spcAft, eff, nominal) * outer
  );
}

function availableHeight(m: Model, rect: { readonly h: number }, insets: Insets): number {
  return m.useInsets ? rect.h - insets.t - insets.b : rect.h;
}

function choose(m: Model, p: Probe, lines: number): Rung {
  const available = availableHeight(m, p.rect, p.insets);
  const floor = m.ladder[m.ladder.length - 1];
  if (floor === undefined) throw new Error('empty ladder');
  if (m.tallestFitting) {
    let best: Rung | undefined;
    let bestHeight = -1;
    for (const rung of m.ladder) {
      const h = textHeight(m, p, lines, rung);
      if (h <= available + EPS && h > bestHeight) {
        best = rung;
        bestHeight = h;
      }
    }
    return best ?? floor;
  }
  for (const rung of m.ladder) {
    if (textHeight(m, p, lines, rung) <= available + EPS) return rung;
  }
  return floor;
}

const BASE_MODEL: Model = {
  name: 'measured',
  size: sizeWinner.name,
  advance: advanceWinner.name,
  para: 'lineBoxRawPoints',
  block: 'measured',
  spcFirstLastParaDefault: false,
  useInsets: true,
  tallestFitting: false,
  ladder: LADDER,
};

const scored = cases.filter((c) => c.fixedLines);
const wrapped = cases.filter((c) => !c.fixedLines);

const RIVALS: readonly Model[] = [
  BASE_MODEL,
  { ...BASE_MODEL, name: 'size: rounded even at 100%', size: 'roundWholeAlways' },
  { ...BASE_MODEL, name: 'size: no rounding', size: 'exact' },
  { ...BASE_MODEL, name: 'size: round half to even', size: 'roundHalfEven' },
  { ...BASE_MODEL, name: 'size: floor', size: 'floorWhole' },
  { ...BASE_MODEL, name: 'size: nearest half point', size: 'roundHalfPoint' },
  { ...BASE_MODEL, name: 'reduction multiplies the advance', advance: 'multiply' },
  {
    ...BASE_MODEL,
    name: 'reduction also reduces exact points',
    advance: 'subtractAndReducePoints',
  },
  { ...BASE_MODEL, name: 'block: n x advance', block: 'n x advance' },
  { ...BASE_MODEL, name: 'block: last line is the natural 1.2 box', block: 'natural box' },
  { ...BASE_MODEL, name: 'block: floored at min(advance, the natural box)', block: 'min floor' },
  { ...BASE_MODEL, name: 'block: the floor threshold is 1.2 everywhere', block: 'floor 1.2' },
  { ...BASE_MODEL, name: 'block: the fitted line with no floor', block: 'no floor' },
  { ...BASE_MODEL, name: 'spcPts scales with the font', para: 'lineBoxScaledPoints' },
  { ...BASE_MODEL, name: 'spcPct is a fraction of the size', para: 'sizeRawPoints' },
  { ...BASE_MODEL, name: 'spcPct is a fraction of the unscaled box', para: 'nominalBoxRawPoints' },
  { ...BASE_MODEL, name: 'spcFirstLastPara defaults true', spcFirstLastParaDefault: true },
  { ...BASE_MODEL, name: 'the fit test ignores the insets', useInsets: false },
  { ...BASE_MODEL, name: 'take the tallest rung that fits', tallestFitting: true },
  {
    ...BASE_MODEL,
    name: 'the reduction is always 20%',
    ladder: LADDER.filter((r) => r.lnSpcReduction === 20000 || r.fontScale === 100000),
  },
];

function scoreModel(m: Model, subset: readonly Case[]): Score {
  let fits = 0;
  const misses: string[] = [];
  for (const c of subset) {
    const got = choose(m, c.probe, c.lines);
    if (
      got.fontScale === c.observed.fontScale &&
      got.lnSpcReduction === c.observed.lnSpcReduction
    ) {
      fits += 1;
    } else if (misses.length < 10) {
      misses.push(
        `${c.probe.id} box=${String(c.probe.rect.h)}pt lines=${String(c.lines)} ` +
          `predicted ${rungKey(got)} observed ${rungKey(c.observed)}`,
      );
    }
  }
  return { name: m.name, fits, total: subset.length, misses };
}

/**
 * Both instruments, one score.
 *
 * The emitted rungs cannot tell the three block-height readings apart: they
 * differ only in what the *last* line measures, and a rung is a step function
 * of the whole height, so a difference of a point in one line rarely crosses a
 * threshold. The block heights PowerPoint reports for text it is not
 * autofitting separate them at once. Neither instrument alone settles the rule,
 * so a candidate has to survive both, and the score is the sum.
 */
function scoreCombined(m: Model): Score {
  const emitted = scoreModel(m, scored);
  const blockModel = BLOCK_MODELS.find((x) => x.name === m.block);
  if (blockModel === undefined) throw new Error(`model ${m.name} names no block rule`);
  let fits = 0;
  const misses: string[] = [];
  for (const b of blockSubset) {
    const eff = effSize(b.probe.sz / 100, (b.probe.stored.fontScale ?? 100000) / 100000);
    const predicted = blockModel.height(b.advance, eff, b.probe.face, b.lines);
    if (Math.abs(predicted - b.blockHeight) < 0.02) fits += 1;
    else if (misses.length < 6) {
      misses.push(
        `${b.probe.id} block predicted ${predicted.toFixed(3)} measured ${b.blockHeight.toFixed(3)}`,
      );
    }
  }
  return {
    name: m.name,
    fits: emitted.fits + fits,
    total: emitted.total + blockSubset.length,
    misses: [...emitted.misses, ...misses],
  };
}

const scores = RIVALS.map(scoreCombined);
table(scores);
const winner = chooseUnique('the autofit rule', scores);
say();
say(`    ${winner.name}: ${String(winner.fits)} / ${String(winner.total)}`);
say();

/* ------------------------------------------ F: is a stored scale used at all */

say('F. Whether a stored scale is used at all');
const verbatim: string[] = [];
for (const b of blocks) {
  if (!b.probe.id.startsWith('stored-x')) continue;
  const nominal = b.probe.sz / 100;
  const eff = effSize(nominal, (b.probe.stored.fontScale ?? 100000) / 100000);
  const stated = advanceOf(eff, b.probe.stored.lnSpcReduction ?? 0, b.probe.lnSpc);
  const recomputed = NATURAL * nominal;
  verbatim.push(
    `    ${b.probe.id}: measured ${b.advance.toFixed(3)}pt, stated ${stated.toFixed(3)}pt, ` +
      `a recomputation would give ${recomputed.toFixed(3)}pt`,
  );
}
for (const v of verbatim) say(v);
findings['storedApplied'] = verbatim;
say();

/* ------------------------------------------------------------- G: spAutoFit */

say('G. spAutoFit');

const spKept = spCases.filter((s) => Math.abs(s.emu / EMU_PER_POINT - s.boxPt) < 1e-6);
const spMoved = spCases.filter((s) => !spKept.includes(s));

interface SpRatio {
  readonly id: string;
  readonly face: string;
  readonly ratio: number;
}
const spRatios: SpRatio[] = [];
for (const s of spMoved) {
  const eff = s.probe.sz / 100;
  const advance = advanceOf(eff, 0, s.probe.lnSpc);
  const needed =
    (s.lines - 1) * advance +
    lastLineOf(advance, eff, s.probe.face) +
    s.probe.insets.t +
    s.probe.insets.b;
  spRatios.push({ id: s.probe.id, face: s.probe.face, ratio: s.emu / EMU_PER_POINT / needed });
}
const ratios = spRatios.map((r) => r.ratio);
const ratioMin = Math.min(...ratios);
const ratioMax = Math.max(...ratios);
say(`    ${String(spCases.length)} probes; ${String(spKept.length)} kept their authored height`);
say(`    of the ${String(spMoved.length)} that moved, height / (text + tIns + bIns) spans`);
say(`      ${ratioMin.toFixed(9)} .. ${ratioMax.toFixed(9)}`);
for (const face of [...new Set(spRatios.map((r) => r.face))]) {
  const subset = spRatios.filter((r) => r.face === face).map((r) => r.ratio);
  say(
    `      ${face.padEnd(16)} ${Math.min(...subset).toFixed(9)} .. ${Math.max(...subset).toFixed(9)}` +
      `  (${String(subset.length)} probes)`,
  );
}
const SP_FACTOR = (ratioMin + ratioMax) / 2;
say(
  `    taken as ${SP_FACTOR.toFixed(9)}; the spread is ${(ratioMax - ratioMin).toExponential(2)}`,
);

// The hysteresis: which boxes were left alone, and how far from the target.
const hyst = spCases
  .filter((s) => s.probe.deck === 'sp-hyst')
  .map((s) => {
    const eff = s.probe.sz / 100;
    const advance = advanceOf(eff, 0, s.probe.lnSpc);
    const needed =
      (s.lines - 1) * advance +
      lastLineOf(advance, eff, s.probe.face) +
      s.probe.insets.t +
      s.probe.insets.b;
    return { box: s.boxPt, target: SP_FACTOR * needed, got: s.emu / EMU_PER_POINT };
  })
  .sort((a, b) => a.box - b.box);
const untouched = hyst.filter((h) => Math.abs(h.got - h.box) < 1e-6);
say();
if (untouched.length > 0) {
  const boxes = untouched.map((h) => h.box);
  const target = hyst[0]?.target ?? 0;
  say(
    `    the hysteresis sweep left ${String(untouched.length)} boxes alone, from ` +
      `${String(Math.min(...boxes))}pt to ${String(Math.max(...boxes))}pt, against a target of ` +
      `${target.toFixed(3)}pt`,
  );
} else {
  say('    every box in the hysteresis sweep was resized');
}
say();

/* -------------------------------------------------------------- H: wrapping */

say('H. Wrapping');
let wrapFits = 0;
let wrapProven = 0;
for (const c of wrapped) {
  const available = availableHeight(BASE_MODEL, c.probe.rect, c.probe.insets);
  if (textHeight(BASE_MODEL, c.probe, c.lines, c.observed) <= available + EPS) wrapFits += 1;
  const index = LADDER.findIndex(
    (r) => r.fontScale === c.observed.fontScale && r.lnSpcReduction === c.observed.lnSpcReduction,
  );
  if (index <= 0) {
    wrapProven += 1;
    continue;
  }
  const previous = LADDER[index - 1];
  if (previous === undefined) throw new Error('unreachable');
  // A larger font can only produce more lines, never fewer, so a rung that
  // overflows at the *observed* count certainly overflows at its own.
  if (textHeight(BASE_MODEL, c.probe, c.lines, previous) > available + EPS) wrapProven += 1;
}
say(
  `    ${String(wrapFits)} / ${String(wrapped.length)} chose a rung that fits the line count observed`,
);
say(
  `    ${String(wrapProven)} / ${String(wrapped.length)} additionally prove the rung above overflows ` +
    'without needing a line breaker',
);
if (wrapFits !== wrapped.length) throw new Error('a wrapped probe chose a rung that does not fit');
say();

/* -------------------------------------------------------------- the fixture */

interface SweepRow {
  readonly deck: string;
  readonly sz: number;
  readonly face: string;
  readonly lines: number;
  readonly lnSpc: Spacing | null;
  readonly spcBef: Spacing | null;
  readonly spcAft: Spacing | null;
  readonly spcFirstLastPara: boolean | null;
  readonly insets: Insets;
  readonly boxPt: readonly number[];
  readonly rung: readonly number[];
}

const ladderIndex = new Map(LADDER.map((r, i) => [rungKey(r), i]));
const groups = new Map<string, Case[]>();
for (const c of scored) {
  const p = c.probe;
  const key = JSON.stringify([
    c.deck,
    p.sz,
    p.face,
    p.paragraphs.length,
    p.lnSpc ?? null,
    p.spcBef ?? null,
    p.spcAft ?? null,
    p.spcFirstLastPara ?? null,
    p.insets,
  ]);
  const g = groups.get(key);
  if (g === undefined) groups.set(key, [c]);
  else g.push(c);
}

const sweeps: SweepRow[] = [];
for (const g of groups.values()) {
  const sorted = [...g].sort((a, b) => a.probe.rect.h - b.probe.rect.h);
  const first = sorted[0];
  if (first === undefined) throw new Error('unreachable');
  const p = first.probe;
  sweeps.push({
    deck: first.deck,
    sz: p.sz,
    face: p.face,
    lines: first.lines,
    lnSpc: p.lnSpc ?? null,
    spcBef: p.spcBef ?? null,
    spcAft: p.spcAft ?? null,
    spcFirstLastPara: p.spcFirstLastPara ?? null,
    insets: p.insets,
    boxPt: sorted.map((c) => c.probe.rect.h),
    rung: sorted.map((c) => {
      const i = ladderIndex.get(rungKey(c.observed));
      if (i === undefined) throw new Error(`rung ${rungKey(c.observed)} is not in the ladder`);
      return i;
    }),
  });
}

/**
 * Drop the keys whose value is null.
 *
 * Most block rows state no line spacing, no paragraph spacing and no font
 * scale, and 580 rows times six nulls is a third of this fixture. Nothing is
 * lost: every one of those fields means "the file stated nothing", so an absent
 * key and a null one say the same thing.
 */
function dropNulls<T extends Record<string, unknown>>(row: T): Partial<T> {
  const out: Partial<T> = {};
  for (const [key, value] of Object.entries(row)) {
    if (value !== null) out[key as keyof T] = value as T[keyof T];
  }
  return out;
}

const round = (x: number, places = 3): number => {
  const k = 10 ** places;
  return Math.round(x * k) / k;
};

const fixture = {
  $comment:
    'Experiment T4, sub-phase 3.4. What PowerPoint does with a:normAutofit and a:spAutoFit, ' +
    'measured by making PowerPoint compute an autofit and reading the numbers it wrote into its own ' +
    'file. Every value here is PowerPoint 16.0.20326 output, not a reading of ECMA-376. ' +
    'See docs/adr/phase-3-text/0030-autofit.md.',
  measuredOn: {
    application: 'Microsoft PowerPoint',
    version: '16.0',
    build: '20326',
    platform: 'Windows 11 Enterprise 10.0.26200',
  },
  method: {
    trigger:
      'PowerPoint recomputes autofit only while the presentation has a window. Opened with ' +
      'WithWindow:=msoFalse, every probe came back unshrunk however hard the text frame was poked - ' +
      'setting TextFrame2.AutoSize and re-editing the text both did nothing. The window may be ' +
      'minimised: the emitted scales were checked against a run with it restored and are identical.',
    recompute:
      'TextFrame2.AutoSize is driven to msoAutoSizeNone and then to msoAutoSizeTextToFitShape (or ' +
      'msoAutoSizeShapeToFitText for the spAutoFit probes), the deck is saved, and the emitted ' +
      'a:normAutofit is read out of the saved XML.',
    view: 'The view decks are opened read-only, read, and closed without being saved.',
  },
  ladder: LADDER,
  ladderOrder:
    'Tried in order: font scale descending, and within one font scale the smaller reduction first. ' +
    'Not the tallest rung that fits - at 32pt and at 11pt, rounding makes (85000, 10000) taller than ' +
    '(92500, 20000), and PowerPoint still takes the second.',
  lineBox: {
    naturalFactor: NATURAL,
    advanceCoefficient: ADVANCE_COEFFICIENT,
    perFace,
    perFaceResidualPt: perFaceResidual,
    floorRatio,
    floorBracket,
    browserComparison,
    browserVerdict:
      'Chromium TextMetrics cannot supply either constant. fontBoundingBoxDescent is within 0.0013 ' +
      'of b on Courier New, Tahoma and Verdana and 0.016 out on Arial - and Arial and Verdana ' +
      'report descents 0.002 apart while their measured b differs by 0.020, so no function of the ' +
      'descent alone produces both. fontBoundingBoxAscent + Descent puts Verdana above Courier New ' +
      'where C puts it below.',
    rule:
      'lastLine = max(min(advance, 1.2 x effectiveSize), a x advance + b x effectiveSize), with a ' +
      'the same for every face and b a property of the typeface. At single spacing the first branch ' +
      'wins and the last line is exactly the advance, which is why a block of n lines measures ' +
      'n x advance there and why sub-phase 3.2 could not separate the two.',
  },
  rules: {
    effectiveSize:
      'round(nominal points x fontScale) to a whole point, halves up - but only when a scale ' +
      'applies. A fontScale of exactly 100% leaves a fractional size alone: 10.5pt stays 10.5pt.',
    advance:
      'a:lnSpc percent minus lnSpcReduction, times 1.2, times the effective size. a:lnSpc in exact ' +
      'points is used as given and ignores lnSpcReduction entirely.',
    paragraphSpacing:
      'a:spcPct is a fraction of the scaled 1.2 line box; a:spcPts is used raw and does not follow ' +
      'the font scale. spcFirstLastPara defaults false, which drops the first paragraph spcBef and ' +
      'the last spcAft.',
    fit:
      '(lines - 1) x advance + lastLine, plus the paragraph spacing terms, against the box height ' +
      'minus tIns and bIns. The first rung of the ladder that fits wins.',
    storedValues:
      'A stored fontScale or lnSpcReduction is applied exactly as written, whether or not it is a ' +
      'rung, and opening a file never recomputes it.',
    shapeAutofit:
      'a:spAutoFit sets the shape height to the text height plus both insets, times a constant just ' +
      'under one per cent - the same constant on every face measured, so it is not a font metric.',
  },
  spAutoFit: {
    factor: SP_FACTOR,
    factorMin: ratioMin,
    factorMax: ratioMax,
    factorSpread: ratioMax - ratioMin,
    factorNote:
      'The bracket contains 517/512 = 1.009765625 exactly, which is what the package uses: a ' +
      'dyadic number inside a measured range is a better guess at what a program computed than ' +
      'the midpoint of the range is.',
    probes: spCases.map((s) => ({
      id: s.probe.id,
      face: s.probe.face,
      sz: s.probe.sz,
      lines: s.lines,
      tIns: s.probe.insets.t,
      bIns: s.probe.insets.b,
      boxPt: s.boxPt,
      emu: s.emu,
    })),
  },
  scores: scores.map((s) => ({ model: s.name, fits: s.fits, total: s.total })),
  sizeScores: sizeScores.map((s) => ({ model: s.name, fits: s.fits, total: s.total })),
  advanceScores: advanceScores.map((s) => ({ model: s.name, fits: s.fits, total: s.total })),
  blockScores: blockScores.map((s) => ({ model: s.name, fits: s.fits, total: s.total })),
  findings,
  sweeps,
  blocks: blocks.map((b) =>
    dropNulls({
      id: b.probe.id,
      deck: b.probe.deck,
      face: b.probe.face,
      sz: b.probe.sz,
      autofit: b.probe.autofit,
      fontScale: b.probe.stored.fontScale ?? null,
      lnSpcReduction: b.probe.stored.lnSpcReduction ?? null,
      lnSpc: b.probe.lnSpc ?? null,
      spcBef: b.probe.spcBef ?? null,
      spcAft: b.probe.spcAft ?? null,
      spcFirstLastPara: b.probe.spcFirstLastPara ?? null,
      paragraphs: b.probe.paragraphs.length,
      lines: b.lines,
      // The mean distance between line tops, which is the *line* advance only
      // when nothing else sits between two lines. A probe stating spcBef or
      // spcAft has that spacing inside every gap, so the number would be a sum of
      // two rules wearing the name of one - and a reader who used it as the
      // advance would be wrong by exactly the paragraph spacing. Omitted there;
      // `blockHeightPt` is what those probes measure.
      advancePt:
        b.probe.spcBef === undefined && b.probe.spcAft === undefined ? round(b.advance) : null,
      blockHeightPt: round(b.blockHeight),
    }),
  ),
  wrap: wrapped.map((c) => ({
    id: c.probe.id,
    widthPt: c.probe.rect.w,
    heightPt: c.probe.rect.h,
    sz: c.probe.sz,
    face: c.probe.face,
    text: c.probe.paragraphs.join('\n'),
    lines: c.lines,
    rung: ladderIndex.get(rungKey(c.observed)) ?? -1,
  })),
  decks: inputs.decks,
};

const repoFixture = resolve(process.cwd(), 'corpus', 'ground-truth', 'autofit.json');
writeFileSync(repoFixture, `${JSON.stringify(fixture, null, 2)}\n`, 'utf8');
say(`wrote ${repoFixture}`);
say(
  `  ${String(scored.length)} scored rungs, ${String(blocks.length)} block readings, ` +
    `${String(spCases.length)} spAutoFit probes, ${String(inputs.decks.length)} decks`,
);

writeFileSync(join(work, 'autofit-report.txt'), report.join('\n'), 'utf8');
console.log(`wrote ${join(work, 'autofit-report.txt')}`);

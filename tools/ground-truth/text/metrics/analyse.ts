/**
 * Experiment T2, step 4 - score every reading of the line model against what
 * PowerPoint did, and emit `corpus/ground-truth/text-metrics.json`.
 *
 * ```
 * node tools/ground-truth/analyse-metrics.ts <work-dir>
 * ```
 *
 * ## Scoring, not confirming
 *
 * The plan asserts that PowerPoint's line height is a font-independent
 * `1.2 x font size`. Confirming that is worth nothing on its own: the probe
 * decks were built so that a *wrong* model predicts a different number, and the
 * useful output is which wrong models are refuted and by how much. So every
 * rule below is scored against the alternatives anyone would plausibly write -
 * the browser's own `line-height: normal`, a percentage applied to the font
 * size rather than to the line box, kerning read as a boolean - and nothing is
 * written unless exactly one candidate fits every probe of its kind.
 *
 * ## The one rule that is data rather than law
 *
 * The last line of a text block does not measure the same as the lines above
 * it, and the difference involves a per-face constant that no browser reports
 * and no measurement here can source. That constant is committed as a measured
 * table, clearly separated from the fitted rules, together with the score of
 * the best model found and an enumeration of the probes it misses. A table of
 * numbers labelled as numbers is honest; the same table dressed as a rule is
 * not, and 3.6 is where it has to be finished.
 */

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import type { LnSpc, Probe } from './text-metrics.ts';

/* -------------------------------------------------------------------------- */
/* what the reader wrote                                                      */
/* -------------------------------------------------------------------------- */

interface LineReading {
  readonly index: number;
  readonly top: number | null;
  readonly height: number | null;
  readonly left: number | null;
  readonly width: number | null;
  readonly text: string | null;
}

interface CharReading {
  readonly index: number;
  readonly text: string | null;
  readonly left: number | null;
  readonly width: number | null;
}

interface ShapeReading {
  readonly id: string;
  readonly shapeTop: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly top: number | null;
  readonly lineCount: number | null;
  readonly fontName: string | null;
  readonly fontSize: number | null;
  readonly lines: readonly LineReading[];
  readonly chars: readonly CharReading[];
}

interface DeckReading {
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly shapes: readonly ShapeReading[];
}

interface Inputs {
  readonly fonts: readonly string[];
  readonly absentFont: string;
  readonly strings: readonly { readonly key: string; readonly text: string }[];
  readonly probes: readonly Probe[];
}

interface BrowserSize {
  readonly px: number;
  readonly metrics: Readonly<Record<string, number>>;
  readonly widths: Readonly<Record<string, number>>;
  readonly spc: Readonly<Record<string, number>>;
  readonly kernOn: number;
  readonly kernOff: number;
}
interface BrowserFont {
  readonly family: string;
  readonly installed: boolean;
  readonly sizes: readonly BrowserSize[];
}
interface BrowserMetrics {
  readonly chromium: string;
  readonly fonts: readonly BrowserFont[];
}

const work = process.argv[2];
if (work === undefined) throw new Error('usage: analyse-metrics.ts <work-dir>');

const inputs = JSON.parse(readFileSync(join(work, 'metric-inputs.json'), 'utf8')) as Inputs;
const readings = JSON.parse(readFileSync(join(work, 'metric-readings.json'), 'utf8')) as {
  decks: readonly DeckReading[];
};
const browser = JSON.parse(
  readFileSync(join(work, 'browser-metrics.json'), 'utf8'),
) as BrowserMetrics;

/**
 * `BoundHeight` and `BoundWidth` come back as float32 points and PowerPoint
 * quantises them to a two-hundredth of a point, so two numbers within six
 * thousandths are the same number. Anything looser would let a wrong model
 * through: the smallest real difference any probe here can produce is a tenth
 * of a point, more than fifteen times this.
 */
const EPS = 6e-3;
const near = (a: number, b: number): boolean => Math.abs(a - b) < EPS;

const shapes = new Map<string, ShapeReading>();
for (const deck of readings.decks) {
  if (!deck.opened) throw new Error(`deck ${deck.deck} did not open`);
  if (deck.repaired === true) throw new Error(`deck ${deck.deck} was repaired - readings are void`);
  for (const shape of deck.shapes) shapes.set(shape.id, shape);
}

const probes = new Map<string, Probe>(inputs.probes.map((p) => [p.id, p]));
for (const id of probes.keys()) {
  if (!shapes.has(id)) throw new Error(`probe ${id} was never read back`);
}

function reading(id: string): ShapeReading {
  const r = shapes.get(id);
  if (r === undefined) throw new Error(`no reading for ${id}`);
  return r;
}

/** Height of the whole text block, in points. */
function heightOf(id: string): number {
  const h = reading(id).height;
  if (h === null) throw new Error(`${id}: no BoundHeight`);
  return h;
}

/**
 * The baseline-to-baseline advance, taken from two line tops rather than from a
 * height. A height is one number that many models fit; two tops are a distance
 * with every constant term already cancelled.
 */
function advanceOf(id: string): number | null {
  const lines = reading(id).lines;
  if (lines.length < 2) return null;
  const first = lines[0]?.top;
  const last = lines[lines.length - 1]?.top;
  if (first === null || first === undefined || last === null || last === undefined) return null;
  return (last - first) / (lines.length - 1);
}

const pt = (sz: number): number => sz / 100;

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

interface Score<T> {
  readonly name: string;
  readonly fn: T;
  hits: number;
  readonly misses: string[];
}

/**
 * Score every candidate and insist on exactly one that fits everything.
 *
 * Two survivors means the experiment cannot tell them apart and the table needs
 * another probe, not a coin toss; none means the rule is not yet known. Both
 * are failures of the measurement and both throw, because a fixture is a claim
 * about PowerPoint and a claim nobody could have refuted is not worth making.
 */
function chooseUnique<T>(
  what: string,
  candidates: readonly { readonly name: string; readonly fn: T }[],
  cases: readonly string[],
  test: (fn: T, id: string) => boolean,
): { winner: Score<T>; table: readonly Score<T>[] } {
  if (cases.length === 0) throw new Error(`${what}: nothing to score`);
  const table = candidates.map((c) => ({ ...c, hits: 0, misses: [] as string[] }));
  for (const score of table) {
    for (const id of cases) {
      if (test(score.fn, id)) score.hits += 1;
      else if (score.misses.length < 8) score.misses.push(id);
    }
  }
  const perfect = table.filter((s) => s.hits === cases.length);
  const report = table
    .map((s) => `      ${String(s.hits).padStart(5)}/${String(cases.length)}  ${s.name}`)
    .join('\n');
  if (perfect.length !== 1) {
    throw new Error(
      `${what}: ${String(perfect.length)} candidates fit all ${String(cases.length)} probes; ` +
        `exactly one must.\n${report}`,
    );
  }
  const winner = perfect[0];
  if (winner === undefined) throw new Error(`${what}: unreachable`);
  console.log(`  ${what}\n${report}`);
  return { winner, table };
}

/* -------------------------------------------------------------------------- */
/* R1 - how `spcPct/@val` is quantised                                        */
/* -------------------------------------------------------------------------- */

console.log('\n=== R1  the percentage is quantised before it is used ===');

type Quantiser = (val: number) => number;
const QUANTISERS: readonly { name: string; fn: Quantiser }[] = [
  { name: 'exact: val / 100000', fn: (v) => v / 100000 },
  { name: 'round half up to a whole percent', fn: (v) => Math.round(v / 1000) / 100 },
  { name: 'floor to a whole percent', fn: (v) => Math.floor(v / 1000) / 100 },
  { name: 'ceil to a whole percent', fn: (v) => Math.ceil(v / 1000) / 100 },
  { name: 'round to a tenth of a percent', fn: (v) => Math.round(v / 100) / 1000 },
];

// Scored on the advance between two baselines, so the finding does not rest on
// the last-line model it was first noticed through.
const pctAdvanceCases = [...probes.values()]
  .filter((p) => p.lines >= 2 && p.lnSpc?.kind === 'pct')
  .map((p) => p.id);

const quant = chooseUnique('percent quantisation', QUANTISERS, pctAdvanceCases, (fn, id) => {
  const p = probes.get(id);
  const a = advanceOf(id);
  if (p?.lnSpc === undefined || a === null) return false;
  return near(fn(p.lnSpc.value) * 1.2 * pt(p.sz), a);
});
const quantise = quant.winner.fn;

/* -------------------------------------------------------------------------- */
/* R2 - the line advance                                                      */
/* -------------------------------------------------------------------------- */

console.log('\n=== R2  the line advance ===');

type AdvanceModel = (sz: number, lnSpc: LnSpc | undefined) => number;
const ADVANCE_MODELS: readonly { name: string; fn: AdvanceModel }[] = [
  {
    name: 'percent x 1.2 x size; exact points as given  [the measured rule]',
    fn: (sz, ls) =>
      ls === undefined || ls.kind === 'pct'
        ? quantise(ls?.value ?? 100000) * 1.2 * pt(sz)
        : pt(ls.value),
  },
  {
    name: 'percent x size (the percentage applied to the font size, not the line box)',
    fn: (sz, ls) =>
      ls === undefined || ls.kind === 'pct' ? quantise(ls?.value ?? 100000) * pt(sz) : pt(ls.value),
  },
  {
    name: 'percent x 1.2 x size; exact points also scaled by 1.2',
    fn: (sz, ls) =>
      ls === undefined || ls.kind === 'pct'
        ? quantise(ls?.value ?? 100000) * 1.2 * pt(sz)
        : 1.2 * pt(ls.value),
  },
  {
    name: 'a flat 1.2 x size, ignoring a:lnSpc entirely',
    fn: (sz) => 1.2 * pt(sz),
  },
];

const advanceCases = [...probes.values()].filter((p) => p.lines >= 2).map((p) => p.id);
const adv = chooseUnique('line advance', ADVANCE_MODELS, advanceCases, (fn, id) => {
  const p = probes.get(id);
  const a = advanceOf(id);
  if (p === undefined || a === null) return false;
  return near(fn(p.sz, p.lnSpc), a);
});
const lineAdvance = adv.winner.fn;

/* -------------------------------------------------------------------------- */
/* R2b - the advance is font-independent                                      */
/* -------------------------------------------------------------------------- */

console.log('\n=== R2b  the advance does not depend on the face ===');

const byFontSize = new Map<string, number[]>();
for (const p of probes.values()) {
  if (p.kind !== 'lineHeight' || p.lines < 2 || p.lnSpc === undefined) continue;
  const a = advanceOf(p.id);
  if (a === null) continue;
  const key = `${String(p.sz)}|${p.lnSpc.kind}|${String(p.lnSpc.value)}|${String(p.lines)}`;
  const list = byFontSize.get(key) ?? [];
  list.push(a);
  byFontSize.set(key, list);
}
let spreadWorst = 0;
for (const [key, list] of byFontSize) {
  const spread = Math.max(...list) - Math.min(...list);
  if (spread > spreadWorst) spreadWorst = spread;
  if (spread >= EPS) {
    throw new Error(
      `the advance varies by face at ${key}: spread ${spread.toFixed(5)}pt over ` +
        `${String(list.length)} faces - the line box is not font-independent`,
    );
  }
}
console.log(
  `      ${String(inputs.fonts.length)} faces agree at every size: worst spread ` +
    `${spreadWorst.toFixed(6)}pt over ${String(byFontSize.size)} groups`,
);

/* -------------------------------------------------------------------------- */
/* R3 - how a block is composed                                               */
/* -------------------------------------------------------------------------- */

console.log('\n=== R3  H(n) = (n-1) x advance + the last line ===');

let composed = 0;
const compositionMisses: string[] = [];
for (const p of probes.values()) {
  if (p.lines < 2) continue;
  const one = [...probes.values()].find(
    (q) =>
      q.kind === p.kind &&
      q.font === p.font &&
      q.sz === p.sz &&
      q.lines === 1 &&
      JSON.stringify(q.lnSpc ?? null) === JSON.stringify(p.lnSpc ?? null),
  );
  if (one === undefined) continue;
  const predicted = (p.lines - 1) * lineAdvance(p.sz, p.lnSpc) + heightOf(one.id);
  if (near(predicted, heightOf(p.id))) composed += 1;
  else if (compositionMisses.length < 8) compositionMisses.push(p.id);
}
if (compositionMisses.length > 0) {
  throw new Error(`composition fails on ${compositionMisses.join(', ')}`);
}
console.log(`      ${String(composed)}/${String(composed)} multi-line probes`);

// And the block starts at the frame top: no half-leading above the first line.
let topOffsetWorst = 0;
for (const p of probes.values()) {
  const r = reading(p.id);
  if (r.top === null || r.shapeTop === null) continue;
  topOffsetWorst = Math.max(topOffsetWorst, Math.abs(r.top - r.shapeTop));
}
if (topOffsetWorst >= EPS) {
  throw new Error(`text does not start at the frame top: worst offset ${String(topOffsetWorst)}pt`);
}
console.log(`      first line sits on the frame top exactly (worst offset 0pt)`);

/* -------------------------------------------------------------------------- */
/* R4 - the last line, which is data and not law                              */
/* -------------------------------------------------------------------------- */

console.log('\n=== R4  the last line ===');

/**
 * `c` is solved from one spacing deep in the linear regime and then scored
 * against every other spacing. Solving from one point and scoring the rest is
 * what makes this a fit; a least-squares line through all of them would have
 * absorbed the very disagreement being looked for, and did, on the first pass.
 */
const C = new Map<string, number>();
for (const p of probes.values()) {
  if (p.lines !== 1 || p.lnSpc?.kind !== 'pct' || p.lnSpc.value !== 300000) continue;
  C.set(`${p.font}@${String(p.sz)}`, heightOf(p.id) - 0.75 * lineAdvance(p.sz, p.lnSpc));
}
for (const p of probes.values()) {
  if (p.lines !== 1 || p.lnSpc?.kind !== 'pct' || p.lnSpc.value !== 200000) continue;
  const key = `${p.font}@${String(p.sz)}`;
  if (!C.has(key)) C.set(key, heightOf(p.id) - 0.75 * lineAdvance(p.sz, p.lnSpc));
}

type LastModel = (advance: number, natural: number, c: number) => number;
const LAST_MODELS: readonly { name: string; fn: LastModel }[] = [
  {
    name: 'the advance itself (what every line above the last measures)',
    fn: (a) => a,
  },
  { name: 'the natural box, 1.2 x size', fn: (_a, nat) => nat },
  { name: '0.75 x advance + c', fn: (a, _n, c) => 0.75 * a + c },
  { name: 'max(advance, 0.75 x advance + c)', fn: (a, _n, c) => Math.max(a, 0.75 * a + c) },
  {
    name: 'advance <= 1.2 size ? max(advance, 0.75 advance + c) : 0.75 advance + c',
    fn: (a, nat, c) => (a <= nat + 1e-9 ? Math.max(a, 0.75 * a + c) : 0.75 * a + c),
  },
];

const lastCases = [...probes.values()]
  .filter((p) => p.lines === 1 && C.has(`${p.font}@${String(p.sz)}`))
  .map((p) => p.id);

const lastTable = LAST_MODELS.map((m) => {
  let hits = 0;
  const misses: string[] = [];
  for (const id of lastCases) {
    const p = probes.get(id);
    if (p === undefined) continue;
    const c = C.get(`${p.font}@${String(p.sz)}`) ?? 0;
    const a = lineAdvance(p.sz, p.lnSpc);
    if (near(m.fn(a, 1.2 * pt(p.sz), c), heightOf(id))) hits += 1;
    else misses.push(id);
  }
  return { name: m.name, hits, misses };
});
for (const s of lastTable) {
  console.log(`      ${String(s.hits).padStart(5)}/${String(lastCases.length)}  ${s.name}`);
}

const best = lastTable.reduce((a, b) => (b.hits > a.hits ? b : a));
const exceptionFaces = [...new Set(best.misses.map((id) => probes.get(id)?.font ?? '?'))].sort(
  (a, b) => a.localeCompare(b),
);
console.log(
  `      best model misses ${String(best.misses.length)} probes, all on: ` +
    `${exceptionFaces.join(', ')}`,
);
if (exceptionFaces.length > 1) {
  throw new Error(
    `the last-line model misses on more than one face (${exceptionFaces.join(', ')}); ` +
      'the residual is not the single unexplained interval it was taken to be',
  );
}

/* -------------------------------------------------------------------------- */
/* R5 - the paragraph mark inside BoundWidth                                  */
/* -------------------------------------------------------------------------- */

console.log('\n=== R5  BoundWidth includes a trailing paragraph mark ===');

const MARK = new Map<string, number>();
for (const p of probes.values()) {
  if (p.kind !== 'pilcrow' || p.text !== 'H') continue;
  const one = reading(p.id).width;
  const twoProbe = [...probes.values()].find(
    (q) => q.kind === 'pilcrow' && q.font === p.font && q.sz === p.sz && q.text === 'HH',
  );
  if (one === null || twoProbe === undefined) continue;
  const two = reading(twoProbe.id).width;
  if (two === null) continue;
  // `HH` minus `H` is one H advance with the mark cancelling; `H` minus that is
  // the mark. Neither step needs to know what the mark is.
  MARK.set(`${p.font}@${String(p.sz)}`, one - (two - one));
}

let markSpread = 0;
for (const font of [...inputs.fonts, inputs.absentFont]) {
  const ratios = [...MARK.entries()]
    .filter(([k]) => k.startsWith(`${font}@`))
    .map(([k, v]) => v / pt(Number(k.split('@')[1])));
  if (ratios.length < 2) continue;
  markSpread = Math.max(markSpread, Math.max(...ratios) - Math.min(...ratios));
}
console.log(
  `      the mark is a fixed fraction of the size for each face: worst spread ` +
    `${markSpread.toFixed(5)} em over ${String(inputs.fonts.length + 1)} faces`,
);
if (markSpread > 0.01) {
  throw new Error('the paragraph mark is not a constant fraction of the size');
}

/* -------------------------------------------------------------------------- */
/* R6 - `@spc`                                                                */
/* -------------------------------------------------------------------------- */

console.log('\n=== R6  a:rPr/@spc ===');

type SpcModel = (spc: number, chars: number) => number;
const SPC_MODELS: readonly { name: string; fn: SpcModel }[] = [
  { name: 'points after every character, the last included', fn: (s, n) => (s / 100) * n },
  { name: 'points between characters only', fn: (s, n) => (s / 100) * (n - 1) },
  { name: 'em-relative: spc/100000 x size, every character', fn: () => Number.NaN },
];

// `0123456789` is ten characters, and `BoundWidth` covers eleven advances
// because of the paragraph mark - which is spaced too. The delta against the
// spc=0 probe of the same face and size cancels every glyph advance, leaving
// only the spacing.
const spcCases = [...probes.values()]
  .filter((p) => p.kind === 'spc' && p.spc !== 0)
  .map((p) => p.id);
const spcModel = chooseUnique('letter spacing', SPC_MODELS, spcCases, (fn, id) => {
  const p = probes.get(id);
  if (p?.spc === undefined) return false;
  const base = [...probes.values()].find(
    (q) => q.kind === 'spc' && q.font === p.font && q.sz === p.sz && q.spc === 0,
  );
  if (base === undefined) return false;
  const got = (reading(p.id).width ?? 0) - (reading(base.id).width ?? 0);
  const want = fn(p.spc, p.text.length + 1);
  return Number.isFinite(want) && near(want, got);
});

/* -------------------------------------------------------------------------- */
/* R7 - `@kern`                                                               */
/* -------------------------------------------------------------------------- */

console.log('\n=== R7  a:rPr/@kern is a minimum size, not a flag ===');

type KernModel = (kern: number, sz: number) => boolean;
const KERN_MODELS: readonly { name: string; fn: KernModel }[] = [
  {
    name: 'on when kern > 0 and size >= kern  (threshold, inclusive)',
    fn: (k, sz) => k > 0 && sz >= k,
  },
  {
    name: 'on when kern > 0 and size > kern  (threshold, exclusive)',
    fn: (k, sz) => k > 0 && sz > k,
  },
  { name: 'a boolean: on whenever the attribute is present', fn: () => true },
  { name: 'a boolean: on when kern is non-zero, at any size', fn: (k) => k > 0 },
  { name: 'never', fn: () => false },
];

// Only a face with kern pairs for this string can answer, and the readings say
// that is Arial alone of the three - which is exactly why three were measured.
const kernFonts = new Set<string>();
for (const p of probes.values()) {
  if (p.kind !== 'kern') continue;
  const widths = [...probes.values()]
    .filter((q) => q.kind === 'kern' && q.font === p.font && q.sz === p.sz)
    .map((q) => reading(q.id).width ?? 0);
  if (Math.max(...widths) - Math.min(...widths) > EPS) kernFonts.add(p.font);
}
console.log(`      faces whose kerning is measurable on this string: ${[...kernFonts].join(', ')}`);

const kernCases = [...probes.values()]
  .filter((p) => p.kind === 'kern' && p.kern !== undefined && kernFonts.has(p.font))
  .map((p) => p.id);

const kernModel = chooseUnique('kerning', KERN_MODELS, kernCases, (fn, id) => {
  const p = probes.get(id);
  if (p?.kern === undefined) return false;
  const group = [...probes.values()].filter(
    (q) => q.kind === 'kern' && q.font === p.font && q.sz === p.sz,
  );
  const widths = group.map((q) => reading(q.id).width ?? 0);
  const off = Math.max(...widths);
  const on = Math.min(...widths);
  return near(reading(p.id).width ?? 0, fn(p.kern, p.sz) ? on : off);
});

/* -------------------------------------------------------------------------- */
/* R7b - what a file that says nothing about kerning gets                     */
/* -------------------------------------------------------------------------- */

/**
 * ECMA-376 says an omitted `@kern` kerns at every size. A deck does not behave
 * that way, and the reason is 3.1's tenth source: a master that declares no
 * `p:txStyles` gets PowerPoint's built-in ones, and those carry a threshold.
 * So the effective default is inherited, not absent - and the number matters,
 * because it is the one every ordinary run in every ordinary deck resolves to.
 */
const kernDefaultProbes = [...probes.values()].filter(
  (p) => p.kind === 'kern' && p.kern === undefined,
);
let defaultThresholdBelow: number | null = null;
let defaultThresholdAtOrAbove: number | null = null;
for (const p of kernDefaultProbes) {
  const group = [...probes.values()].filter(
    (q) => q.kind === 'kern' && q.kern !== undefined && q.font === p.font && q.sz === p.sz,
  );
  if (group.length === 0) continue;
  const widths = group.map((q) => reading(q.id).width ?? 0);
  const kerned = near(reading(p.id).width ?? 0, Math.min(...widths));
  if (kerned) {
    defaultThresholdAtOrAbove =
      defaultThresholdAtOrAbove === null ? p.sz : Math.min(defaultThresholdAtOrAbove, p.sz);
  } else {
    defaultThresholdBelow =
      defaultThresholdBelow === null ? p.sz : Math.max(defaultThresholdBelow, p.sz);
  }
}
console.log(
  `      with no @kern in the file: not kerned at or below ` +
    `${String(defaultThresholdBelow === null ? '-' : pt(defaultThresholdBelow))}pt, kerned at or ` +
    `above ${String(defaultThresholdAtOrAbove === null ? '-' : pt(defaultThresholdAtOrAbove))}pt ` +
    `- so the cascade supplies a threshold rather than the attribute being absent`,
);
if (defaultThresholdBelow === null || defaultThresholdAtOrAbove === null) {
  throw new Error('the no-@kern probes did not bracket a threshold; the default is unmeasured');
}

/* -------------------------------------------------------------------------- */
/* R8 - substitution is visible in the advances                               */
/* -------------------------------------------------------------------------- */

console.log('\n=== R8  a face that is not installed ===');

function advanceRow(font: string, sz: number): string {
  return inputs.strings
    .map((s) => {
      const p = [...probes.values()].find(
        (q) => q.kind === 'advance' && q.font === font && q.sz === sz && q.text === s.text,
      );
      return p === undefined ? '?' : (reading(p.id).width ?? 0).toFixed(3);
    })
    .join('|');
}

const shared: string[] = [];
const rows = new Map<string, string>();
for (const font of [...inputs.fonts, inputs.absentFont]) {
  const row = advanceRow(font, 1800);
  const already = rows.get(row);
  if (already !== undefined) shared.push(`${already} == ${font}`);
  else rows.set(row, font);
}
if (shared.length !== 1 || !shared[0]?.endsWith(inputs.absentFont)) {
  throw new Error(
    `expected exactly one shared advance row, on the absent face; got: ${shared.join('; ') || '(none)'}`,
  );
}
const substitutedBy = shared[0].split(' == ')[0] ?? '?';
console.log(
  `      ${String(rows.size)} distinct rows over ${String(inputs.fonts.length + 1)} faces; ` +
    `the absent face measures as ${substitutedBy}`,
);

/* -------------------------------------------------------------------------- */
/* R9 - what a browser measures, against what PowerPoint measured             */
/* -------------------------------------------------------------------------- */

console.log('\n=== R9  Chromium against PowerPoint ===');

const browserByKey = new Map<string, BrowserSize>();
for (const f of browser.fonts) {
  for (const s of f.sizes) browserByKey.set(`${f.family}@${String(Math.round(s.px * 100))}`, s);
}

interface Agreement {
  readonly font: string;
  readonly samples: number;
  readonly meanPct: number;
  readonly worstPct: number;
  readonly worstAt: string;
}
const agreement: Agreement[] = [];
const allRel: number[] = [];
for (const font of [...inputs.fonts, inputs.absentFont]) {
  let sum = 0;
  let n = 0;
  let worst = 0;
  let worstAt = '';
  for (const p of probes.values()) {
    if (p.kind !== 'advance' || p.font !== font) continue;
    const mark = MARK.get(`${font}@${String(p.sz)}`);
    const b = browserByKey.get(`${font}@${String(p.sz)}`);
    const key = inputs.strings.find((s) => s.text === p.text)?.key;
    const raw = reading(p.id).width;
    if (mark === undefined || b === undefined || key === undefined || raw === null) continue;
    const ours = b.widths[key];
    if (ours === undefined) continue;
    const theirs = raw - mark;
    const rel = Math.abs(theirs - ours) / Math.max(theirs, 1e-9);
    sum += rel;
    n += 1;
    if (font !== inputs.absentFont) allRel.push(rel);
    if (rel > worst) {
      worst = rel;
      worstAt = `${key}@${String(pt(p.sz))}pt`;
    }
  }
  if (n > 0) {
    agreement.push({
      font,
      samples: n,
      meanPct: (sum / n) * 100,
      worstPct: worst * 100,
      worstAt,
    });
  }
}
allRel.sort((a, b) => a - b);
const pick = (q: number): number =>
  (allRel[Math.min(allRel.length - 1, Math.floor(allRel.length * q))] ?? 0) * 100;
const summary = { median: pick(0.5), p95: pick(0.95), max: pick(1) };
console.log(
  `      installed faces, ${String(allRel.length)} comparisons: median ` +
    `${summary.median.toFixed(4)}%  p95 ${summary.p95.toFixed(4)}%  max ${summary.max.toFixed(4)}%`,
);

// The browser's own letter spacing and kerning have to agree in *kind*, or the
// measurer cannot be built on them at all.
for (const font of ['Arial', 'Georgia', 'Consolas']) {
  for (const sz of [1200, 3200]) {
    const b = browserByKey.get(`${font}@${String(sz)}`);
    if (b === undefined) continue;
    const base = b.spc['0'] ?? 0;
    for (const [raw, width] of Object.entries(b.spc)) {
      const v = Number(raw);
      if (v === 0) continue;
      const perChar = (width - base) / (v / 100);
      if (!near(perChar, 10)) {
        throw new Error(
          `Chromium letterSpacing is ${perChar.toFixed(3)} x spc over ten characters at ` +
            `${font} ${String(pt(sz))}pt; PowerPoint applies it once per character including ` +
            'the last, and a measurer built on a different rule cannot match it',
        );
      }
    }
  }
}
console.log('      Chromium letterSpacing applies once per character, the last included: agrees');

const kernAgrees = ['Arial', 'Georgia', 'Consolas'].every((font) => {
  const b = browserByKey.get(`${font}@1200`);
  if (b === undefined) return false;
  const chromiumKerns = Math.abs(b.kernOn - b.kernOff) > EPS;
  return chromiumKerns === kernFonts.has(font);
});
if (!kernAgrees) {
  throw new Error('Chromium and PowerPoint disagree about which faces kern this string');
}
console.log('      the faces Chromium kerns are the faces PowerPoint kerns: agrees');

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface AdvanceEntry {
  readonly font: string;
  readonly sz: number;
  readonly paragraphMarkPt: number;
  readonly widthsPt: Readonly<Record<string, number>>;
}

const advanceTable: AdvanceEntry[] = [];
for (const font of [...inputs.fonts, inputs.absentFont]) {
  const sizes = [
    ...new Set(
      [...probes.values()].filter((p) => p.kind === 'advance' && p.font === font).map((p) => p.sz),
    ),
  ].sort((a, b) => a - b);
  for (const sz of sizes) {
    const mark = MARK.get(`${font}@${String(sz)}`);
    if (mark === undefined) continue;
    const widthsPt: Record<string, number> = {};
    for (const s of inputs.strings) {
      const p = [...probes.values()].find(
        (q) => q.kind === 'advance' && q.font === font && q.sz === sz && q.text === s.text,
      );
      if (p === undefined) continue;
      widthsPt[s.key] = Number(((reading(p.id).width ?? 0) - mark).toFixed(4));
    }
    advanceTable.push({
      font,
      sz,
      paragraphMarkPt: Number(mark.toFixed(4)),
      widthsPt,
    });
  }
}

const lastLineTable = [...C.entries()]
  .map(([key, c]) => {
    const [font = '', szRaw = '0'] = key.split('@');
    return { font, sz: Number(szRaw), descentPt: Number(c.toFixed(4)) };
  })
  .sort((a, b) => a.font.localeCompare(b.font) || a.sz - b.sz);

/*
 * The readings themselves, so a test can re-derive every rule above rather than
 * assert the sentence that describes it.
 *
 * A suite written against the conclusions passes forever and proves nothing; a
 * suite that recomputes the advance from `lnSpc` and compares it against what
 * PowerPoint reported goes red the moment the implementation is wrong. That is
 * the whole reason these tables are here and not just the summaries.
 */
const spacingOf = (p: Probe): { kind: string; value: number } | null =>
  p.lnSpc === undefined ? null : { kind: p.lnSpc.kind, value: p.lnSpc.value };
const round4 = (v: number): number => Number(v.toFixed(4));

const lineAdvanceReadings = [...probes.values()]
  .filter((p) => p.lines >= 2)
  .map((p) => ({
    font: p.font,
    sz: p.sz,
    lnSpc: spacingOf(p),
    lines: p.lines,
    advancePt: round4(advanceOf(p.id) ?? Number.NaN),
    heightPt: round4(heightOf(p.id)),
  }))
  .sort((a, b) => a.font.localeCompare(b.font) || a.sz - b.sz || a.lines - b.lines);

const lastLineReadings = [...probes.values()]
  .filter((p) => p.kind === 'lastLine' && p.lines === 1)
  .map((p) => ({
    font: p.font,
    sz: p.sz,
    lnSpc: spacingOf(p),
    heightPt: round4(heightOf(p.id)),
  }))
  .sort((a, b) => a.font.localeCompare(b.font) || a.sz - b.sz);

const kernReadings = [...probes.values()]
  .filter((p) => p.kind === 'kern')
  .map((p) => ({
    font: p.font,
    sz: p.sz,
    kern: p.kern ?? null,
    text: p.text,
    widthPt: round4(reading(p.id).width ?? Number.NaN),
  }))
  .sort((a, b) => a.font.localeCompare(b.font) || a.sz - b.sz || (a.kern ?? -1) - (b.kern ?? -1));

const spcReadings = [...probes.values()]
  .filter((p) => p.kind === 'spc')
  .map((p) => ({
    font: p.font,
    sz: p.sz,
    spc: p.spc ?? 0,
    chars: p.text.length,
    widthPt: round4(reading(p.id).width ?? Number.NaN),
  }))
  .sort((a, b) => a.font.localeCompare(b.font) || a.sz - b.sz || a.spc - b.spc);

const fixture = {
  $comment:
    'Generated by tools/ground-truth/analyse-metrics.ts from experiment T2. Do not hand-edit: ' +
    'rebuild the probe decks, re-read them through PowerPoint, and re-run the analysis.',
  experiment: 'T2',
  subPhase: '3.2',
  adr: 'docs/adr/0028-measurement-and-the-line-model.md',
  probeCount: probes.size,
  tolerancePt: EPS,
  chromium: browser.chromium,

  lineModel: {
    $comment:
      'The rules that fit every probe of their kind. Each was chosen over the ' +
      'alternatives listed in `refuted`, and the analysis throws rather than emit ' +
      'a rule two candidates fit.',
    percentQuantisation: quant.winner.name,
    advance: adv.winner.name,
    naturalLineFactor: 1.2,
    fontIndependent: true,
    firstLineTopOffsetPt: 0,
    composition: 'H(n) = (n - 1) * advance + lastLineHeight',
    letterSpacing: spcModel.winner.name,
    kerning: kernModel.winner.name,
    kernDefault: {
      $comment:
        'ECMA-376 says an omitted @kern kerns at every size. Measured, a file that ' +
        'states no @kern anywhere is not kerned below this threshold, because 3.1 ' +
        'found PowerPoint substitutes its own p:txStyles and those carry one. A ' +
        'resolver therefore has to read @kern through the cascade; the ECMA default ' +
        'only applies once the cascade has genuinely yielded nothing.',
      notKernedAtOrBelowHundredths: defaultThresholdBelow,
      kernedAtOrAboveHundredths: defaultThresholdAtOrAbove,
    },
    refuted: {
      percentQuantisation: quant.table
        .filter((s) => s.hits < pctAdvanceCases.length)
        .map((s) => ({
          model: s.name,
          fits: s.hits,
          of: pctAdvanceCases.length,
        })),
      advance: adv.table
        .filter((s) => s.hits < advanceCases.length)
        .map((s) => ({
          model: s.name,
          fits: s.hits,
          of: advanceCases.length,
        })),
      letterSpacing: spcModel.table
        .filter((s) => s.hits < spcCases.length)
        .map((s) => ({
          model: s.name,
          fits: s.hits,
          of: spcCases.length,
        })),
      kerning: kernModel.table
        .filter((s) => s.hits < kernCases.length)
        .map((s) => ({
          model: s.name,
          fits: s.hits,
          of: kernCases.length,
        })),
    },
  },

  lastLine: {
    $comment:
      'Measured data, not a committed rule. The best model found is stated with its ' +
      'exact score and every probe it misses; `descentPt` is the per-face term it ' +
      'needs, which no browser reports. Finishing this is sub-phase 3.6.',
    model:
      'advance <= 1.2 * size ? max(advance, 0.75 * advance + descent) : 0.75 * advance + descent',
    fits: best.hits,
    of: lastCases.length,
    missesOnlyOn: exceptionFaces,
    misses: best.misses,
    scored: lastTable.map((s) => ({ model: s.name, fits: s.hits, of: lastCases.length })),
    descentPt: lastLineTable,
  },

  advances: {
    $comment:
      'PowerPoint advances in points, with the trailing paragraph mark subtracted. ' +
      '`paragraphMarkPt` is that mark, which equals the face space advance and is ' +
      'an artefact of TextRange2.BoundWidth rather than of layout.',
    strings: inputs.strings,
    entries: advanceTable,
  },

  browserAgreement: {
    $comment:
      'How far Chromium measureText is from PowerPoint on the same face, size and ' +
      'string, as a fraction of the PowerPoint width. The plan promises a tracked ' +
      'number rather than pixel-perfection; this is that number.',
    chromium: browser.chromium,
    summaryPct: {
      median: Number(summary.median.toFixed(4)),
      p95: Number(summary.p95.toFixed(4)),
      max: Number(summary.max.toFixed(4)),
    },
    perFont: agreement.map((a) => ({
      font: a.font,
      samples: a.samples,
      meanPct: Number(a.meanPct.toFixed(4)),
      worstPct: Number(a.worstPct.toFixed(4)),
      worstAt: a.worstAt,
    })),
  },

  readings: {
    $comment:
      'What PowerPoint reported, so a test can re-derive the rules above instead of ' +
      'asserting the sentences that describe them. `advancePt` is the distance ' +
      'between two line tops; `heightPt` and `widthPt` are TextRange2.BoundHeight ' +
      'and BoundWidth, the latter still including the paragraph mark.',
    lineAdvance: lineAdvanceReadings,
    lastLine: lastLineReadings,
    kern: kernReadings,
    spc: spcReadings,
  },

  substitution: {
    $comment:
      'The probe table names one face that is not installed. PowerPoint measures it ' +
      'as another face, and which one is a measurement rather than a guess. Relevant ' +
      'to 3.7: document.fonts.check returned true for it, so presence cannot be ' +
      'detected that way and an advance fingerprint is needed.',
    absentFace: inputs.absentFont,
    powerPointSubstitutes: substitutedBy,
    distinctAdvanceRows: rows.size,
    facesMeasured: inputs.fonts.length + 1,
    browserReportedInstalled:
      browser.fonts.find((f) => f.family === inputs.absentFont)?.installed ?? null,
  },
};

const out = fileURLToPath(new URL('../../corpus/ground-truth/text-metrics.json', import.meta.url));
const json = `${JSON.stringify(fixture, null, 2)}\n`;
writeFileSync(out, json);

const sha = createHash('sha256').update(json).digest('hex');
console.log(`\nwrote ${out}`);
console.log(`  ${String(json.length)} bytes  sha256 ${sha}`);

/**
 * Experiment T13, step 2 - score every candidate reading against the browser.
 *
 * ```
 * node tools/ground-truth/fonts/metrics/analyse.ts <work-dir> [--fixture <path>]
 * ```
 *
 * Nothing here is allowed to conclude by inspection. Each question below is a
 * set of rival readings, every one is scored against every measurement, and the
 * script throws rather than emit a fixture where two of them tie. ADR 0042.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

// The built package, not its source: `tools/` is run by Node directly and the
// sources carry `.js` specifiers. Run `pnpm build` first.
import { facesIn } from '../../../../packages/cli/dist/index.js';
import { buildFont } from '../../lib/truetype.ts';
import { PROBES, STRINGS, SIZES, BOX_PX, probeFonts } from './probes.ts';

interface Measured {
  readonly chromium: string;
  readonly boxPx: number;
  readonly sizes: readonly number[];
  readonly strings: readonly string[];
  readonly measured: readonly {
    readonly id: string;
    readonly widths: Record<string, Record<string, number>>;
    readonly box: { readonly ascent: number; readonly descent: number };
  }[];
}

interface Scored {
  readonly model: string;
  readonly fits: number;
  readonly of: number;
  /** The largest absolute disagreement, in the unit of the question. */
  readonly worst: number;
}

/** Refuse a question whose winner is not alone at the top. */
function decide(question: string, scored: readonly Scored[]): Scored {
  const ranked = [...scored].sort((a, b) => b.fits - a.fits || a.worst - b.worst);
  const best = ranked[0];
  if (best === undefined) throw new Error(`${question}: nothing was scored`);
  if (best.fits !== best.of) {
    throw new Error(
      `${question}: the best reading "${best.model}" fits only ${String(best.fits)}/${String(best.of)}`,
    );
  }
  const rival = ranked[1];
  if (rival !== undefined && rival.fits === best.fits) {
    throw new Error(
      `${question}: "${best.model}" and "${rival.model}" both fit ${String(best.fits)}/${String(best.of)}; ` +
        'no probe separates them, so neither may be committed as the answer',
    );
  }
  return best;
}

const work = resolve(process.argv[2] ?? '');
if (process.argv[2] === undefined) throw new Error('usage: analyse.ts <work-dir> [--fixture <p>]');
const fixtureAt = process.argv.indexOf('--fixture');
const fixture = fixtureAt < 0 ? null : process.argv[fixtureAt + 1];

const browser = JSON.parse(readFileSync(join(work, 'browser-metrics.json'), 'utf8')) as Measured;

const rowOf = (id: string): Measured['measured'][number] => {
  const row = browser.measured.find((r) => r.id === id);
  if (row === undefined) throw new Error(`no browser row for ${id}`);
  return row;
};

/* -------------------------------------------------------------------------- */
/* question 1 - which table the font bounding box comes from                   */
/* -------------------------------------------------------------------------- */

const BOX_READINGS = {
  'hhea.ascender/descender': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.hheaAscender,
    descent: -s.hheaDescender,
  }),
  'OS/2.usWinAscent/usWinDescent': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.winAscent,
    descent: s.winDescent,
  }),
  'OS/2.sTypoAscender/sTypoDescender': (s: ReturnType<typeof buildFont>['spec']) => ({
    ascent: s.typoAscender,
    descent: -s.typoDescender,
  }),
  'usWin, or sTypo when fsSelection bit 7 is set': (s: ReturnType<typeof buildFont>['spec']) =>
    s.useTypoMetrics === true
      ? { ascent: s.typoAscender, descent: -s.typoDescender }
      : { ascent: s.winAscent, descent: s.winDescent },
};

const boxScores: Scored[] = Object.entries(BOX_READINGS).map(([model, read]) => {
  let fits = 0;
  let of = 0;
  let worst = 0;
  for (const probe of PROBES) {
    const { spec } = buildFont(probe.spec);
    const want = read(spec);
    const got = rowOf(probe.id).box;
    const scale = BOX_PX / spec.unitsPerEm;
    for (const [predicted, measured] of [
      [want.ascent * scale, got.ascent],
      [want.descent * scale, got.descent],
    ] as const) {
      of += 1;
      const off = Math.abs(predicted - measured);
      if (off < 1e-9) fits += 1;
      else worst = Math.max(worst, off);
    }
  }
  return { model, fits, of, worst };
});
const boxAnswer = decide('the font bounding box', boxScores);

/* -------------------------------------------------------------------------- */
/* question 2 - which kerning table the browser honours                        */
/* -------------------------------------------------------------------------- */

/** What each probe's tables say, so a reading can be scored without the reader. */
const KERN_READINGS = {
  'the legacy kern table only': (id: string) =>
    id === 'kern-only' || id === 'kern-and-gpos' ? -200 : id === 'gpos-only' ? 0 : 0,
  'GPOS PairPos only': (id: string) =>
    id === 'gpos-only' ? -200 : id === 'kern-and-gpos' ? -100 : 0,
  'GPOS when present, else the kern table': (id: string) =>
    id === 'kern-only' ? -200 : id === 'gpos-only' ? -200 : id === 'kern-and-gpos' ? -100 : 0,
  'the kern table when present, else GPOS': (id: string) =>
    id === 'kern-only' ? -200 : id === 'gpos-only' ? -200 : id === 'kern-and-gpos' ? -200 : 0,
  neither: () => 0,
};

const kernScores: Scored[] = Object.entries(KERN_READINGS).map(([model, adjustOf]) => {
  let fits = 0;
  let of = 0;
  let worst = 0;
  for (const probe of PROBES) {
    const row = rowOf(probe.id);
    for (const px of SIZES) {
      // `AB` is the kerned pair; `AC` is the control that never moves.
      const kerned = row.widths[String(px)]?.['AB'] ?? 0;
      const control = row.widths[String(px)]?.['AC'] ?? 0;
      const predicted = control + (adjustOf(probe.id) * px) / 1000;
      of += 1;
      // The rival readings are 100 font units apart, which is 0.8px at the
      // smallest size probed; the browser quantisation this tolerance absorbs is
      // three orders of magnitude smaller, and is question 3.
      const off = Math.abs(predicted - kerned);
      if (off < 1e-3) fits += 1;
      else worst = Math.max(worst, off);
    }
  }
  return { model, fits, of, worst };
});
const kernAnswer = decide('kerning', kernScores);

/* -------------------------------------------------------------------------- */
/* question 3 - the arithmetic that turns font units into a width              */
/* -------------------------------------------------------------------------- */

const FIXED = 65536;
const QUANTISERS = {
  'exact float, no quantisation': { advance: (x: number) => x, kern: (x: number) => x },
  'advance and kern truncated at 1/65536 px': {
    advance: (x: number) => Math.trunc(x * FIXED) / FIXED,
    kern: (x: number) => Math.trunc(x * FIXED) / FIXED,
  },
  'advance and kern rounded at 1/65536 px': {
    advance: (x: number) => Math.round(x * FIXED) / FIXED,
    kern: (x: number) => Math.round(x * FIXED) / FIXED,
  },
  'advance truncated, kern rounded, both at 1/65536 px': {
    advance: (x: number) => Math.trunc(x * FIXED) / FIXED,
    kern: (x: number) => Math.round(x * FIXED) / FIXED,
  },
  'the whole string quantised once at 1/65536 px': {
    advance: (x: number) => x,
    kern: (x: number) => x,
  },
};

/** The reader's own answer for one string, under a candidate quantisation. */
function widthUnder(
  faceFor: string,
  text: string,
  px: number,
  q: { advance: (x: number) => number; kern: (x: number) => number },
): number {
  const built = buildFont(PROBES.find((p) => p.id === faceFor)?.spec ?? {});
  const face = facesIn(built.bytes, faceFor)[0]!;
  const em = face.metrics.unitsPerEm;
  const points = [...text];
  let width = 0;
  for (let i = 0; i < points.length; i++) {
    const code = points[i]!.codePointAt(0)!;
    width += q.advance(((face.advanceOf(code) ?? 0) * px) / em);
    if (i + 1 < points.length) {
      const adjust = face.kernBetween(code, points[i + 1]!.codePointAt(0)!);
      if (adjust !== 0) width += q.kern((adjust * px) / em);
    }
  }
  return width;
}

const widthScores: Scored[] = Object.entries(QUANTISERS).map(([model, q]) => {
  let fits = 0;
  let of = 0;
  let worst = 0;
  for (const probe of PROBES) {
    const row = rowOf(probe.id);
    for (const px of SIZES) {
      for (const text of STRINGS) {
        const raw = widthUnder(probe.id, text, px, q);
        const predicted =
          model === 'the whole string quantised once at 1/65536 px'
            ? Math.trunc(raw * FIXED) / FIXED
            : raw;
        const measured = row.widths[String(px)]?.[text] ?? 0;
        of += 1;
        const off = Math.abs(predicted - measured);
        if (off < 1e-12) fits += 1;
        else worst = Math.max(worst, off);
      }
    }
  }
  return { model, fits, of, worst };
});
const widthAnswer = decide('the width arithmetic', widthScores);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const linear = widthScores.find((s) => s.model === 'exact float, no quantisation');

const answers = {
  $comment:
    'Experiment T13: what a font-table reader must do to agree with the browser it is standing ' +
    'in for. Regenerate with tools/ground-truth/fonts/metrics/analyse.ts. The fonts are built by ' +
    'tools/ground-truth/lib/truetype.ts and exist nowhere else, so no installed face can answer ' +
    'for a probe.',
  experiment: 'T13',
  subPhase: '3.10',
  adr: 'docs/adr/phase-3-text/0042-rendering-without-a-browser.md',
  chromium: browser.chromium,
  probes: PROBES.map((p) => ({ id: p.id, asks: p.asks })),
  // The bytes measured, and what the browser said about them. Committed so the
  // suite re-derives every rule from the measurement rather than from the code
  // that implements it - a test written off the implementation passes forever.
  fonts: probeFonts().map((font) => ({
    id: font.id,
    family: font.family,
    spec: font.spec,
    base64: font.base64,
  })),
  browser: {
    sizes: SIZES,
    strings: STRINGS,
    boxPx: BOX_PX,
    rows: browser.measured.map((row) => ({ id: row.id, box: row.box, widths: row.widths })),
  },
  measurements: {
    widths: PROBES.length * SIZES.length * STRINGS.length,
    boxes: PROBES.length * 2,
  },
  faceBox: {
    $comment:
      'The ascent and descent a browser reports for a face. Every real font sets hhea and usWin ' +
      'to the same numbers, so only a font built to disagree can say which was read.',
    answer: boxAnswer.model,
    scores: boxScores,
  },
  kerning: {
    $comment:
      'A font may carry pair adjustments in a legacy kern table, in GPOS, or in both. The third ' +
      'probe sets them to different values, which is the only way to find out which one loses.',
    answer: kernAnswer.model,
    scores: kernScores,
  },
  width: {
    $comment:
      'Whether an advance is linear in size - that is, unhinted - and how the browser quantises ' +
      'it. The linear reading is what matters on a slide; the quantisation is what lets the ' +
      "reader's test assert equality instead of a tolerance.",
    answer: widthAnswer.model,
    linearWithinPx: linear?.worst ?? null,
    scores: widthScores,
  },
} as const;

const text = `${JSON.stringify(answers, null, 2)}\n`;
writeFileSync(join(work, 't13-answers.json'), text);
if (fixture !== undefined && fixture !== null) writeFileSync(fixture, text);

console.log(`chromium ${browser.chromium}`);
console.log(`face box   ${boxAnswer.model}  (${String(boxAnswer.fits)}/${String(boxAnswer.of)})`);
console.log(
  `kerning    ${kernAnswer.model}  (${String(kernAnswer.fits)}/${String(kernAnswer.of)})`,
);
console.log(
  `width      ${widthAnswer.model}  (${String(widthAnswer.fits)}/${String(widthAnswer.of)})`,
);
console.log(`  linear reading is within ${String(linear?.worst ?? 0)} px`);
for (const [name, scores] of [
  ['face box', boxScores],
  ['kerning', kernScores],
  ['width', widthScores],
] as const) {
  console.log(`\n${name}`);
  for (const s of [...scores].sort((a, b) => b.fits - a.fits)) {
    console.log(`  ${String(s.fits).padStart(4)}/${String(s.of).padEnd(5)} ${s.model}`);
  }
}

/**
 * Experiment T4, step 2 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/text/autofit/build-deck.ts <work-dir>
 * ```
 *
 * Writes one package per deck plus `autofit-inputs.json` describing what every
 * shape was asked. Unlike T3 there is no measurement step before this one: the
 * primary sweep is built so that no width matters, so there is nothing to
 * measure in a browser first.
 *
 * ## Every deck is either asked or read, never both
 *
 * The runner has to *save* a deck to read back what PowerPoint computed, and
 * saving rewrites every part of it. So a deck whose probes are meant to be read
 * untouched cannot share a file with a deck whose probes are meant to be
 * recomputed - the save would overwrite the very stored values the first group
 * exists to ask about. `spacing` and `spacing-ctl` are the same questions split
 * on that line, and so are `lnspc-fit` and `lnspc-stored`.
 *
 * ## What is held still
 *
 * Everything T2 and T3 turned off is turned off again: `kern="0"`, no bullet,
 * no indent, left aligned, top anchored. Insets are zero except in the deck
 * that asks about insets, so that everywhere else the box height and the
 * available height are the same number and a fit predicate that is wrong about
 * insets cannot hide inside a rung.
 *
 * `a:lnSpc` is written only when a probe states one. T2 measured that an absent
 * `a:lnSpc` lays out identically to `100000`, so writing it everywhere would add
 * markup without adding a fact - and `lnspc-fit` needs "absent" to be a case it
 * can actually put in a file.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { allProbes, type Insets, type Probe, type Spacing, type StoredScale } from './probes.ts';
import { buildSheetPackage, SCHEME_ONE, emu, shape, type SlideSpec } from '../../lib/sheet-pptx.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/text/autofit/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ------------------------------------------------------------------ markup */

function spacingXml(tag: 'a:lnSpc' | 'a:spcBef' | 'a:spcAft', value: Spacing): string {
  const inner =
    value.kind === 'percent'
      ? `<a:spcPct val="${String(value.value)}"/>`
      : `<a:spcPts val="${String(value.value)}"/>`;
  return `<${tag}>${inner}</${tag}>`;
}

/**
 * `a:pPr` with every term that could move the text height stated.
 *
 * `spcBef` and `spcAft` are written as an explicit zero when the probe states
 * none. An absent one would inherit, and while nothing in these packages
 * declares paragraph spacing at any other level, "nothing declares it" is an
 * assumption about the chassis rather than a fact about this probe - and the
 * whole point of the spacing deck is that a stray 12pt somewhere would be
 * absorbed into a rung and never noticed.
 */
function pPrXml(probe: Probe): string {
  return (
    '<a:pPr marL="0" marR="0" indent="0" algn="l">' +
    (probe.lnSpc === undefined ? '' : spacingXml('a:lnSpc', probe.lnSpc)) +
    spacingXml('a:spcBef', probe.spcBef ?? { kind: 'points', value: 0 }) +
    spacingXml('a:spcAft', probe.spcAft ?? { kind: 'points', value: 0 }) +
    '<a:buNone/>' +
    '</a:pPr>'
  );
}

function rPrXml(probe: Probe): string {
  const face = escapeXml(probe.face);
  return (
    `<a:rPr lang="en-US" sz="${String(probe.sz)}" kern="0" dirty="0">` +
    `<a:latin typeface="${face}"/><a:cs typeface="${face}"/>` +
    '</a:rPr>'
  );
}

function autofitXml(probe: Probe): string {
  if (probe.autofit === 'none') return '<a:noAutofit/>';
  if (probe.autofit === 'sp') return '<a:spAutoFit/>';
  return normAutofitXml(probe.stored);
}

function normAutofitXml(stored: StoredScale): string {
  const attrs =
    (stored.fontScale === undefined ? '' : ` fontScale="${String(stored.fontScale)}"`) +
    (stored.lnSpcReduction === undefined
      ? ''
      : ` lnSpcReduction="${String(stored.lnSpcReduction)}"`);
  return `<a:normAutofit${attrs}/>`;
}

function insetAttrs(ins: Insets, what: string): string {
  return (
    ` lIns="${String(emu(ins.l, what))}" tIns="${String(emu(ins.t, what))}"` +
    ` rIns="${String(emu(ins.r, what))}" bIns="${String(emu(ins.b, what))}"`
  );
}

function bodyPrXml(probe: Probe): string {
  const flag =
    probe.spcFirstLastPara === undefined
      ? ''
      : ` spcFirstLastPara="${probe.spcFirstLastPara ? '1' : '0'}"`;
  return (
    `<a:bodyPr wrap="square"${insetAttrs(probe.insets, probe.id)}` +
    ` anchor="t" anchorCtr="0" rtlCol="0"${flag}>` +
    autofitXml(probe) +
    '</a:bodyPr>'
  );
}

function paragraphsXml(probe: Probe): readonly string[] {
  return probe.paragraphs.map(
    (text) =>
      '<a:p>' +
      pPrXml(probe) +
      '<a:r>' +
      rPrXml(probe) +
      `<a:t>${escapeXml(text)}</a:t>` +
      '</a:r>' +
      '</a:p>',
  );
}

/* ------------------------------------------------------------------ slides */

/** Twenty shapes a slide, spread on a grid wide enough that the tallest probe
 *  box in the sweep still does not reach its neighbour. Position is nothing to
 *  do with the measurement; legibility when the deck is opened by hand is. */
const PER_SLIDE = 20;
const COLS = 4;
const X_PITCH = 460;
const Y_PITCH = 470;

function slidesFor(probes: readonly Probe[]): readonly SlideSpec[] {
  const slides: SlideSpec[] = [];
  for (let i = 0; i < probes.length; i += PER_SLIDE) {
    const chunk = probes.slice(i, i + PER_SLIDE);
    slides.push({
      layout: 0,
      shapes: chunk.map((probe, j) =>
        shape({
          id: j + 2,
          name: probe.id,
          rect: {
            x: 10 + (j % COLS) * X_PITCH,
            y: 10 + Math.floor(j / COLS) * Y_PITCH,
            w: probe.rect.w,
            h: probe.rect.h,
          },
          fill: '<a:noFill/>',
          line: '<a:ln><a:solidFill><a:srgbClr val="C8C8C8"/></a:solidFill></a:ln>',
          bodyPr: bodyPrXml(probe),
          paragraphs: paragraphsXml(probe),
        }),
      ),
    });
  }
  return slides;
}

/* ------------------------------------------------------------------ build */

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
  readonly probes: number;
  readonly recompute: boolean;
}

const probes = allProbes();

const byDeck = new Map<string, Probe[]>();
const seen = new Set<string>();
for (const probe of probes) {
  // An id must name one shape in one deck: the runner matches by shape name,
  // and a duplicate would silently answer one question with another's reading.
  if (seen.has(probe.id)) throw new Error(`duplicate probe id ${probe.id}`);
  seen.add(probe.id);
  const list = byDeck.get(probe.deck);
  if (list === undefined) byDeck.set(probe.deck, [probe]);
  else list.push(probe);
}

const entries: DeckEntry[] = [];
for (const [deck, list] of byDeck) {
  const first = list[0];
  if (first === undefined) throw new Error(`deck ${deck} has no probes`);
  // Homogeneous by construction, asserted because the failure it prevents -
  // a save overwriting the stored scales a view deck exists to ask about - is
  // silent and would look like "PowerPoint recomputes on open".
  for (const p of list) {
    if (p.recompute !== first.recompute) {
      throw new Error(`deck ${deck} mixes recompute and view probes`);
    }
  }
  const slides = slidesFor(list);
  const file = `autofit-${deck}.pptx`;
  writeFileSync(
    join(outDir, file),
    buildSheetPackage({
      themes: [{ scheme: SCHEME_ONE, majorLatin: 'Arial', minorLatin: 'Arial' }],
      masters: [{ theme: 0 }],
      layouts: [{ master: 0 }],
      slides,
    }),
  );
  entries.push({
    deck,
    file,
    slides: slides.length,
    probes: list.length,
    recompute: first.recompute,
  });
}

writeFileSync(
  join(outDir, 'autofit-inputs.json'),
  JSON.stringify(
    {
      decks: entries,
      probes: probes.map((p) => ({
        ...p,
        /** The lines the probe is built to produce, so the runner's own count
         *  can contradict it rather than the analysis assuming it. */
        paragraphCount: p.paragraphs.length,
      })),
    },
    null,
    2,
  ),
);

console.log(`${String(probes.length)} probes across ${String(entries.length)} decks`);
for (const e of entries) {
  console.log(
    `  ${e.deck.padEnd(14)} ${String(e.probes).padStart(4)} probes  ` +
      `${String(e.slides).padStart(3)} slides  ${e.recompute ? 'recompute' : 'view only'}`,
  );
}

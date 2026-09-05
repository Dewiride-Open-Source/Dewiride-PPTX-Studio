/**
 * Experiment T2, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/text/metrics/build-deck.ts <out-dir>
 * ```
 *
 * One shape per probe, named with the probe id so the COM reader can match by
 * name rather than by position. Every measurement-hostile default is turned off
 * explicitly in `a:bodyPr` and `a:pPr`: no wrap, no autofit, zero insets, top
 * anchor, zero paragraph spacing, no bullet, no indent. What is left varying is
 * only what the probe is asking about.
 *
 * `wrap="none"` matters most of all. Line breaking is sub-phase 3.3; if the
 * shape width could break a line, every width here would be a measurement of
 * the breaker instead of the measurer.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import {
  buildSheetPackage,
  EMU_PER_POINT,
  SCHEME_ONE,
  shape,
  type SlideSpec,
} from '../../lib/sheet-pptx.ts';
import {
  ABSENT_FONT,
  FONTS,
  LINE_SIZES,
  metricProbes,
  STRINGS,
  type LnSpc,
  type Probe,
} from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/text/metrics/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function lnSpcXml(lnSpc: LnSpc | undefined): string {
  if (lnSpc === undefined) return '';
  const inner =
    lnSpc.kind === 'pct'
      ? `<a:spcPct val="${String(lnSpc.value)}"/>`
      : `<a:spcPts val="${String(lnSpc.value)}"/>`;
  return `<a:lnSpc>${inner}</a:lnSpc>`;
}

/**
 * `a:pPr`, with every term that could add height set to zero.
 *
 * The children are in `CT_TextParagraphProperties` order - `lnSpc`, `spcBef`,
 * `spcAft`, then the bullet group - because 0.6 made `insertInOrder` the only
 * sanctioned way to add an OOXML child and a probe deck that ignores the rule
 * is a probe deck PowerPoint may repair.
 */
function pPrXml(probe: Probe): string {
  return (
    '<a:pPr marL="0" marR="0" indent="0" algn="l">' +
    lnSpcXml(probe.lnSpc) +
    '<a:spcBef><a:spcPts val="0"/></a:spcBef>' +
    '<a:spcAft><a:spcPts val="0"/></a:spcAft>' +
    '<a:buNone/>' +
    '</a:pPr>'
  );
}

/** `a:rPr`, carrying the size, the face, and whichever of `@spc`/`@kern` is under test. */
function rPrXml(probe: Probe, tag: 'a:rPr' | 'a:endParaRPr'): string {
  const face = escapeXml(probe.font);
  const attrs =
    `lang="en-US" sz="${String(probe.sz)}"` +
    (probe.kern === undefined ? '' : ` kern="${String(probe.kern)}"`) +
    (probe.spc === undefined ? '' : ` spc="${String(probe.spc)}"`) +
    ' dirty="0"';
  return (
    `<${tag} ${attrs}>` +
    `<a:latin typeface="${face}"/><a:ea typeface="${face}"/><a:cs typeface="${face}"/>` +
    `</${tag}>`
  );
}

function runXml(probe: Probe): string {
  // `xml:space` is not decoration here: the `trailing` string exists to ask
  // whether PowerPoint measures a trailing space, and without this attribute
  // the question never reaches PowerPoint at all.
  const preserve = probe.text !== probe.text.trim() ? ' xml:space="preserve"' : '';
  return (
    '<a:r>' + rPrXml(probe, 'a:rPr') + `<a:t${preserve}>${escapeXml(probe.text)}</a:t>` + '</a:r>'
  );
}

interface Body {
  readonly bodyPr: string;
  readonly paragraphs: readonly string[];
}

function txBody(probe: Probe): Body {
  const bodyPr =
    '<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" ' +
    'anchor="t" anchorCtr="0" rtlCol="0"><a:noAutofit/></a:bodyPr>';

  const paragraphs: string[] = [];
  if (probe.kind === 'paraLines') {
    for (let i = 0; i < probe.lines; i += 1) {
      paragraphs.push('<a:p>' + pPrXml(probe) + runXml(probe) + '</a:p>');
    }
  } else {
    // `a:br` carries its own `a:rPr`, and the break's size is what sets the
    // height of the line it ends. Giving it anything other than the run's own
    // properties would make the probe measure two sizes at once.
    const parts = [runXml(probe)];
    for (let i = 1; i < probe.lines; i += 1) {
      parts.push('<a:br>' + rPrXml(probe, 'a:rPr') + '</a:br>', runXml(probe));
    }
    paragraphs.push('<a:p>' + pPrXml(probe) + parts.join('') + '</a:p>');
  }

  return { bodyPr, paragraphs };
}

/** Ten shapes to a slide, stacked. Overlap is harmless - each is measured alone. */
const PER_SLIDE = 10;

function slidesFor(probes: readonly Probe[]): SlideSpec[] {
  const slides: SlideSpec[] = [];
  for (let i = 0; i < probes.length; i += PER_SLIDE) {
    const chunk = probes.slice(i, i + PER_SLIDE);
    slides.push({
      layout: 0,
      shapes: chunk.map((probe, j) =>
        shape({
          id: j + 2,
          name: probe.id,
          rect: { x: 10, y: 10 + j * 52, w: 940, h: 48 },
          // A probe with no fill and no line is invisible in PowerPoint, which
          // makes a deck nobody can eyeball if a reading looks wrong.
          fill: '<a:noFill/>',
          line: '<a:ln><a:solidFill><a:srgbClr val="C8C8C8"/></a:solidFill></a:ln>',
          ...txBody(probe),
        }),
      ),
    });
  }
  return slides;
}

const DECKS: readonly { readonly deck: string; readonly kinds: readonly Probe['kind'][] }[] = [
  { deck: 'advance', kinds: ['advance', 'pilcrow'] },
  { deck: 'lineheight', kinds: ['lineHeight'] },
  { deck: 'spacing', kinds: ['paraLines', 'lnSpc', 'spc', 'kern'] },
  { deck: 'lastline', kinds: ['lastLine'] },
  { deck: 'chars', kinds: ['chars'] },
];

const all = metricProbes();
const seen = new Set<string>();
for (const probe of all) {
  if (seen.has(probe.id)) throw new Error(`duplicate probe id: ${probe.id}`);
  seen.add(probe.id);
}

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
  readonly probes: number;
}

const entries: DeckEntry[] = [];
for (const spec of DECKS) {
  const probes = all.filter((p) => spec.kinds.includes(p.kind));
  if (probes.length === 0) throw new Error(`deck ${spec.deck} has no probes`);
  const slides = slidesFor(probes);
  const file = `metrics-${spec.deck}.pptx`;
  writeFileSync(
    join(outDir, file),
    buildSheetPackage({
      // The scheme is irrelevant to a width or a height, so it is the stock
      // one. The font scheme is not irrelevant: every probe names its face
      // outright, and a theme font that differed would be a second answer to
      // the same question if any probe ever forgot to.
      themes: [{ scheme: SCHEME_ONE, majorLatin: 'Arial', minorLatin: 'Arial' }],
      masters: [{ theme: 0 }],
      layouts: [{ master: 0 }],
      slides,
    }),
  );
  entries.push({ deck: spec.deck, file, slides: slides.length, probes: probes.length });
}

writeFileSync(
  join(outDir, 'metric-inputs.json'),
  JSON.stringify(
    {
      emuPerPoint: EMU_PER_POINT,
      fonts: FONTS,
      absentFont: ABSENT_FONT,
      lineSizes: LINE_SIZES,
      strings: STRINGS,
      decks: entries,
      probes: all,
    },
    null,
    2,
  ),
);

console.log(`${String(all.length)} probes across ${String(entries.length)} decks`);
for (const e of entries) {
  console.log(
    `  ${e.deck.padEnd(12)} ${String(e.probes).padStart(4)} probes  ${String(e.slides)} slides`,
  );
}

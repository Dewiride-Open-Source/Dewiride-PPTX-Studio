/**
 * Experiment T8, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/render/text/build-deck.ts <work-dir>
 * ```
 *
 * One package per question and one slide per probe, because an EMF names no
 * shapes: a slide's records are the only thing that identifies which probe drew
 * them. `text-render-inputs.json` is the index `read.ps1` walks.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildSheetPackage, SCHEME_ONE, shape, type SlideSpec } from '../../lib/sheet-pptx.ts';

import { allPackages, type Para, type Probe, type Run } from './probes.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: build-deck.ts <out-dir>');
const outDir: string = arg;
mkdirSync(outDir, { recursive: true });

/* -------------------------------------------------------------------------- */
/* markup                                                                     */
/* -------------------------------------------------------------------------- */

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Reject anything XML 1.0 cannot carry: PowerPoint refuses the package and names nothing. */
function assertProbeText(text: string, what: string): void {
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    const legal =
      code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x20 && code !== 0x7f);
    if (!legal) {
      throw new Error(
        `${what} ${JSON.stringify(text)} holds U+${code.toString(16).toUpperCase().padStart(4, '0')}, which XML 1.0 forbids`,
      );
    }
  }
}

const EMU_PER_POINT = 12700;

/** Hundredths of a point, as `a:rPr/@sz` wants them. */
function hundredths(points: number, what: string): number {
  const value = Math.round(points * 100);
  if (Math.abs(value - points * 100) > 1e-9) {
    throw new Error(`${what} ${String(points)}pt is not a whole hundredth of a point`);
  }
  return value;
}

/** `a:rPr`, with every attribute the run states and nothing it does not. */
function rPrXml(run: Run, tag: 'a:rPr' | 'a:endParaRPr' = 'a:rPr'): string {
  const attrs: string[] = ['lang="en-US"', 'dirty="0"'];
  if (run.szPt !== undefined) attrs.push(`sz="${String(hundredths(run.szPt, 'sz'))}"`);
  if (run.b !== undefined) attrs.push(`b="${run.b ? '1' : '0'}"`);
  if (run.i !== undefined) attrs.push(`i="${run.i ? '1' : '0'}"`);
  if (run.u !== undefined) attrs.push(`u="${run.u}"`);
  if (run.strike !== undefined) attrs.push(`strike="${run.strike}"`);
  if (run.cap !== undefined) attrs.push(`cap="${run.cap}"`);
  if (run.baseline !== undefined) attrs.push(`baseline="${String(run.baseline)}"`);
  if (run.spcPt !== undefined) attrs.push(`spc="${String(hundredths(run.spcPt, 'spc'))}"`);
  if (run.kernPt !== undefined) attrs.push(`kern="${String(hundredths(run.kernPt, 'kern'))}"`);
  const children =
    (run.color === undefined ? '' : `<a:solidFill><a:srgbClr val="${run.color}"/></a:solidFill>`) +
    (run.typeface === undefined
      ? ''
      : `<a:latin typeface="${escapeXml(run.typeface)}"/><a:cs typeface="${escapeXml(run.typeface)}"/>`);
  return children === ''
    ? `<${tag} ${attrs.join(' ')}/>`
    : `<${tag} ${attrs.join(' ')}>${children}</${tag}>`;
}

function paraXml(para: Para): string {
  const pPr = para.algn === undefined ? '' : `<a:pPr algn="${para.algn}"/>`;
  const runs = para.runs
    .map((run) => {
      assertProbeText(run.text, 'run text');
      // `xml:space="preserve"` or a trailing space is not in the file at all,
      // and Q4's trailing-space rows measure nothing.
      return `<a:r>${rPrXml(run)}<a:t xml:space="preserve">${escapeXml(run.text)}</a:t></a:r>`;
    })
    .join('');
  const last = para.runs.at(-1);
  const endPr = last === undefined ? '' : rPrXml(last, 'a:endParaRPr');
  return `<a:p>${pPr}${runs}${endPr}</a:p>`;
}

/** `a:bodyPr`, plus the `a:noAutofit` every probe wants so nothing is rescaled. */
function bodyPrXml(probe: Probe): string {
  const attrs: string[] = [];
  if (probe.bodyRotDeg !== undefined) {
    attrs.push(`rot="${String(Math.round(probe.bodyRotDeg * 60000))}"`);
  }
  if (probe.upright !== undefined) attrs.push(`upright="${probe.upright ? '1' : '0'}"`);
  if (probe.wrap !== undefined) attrs.push(`wrap="${probe.wrap}"`);
  if (probe.insetPt !== undefined) {
    const value = String(Math.round(probe.insetPt * EMU_PER_POINT));
    for (const edge of ['lIns', 'tIns', 'rIns', 'bIns']) attrs.push(`${edge}="${value}"`);
  }
  if (probe.anchor !== undefined) attrs.push(`anchor="${probe.anchor}"`);
  const head = attrs.length === 0 ? '<a:bodyPr' : `<a:bodyPr ${attrs.join(' ')}`;
  return `${head}><a:noAutofit/></a:bodyPr>`;
}

/** One probe as one slide: a reference rectangle, then the text shape over it. */
function slideFor(probe: Probe, index: number): SlideSpec {
  const id = index * 10 + 2;
  return {
    layout: 0,
    shapes: [
      shape({
        id,
        name: `${probe.id}-ref`,
        rect: probe.rect,
        fill: '<a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>',
        line: '<a:ln><a:noFill/></a:ln>',
      }),
      shape({
        id: id + 1,
        name: probe.id,
        rect: probe.rect,
        ...(probe.rotDeg === undefined || probe.rotDeg === 0
          ? {}
          : { rot: Math.round(probe.rotDeg * 60000) }),
        ...(probe.flipH === true ? { flipH: true } : {}),
        ...(probe.flipV === true ? { flipV: true } : {}),
        fill: '<a:noFill/>',
        line: '<a:ln><a:noFill/></a:ln>',
        bodyPr: bodyPrXml(probe),
        paragraphs: probe.paragraphs.map(paraXml),
      }),
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* the packages                                                               */
/* -------------------------------------------------------------------------- */

interface DeckInput {
  readonly deck: string;
  readonly question: string;
  readonly file: string;
  readonly slides: readonly { readonly slide: number; readonly probe: string }[];
}

const decks: DeckInput[] = [];

for (const pkg of allPackages()) {
  const slides = pkg.probes.map((probe, index) => slideFor(probe, index));
  const bytes = buildSheetPackage({
    themes: [{ scheme: SCHEME_ONE }],
    masters: [{ theme: 0, shapes: [] }],
    layouts: [{ master: 0, type: 'blank', name: 'Blank', shapes: [] }],
    slides,
  });
  const file = `t8-${pkg.deck}.pptx`;
  writeFileSync(join(outDir, file), bytes);
  decks.push({
    deck: pkg.deck,
    question: pkg.question,
    file,
    slides: pkg.probes.map((probe, index) => ({ slide: index + 1, probe: probe.id })),
  });
  console.log(`${pkg.deck.padEnd(12)} ${String(pkg.probes.length).padStart(3)} probe(s)  ${file}`);
}

const probeIndex = allPackages().flatMap((pkg) =>
  pkg.probes.map((probe) => ({
    id: probe.id,
    question: probe.question,
    deck: pkg.deck,
    rect: probe.rect,
    bitmap: probe.bitmap === true,
    asks: probe.asks,
  })),
);

writeFileSync(
  join(outDir, 'text-render-inputs.json'),
  `${JSON.stringify({ decks, probes: probeIndex }, null, 2)}\n`,
);
console.log(`\n${String(probeIndex.length)} probe(s) in ${String(decks.length)} package(s)`);

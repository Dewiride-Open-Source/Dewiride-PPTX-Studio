/**
 * Experiment C8, step 1 - one package per probe, built from the styles PowerPoint wrote in step 0.
 *
 * ```
 * node tools/ground-truth/model/tables/styles/build-deck.ts <out-dir>
 * ```
 *
 * Reads `author-styles-log.json` and `pp-styles-off.pptx` from the directory `author.ps1` wrote,
 * and writes `style-<id>.pptx` per probe plus `style-inputs.json` for `read.ps1`.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import {
  buildSheetPackage,
  IDENTITY_CLR_MAP,
  SCHEME_ONE,
  SCHEME_TWO,
  type ClrMapAttrs,
  type ThemeSpec,
} from '../../../lib/sheet-pptx.ts';
import {
  allProbes,
  catalogueFrom,
  EXPORT,
  RECORD_ONLY,
  SCHEME_THREE,
  TABLE_RECT,
  tableFrame,
  tableStylesOf,
  type AuthorLog,
  type Probe,
  type Theme,
} from './probes.ts';

const dir = process.argv[2];
if (dir === undefined) {
  throw new Error('usage: tools/ground-truth/model/tables/styles/build-deck.ts <out-dir>');
}

const THEMES: Readonly<Record<Theme, { theme: ThemeSpec; clrMap: ClrMapAttrs }>> = {
  s1: { theme: { scheme: SCHEME_ONE }, clrMap: IDENTITY_CLR_MAP },
  s2: {
    theme: { scheme: SCHEME_TWO, majorLatin: 'Georgia', minorLatin: 'Verdana' },
    clrMap: { ...IDENTITY_CLR_MAP, bg1: 'dk1', tx1: 'lt1', bg2: 'dk2', tx2: 'lt2' },
  },
  s3: {
    theme: { scheme: SCHEME_THREE, rgbDarkLight: true },
    clrMap: { ...IDENTITY_CLR_MAP, tx1: 'dk2' },
  },
};

/** Every pair of a question's candidates must disagree on some slide, or the question asks nothing. */
function assertSeparable(probes: readonly Probe[]): void {
  const byQuestion = new Map<string, Probe[]>();
  for (const probe of probes) {
    if (probe.control === true || RECORD_ONLY.has(probe.question)) continue;
    byQuestion.set(probe.question, [...(byQuestion.get(probe.question) ?? []), probe]);
  }
  for (const [question, members] of byQuestion) {
    const candidates = Object.keys(members[0]?.slides[0]?.predict ?? {});
    for (const probe of members) {
      for (const slide of probe.slides) {
        const here = Object.keys(slide.predict);
        if (here.join() !== candidates.join()) {
          throw new Error(
            `${probe.id} predicts for ${here.join()}, the question has ${candidates.join()}`,
          );
        }
      }
    }
    for (const a of candidates) {
      for (const b of candidates) {
        if (a >= b) continue;
        const apart = members.some((p) => p.slides.some((s) => s.predict[a] !== s.predict[b]));
        if (!apart) throw new Error(`${question}: no probe tells ${a} from ${b}`);
      }
    }
  }
}

/** Every `control:NAME` names a control slide in the same theme. */
function assertControls(probes: readonly Probe[]): void {
  const controls = new Map<string, Theme>();
  for (const probe of probes) {
    if (probe.control !== true) continue;
    if (probe.slides.length === 1) controls.set(probe.id, probe.theme);
    probe.slides.forEach((_, k) => controls.set(`${probe.id}#${String(k)}`, probe.theme));
  }
  for (const probe of probes) {
    for (const slide of probe.slides) {
      for (const observation of Object.values(slide.predict)) {
        if (!observation.startsWith('control:')) continue;
        const name = observation.slice('control:'.length);
        const theme = controls.get(name);
        if (theme === undefined)
          throw new Error(`${probe.id} predicts ${name}, which is no control`);
        if (theme !== probe.theme)
          throw new Error(`${probe.id} (${probe.theme}) is compared with ${name} (${theme})`);
      }
    }
  }
}

function build(probe: Probe): Uint8Array {
  const { theme, clrMap } = THEMES[probe.theme];
  return buildSheetPackage({
    themes: [theme],
    masters: [{ theme: 0, clrMap }],
    layouts: [{ master: 0, type: 'blank', name: 'Blank' }],
    slides: probe.slides.map((slide, k) => ({
      layout: 0,
      shapes: [tableFrame(`${probe.id}#${String(k)}`, slide.tblPr, slide.tcPr)],
    })),
    ...(probe.part === null
      ? {}
      : {
          tableStyles: {
            xml: probe.part.xml,
            partName: probe.part.partName,
            rel: probe.part.rel,
            part: probe.part.part,
          },
        }),
  });
}

const log = JSON.parse(readFileSync(join(dir, 'author-styles-log.json'), 'utf8')) as AuthorLog;
const authored = tableStylesOf(readFileSync(join(dir, 'pp-styles-off.pptx')));
if (authored === null) throw new Error('pp-styles-off.pptx has no table-style part');
const catalogue = catalogueFrom(log, authored);
const probes = allProbes(catalogue);
assertControls(probes);
assertSeparable(probes);

for (const probe of probes) writeFileSync(join(dir, `style-${probe.id}.pptx`), build(probe));

const inputs = {
  emuPerPoint: 12700,
  table: TABLE_RECT,
  export: EXPORT,
  reread: ['ctl-G1', 'ctl-G2', 'ctl-G3', 'ctl-sig-B'],
  reapply: {
    probes: ['sweep-nopart', 'sweep2-nopart'],
    through: catalogue.ids.get('No Style, No Grid'),
  },
  probes: [...probes]
    .sort((a, b) => Number(b.control === true) - Number(a.control === true))
    .map((probe) => ({
      id: probe.id,
      file: `style-${probe.id}.pptx`,
      question: probe.question,
      theme: probe.theme,
      control: probe.control === true,
      hostile: probe.hostile === true,
      sweep: probe.slides.length > 1,
      part: probe.part,
      slides: probe.slides,
    })),
};
writeFileSync(join(dir, 'style-inputs.json'), JSON.stringify(inputs, null, 2));
console.log(`wrote ${String(probes.length)} packages and style-inputs.json`);

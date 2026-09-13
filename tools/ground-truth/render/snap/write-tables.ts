/**
 * Experiment F3, step 4 - the fixture as Markdown tables for the ADR, on stdout.
 *
 * `node tools/ground-truth/render/snap/write-tables.ts [corpus/ground-truth/snap.json]`
 */

import { readFileSync } from 'node:fs';

import { repoPath } from '../../../repo/root.ts';

interface Score {
  readonly model: string;
  readonly description: string;
  readonly fits: number;
  readonly of: number;
  readonly worstError: number;
}

interface Measure {
  readonly from: number;
  readonly cover: readonly number[];
}

interface AtWidth {
  readonly height: number;
  readonly dpi: number;
  readonly eighth: boolean;
  readonly wholeHeight: boolean;
  readonly cases: number;
  readonly rule: Readonly<Record<string, readonly [number, number]>>;
  readonly ruleMisses: readonly string[];
  readonly mappings: Readonly<Record<string, Readonly<Record<string, readonly [number, number]>>>>;
  readonly partialRows: number;
  readonly quarterRows: number;
}

interface Fixture {
  readonly powerpoint: string;
  readonly widths: readonly number[];
  readonly gridWidths: readonly number[];
  readonly cases: number;
  readonly findings: Readonly<Record<string, unknown>>;
  readonly candidates: Readonly<Record<string, readonly Score[]>>;
  readonly atWidth: Readonly<Record<string, AtWidth>>;
  readonly measured: Readonly<Record<string, Readonly<Record<string, Measure>>>>;
}

const fixturePath = process.argv[2] ?? repoPath('corpus/ground-truth/snap.json');
const fixture = JSON.parse(readFileSync(fixturePath, 'utf8')) as Fixture;

const lines: string[] = [];
const out = (line = ''): void => {
  lines.push(line);
};

out(
  `F3 against PowerPoint ${fixture.powerpoint}: ${String(fixture.cases)} cases at ${fixture.widths.join(', ')} px wide, the rule read at ${fixture.gridWidths.join(', ')}.`,
);
out();

for (const [family, scores] of Object.entries(fixture.candidates)) {
  out(`### ${family}`);
  out();
  out('| model | reading | fits | worst error |');
  out('| ----- | ------- | ---: | ----------: |');
  for (const score of scores) {
    const mark = score.fits === score.of ? '**' : '';
    out(
      `| ${mark}${score.model}${mark} | ${score.description} | ${mark}${String(score.fits)}/${String(score.of)}${mark} | ${score.worstError.toFixed(3)} |`,
    );
  }
  out();
}

/** The inked rows of a profile, as `row: cover` pairs. */
function inked(measure: Measure): string {
  return measure.cover
    .map((cover, i) => [measure.from + i, cover] as const)
    .filter(([, cover]) => cover > 0)
    .map(([row, cover]) => `${String(row)}: ${cover.toFixed(2)}`)
    .join(', ');
}

function profiles(title: string, ids: readonly string[]): void {
  out(`### ${title}`);
  out();
  out('| probe | width | inked rows |');
  out('| ----- | ----: | ---------- |');
  for (const id of ids) {
    const row = fixture.measured[id];
    if (row === undefined) continue;
    for (const width of fixture.widths) {
      const measure = row[String(width)];
      if (measure !== undefined) out(`| ${id} | ${String(width)} | ${inked(measure)} |`);
    }
  }
  out();
}

profiles('The one-point stroke, four offsets', [
  'stroke-h-1-0',
  'stroke-h-1-0_25',
  'stroke-h-1-0_5',
  'stroke-h-1-0_75',
]);
profiles('The two-point stroke, and its 2.5-pixel pen at 1200', [
  'stroke-h-2-0',
  'stroke-h-2-0_25',
  'stroke-h-2-0_5',
  'stroke-h-2-0_75',
]);
profiles('The fill edge', ['fill-0-top', 'fill-0_25-top', 'fill-0_5-top', 'fill-0_75-top']);
profiles('The picture border', ['border-1-0', 'border-1-0_5', 'border-2-0', 'border-2-0_5']);
profiles('The inset stroke', ['inset-1-0', 'inset-1-0_5', 'inset-2-0', 'inset-2-0_5']);
profiles('The half pixels at 240 and 1920', ['tie-6-0', 'tie-10-0', 'tie-0_75-0', 'tie-1_25-0']);
profiles('The ellipse, top and bottom', ['ellipse-1-0', 'ellipse-1-0-bottom']);
profiles('The flat end of a one-point line', ['end-1-0-start', 'end-1-0-end']);

out('### Every width against the rule');
out();
out(
  '| width | dpi | height | eighth | rule fits | partial rows | of them quarters | misses | mappings |',
);
out(
  '| ----: | --: | -----: | ------ | --------: | -----------: | ---------------: | ------ | -------- |',
);
for (const [width, at] of Object.entries(fixture.atWidth)) {
  const fits = Object.values(at.rule).reduce((sum, [f]) => sum + f, 0);
  const mappings = Object.entries(at.mappings)
    .map(([name, byFamily]) => {
      const f = Object.values(byFamily).reduce((sum, [x]) => sum + x, 0);
      const of = Object.values(byFamily).reduce((sum, [, n]) => sum + n, 0);
      return `${name} ${String(f)}/${String(of)}`;
    })
    .join(', ');
  out(
    `| ${width} | ${String(at.dpi)} | ${String(at.height)} | ${at.eighth ? 'yes' : 'no'} | ${String(fits)}/${String(at.cases)} | ${String(at.partialRows)} | ${String(at.quarterRows)} | ${at.ruleMisses.join(' ') || (fits === at.cases ? '' : 'many')} | ${mappings} |`,
  );
}
out();

out('### Findings');
out();
for (const [key, value] of Object.entries(fixture.findings)) {
  out('- `' + key + '`: ' + JSON.stringify(value));
}

console.log(lines.join('\n'));

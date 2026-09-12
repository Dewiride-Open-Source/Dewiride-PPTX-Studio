/**
 * Experiment F2, step 4 - the fixture as Markdown tables for the ADR, on stdout.
 *
 * `node tools/ground-truth/render/zoom/write-tables.ts [corpus/ground-truth/zoom.json]`
 */

import { readFileSync } from 'node:fs';

import { repoPath } from '../../../repo/root.ts';

interface Score {
  readonly model: string;
  readonly description: string;
  readonly fits: number;
  readonly of: number;
  readonly worstErrorPx: number;
}

interface Fixture {
  readonly powerpoint: string;
  readonly widths: readonly number[];
  readonly cases: number;
  readonly findings: Readonly<Record<string, unknown>>;
  readonly candidates: Readonly<Record<string, readonly Score[]>>;
  readonly text: readonly {
    probe: string;
    width: number;
    boxW: number;
    boxH: number;
    withinPx: number;
  }[];
  readonly gradient: readonly { width: number; levels: number; worstLevels: number }[];
  readonly patternCoverage: readonly {
    probe: string;
    width: number;
    periodPx: number | null;
    mean: number;
  }[];
  readonly measured: Readonly<
    Record<string, Readonly<Record<string, Readonly<Record<string, unknown>>>>>
  >;
}

const fixturePath = process.argv[2] ?? repoPath('corpus/ground-truth/zoom.json');
const fixture = JSON.parse(readFileSync(fixturePath, 'utf8')) as Fixture;

const lines: string[] = [];
const out = (line = ''): void => {
  lines.push(line);
};

const cell = (value: unknown): string =>
  value === null || value === undefined
    ? '—'
    : typeof value === 'number'
      ? Number.isInteger(value)
        ? String(value)
        : value.toFixed(2)
      : JSON.stringify(value);

out(
  `F2 against PowerPoint ${fixture.powerpoint}: ${String(fixture.cases)} cases at ${fixture.widths.join(', ')} px wide.`,
);
out();

for (const [family, scores] of Object.entries(fixture.candidates)) {
  out(`### ${family}`);
  out();
  out('| model | reading | fits | worst error (px) |');
  out('| ----- | ------- | ---: | ---------------: |');
  for (const score of scores) {
    const mark = score.fits === score.of ? '**' : '';
    out(
      `| ${mark}${score.model}${mark} | ${score.description} | ${mark}${String(score.fits)}/${String(score.of)}${mark} | ${cell(score.worstErrorPx)} |`,
    );
  }
  out();
}

function byWidth(title: string, probes: readonly string[], key: string): void {
  out(`### ${title}`);
  out();
  out(`| probe | ${fixture.widths.map((w) => `${String(w)} px`).join(' | ')} |`);
  out(`| ----- | ${fixture.widths.map(() => '---:').join(' | ')} |`);
  for (const probe of probes) {
    const row = fixture.measured[probe];
    if (row === undefined) continue;
    out(`| ${probe} | ${fixture.widths.map((w) => cell(row[String(w)]?.[key])).join(' | ')} |`);
  }
  out();
}

const probesOf = (prefix: string): string[] =>
  Object.keys(fixture.measured).filter((id) => id.startsWith(prefix));

byWidth('Hairline ink, pixels per scanline', probesOf('hair-').concat(['border-hair']), 'inkPx');
byWidth('Thin stroke ink, pixels', probesOf('thin-'), 'inkPx');
byWidth('Picture border ink, pixels', ['border-1pt', 'border-1_1pt'], 'inkPx');
byWidth('Dash period, pixels', probesOf('dash-'), 'periodPx');
byWidth('Triangle head height, pixels', probesOf('marker-'), 'boxH');

out('### Pattern period and coverage');
out();
out('| probe | width | period (px) | coverage |');
out('| ----- | ----: | ----------: | -------: |');
for (const row of fixture.patternCoverage) {
  out(`| ${row.probe} | ${String(row.width)} | ${cell(row.periodPx)} | ${row.mean.toFixed(3)} |`);
}
out();

out('### Gradient ramp against the 960 export');
out();
out('| width | distinct levels | worst difference (levels) |');
out('| ----: | --------------: | ------------------------: |');
for (const row of fixture.gradient) {
  out(`| ${String(row.width)} | ${String(row.levels)} | ${String(row.worstLevels)} |`);
}
out();

out('### Text extent against a linear scale of the 960 export');
out();
out('| probe | worst deviation (px) | at width |');
out('| ----- | -------------------: | -------: |');
const worstByProbe = new Map<string, { withinPx: number; width: number }>();
for (const row of fixture.text) {
  const seen = worstByProbe.get(row.probe);
  if (seen === undefined || row.withinPx > seen.withinPx) {
    worstByProbe.set(row.probe, { withinPx: row.withinPx, width: row.width });
  }
}
for (const [probe, worst] of worstByProbe) {
  out(`| ${probe} | ${worst.withinPx.toFixed(2)} | ${String(worst.width)} |`);
}
out();

out('### Findings');
out();
for (const [key, value] of Object.entries(fixture.findings)) {
  out('- `' + key + '`: ' + JSON.stringify(value));
}

console.log(lines.join('\n'));

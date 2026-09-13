/**
 * Gate 3's run, as a page a person looks at.
 *
 * One column per zoom, worst first within it, and for the three worst at each zoom the page's
 * own screenshot beside PowerPoint's grid and the difference between them. The pictures are
 * files rather than data URIs so each can be opened and pointed at on its own.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { RAMP } from '../fidelity/heatmap.ts';
import { HISTOGRAM_EDGES } from '../fidelity/metric/score.ts';

import type { Gate3Run, ScoredSlide, ZoomColumn } from './gate3.ts';

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const STYLE = `
:root { color-scheme: dark; --bg:#0d1117; --panel:#161b22; --line:#30363d; --ink:#e6edf3;
        --dim:#8b949e; --ok:#3fb950; --bad:#f85149; }
* { box-sizing: border-box; }
body { margin:0; background:var(--bg); color:var(--ink); font:14px/1.55 ui-sans-serif,system-ui,sans-serif; }
main { max-width:1400px; margin:0 auto; padding:32px 24px 96px; }
h1 { font-size:26px; margin:0 0 4px; letter-spacing:-0.01em; }
h2 { font-size:18px; margin:48px 0 12px; }
h3 { font-size:15px; margin:28px 0 8px; }
p.sub { color:var(--dim); margin:0 0 28px; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th,td { text-align:left; padding:7px 10px; border-bottom:1px solid var(--line); vertical-align:top; }
th { color:var(--dim); font-weight:600; }
td.n, th.n { text-align:right; font-variant-numeric:tabular-nums; }
code { font:12px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace; color:var(--dim); }
.verdict { display:inline-block; padding:4px 12px; border-radius:999px; font-weight:600; font-size:13px; }
.pass { background:rgba(63,185,80,.15); color:var(--ok); }
.fail { background:rgba(248,81,73,.15); color:var(--bad); }
.note { background:rgba(139,148,158,.15); color:var(--dim); }
ul.gated { columns:2; gap:32px; padding-left:20px; }
ul.gated li { break-inside:avoid; }
.slide { border:1px solid var(--line); border-radius:10px; background:var(--panel); margin:20px 0; overflow:hidden; }
.slide > header { display:flex; flex-wrap:wrap; gap:12px; align-items:baseline;
                  padding:12px 16px; border-bottom:1px solid var(--line); }
.slide h4 { margin:0; font-size:15px; font-family:ui-monospace,monospace; }
.panes { display:grid; grid-template-columns:repeat(3,1fr); gap:1px; background:var(--line); }
.pane { background:var(--bg); padding:10px; min-width:0; }
.pane h5 { margin:0 0 8px; font-size:11px; text-transform:uppercase; letter-spacing:.08em; color:var(--dim); }
.pane img { display:block; width:100%; height:auto; background:#fff; border-radius:3px; }
.pane.heat img { background:var(--bg); }
.figures { display:flex; flex-wrap:wrap; gap:24px; padding:12px 16px; border-top:1px solid var(--line); }
.figure { font-variant-numeric:tabular-nums; }
.figure b { display:block; font-size:18px; font-weight:600; }
.figure span { font-size:11px; color:var(--dim); text-transform:uppercase; letter-spacing:.06em; }
.legend { display:flex; align-items:center; gap:2px; flex-wrap:wrap; margin:8px 0 0; }
.legend i { width:34px; height:12px; display:block; }
.legend small { color:var(--dim); font-size:11px; width:34px; text-align:center; display:block; }
details { margin:12px 0; }
summary { cursor:pointer; color:var(--dim); }
@media (max-width:900px) { .panes { grid-template-columns:1fr; } ul.gated { columns:1; } }
`;

function legend(): string {
  const edges = ['0', ...HISTOGRAM_EDGES.map((edge) => String(edge))];
  return (
    '<div class="legend">' +
    RAMP.map((colour) => `<i style="background:${colour}"></i>`).join('') +
    '</div><div class="legend">' +
    edges.map((edge) => `<small>${edge}</small>`).join('') +
    '</div>'
  );
}

function pct(zoom: number): string {
  return `${String(zoom * 100)} %`;
}

/** A column's name: its zoom, and the display when it is the 2x pass. */
export function columnLabel(column: ZoomColumn, ratio: number): string {
  return column.surface === 'ratio'
    ? `${pct(column.zoom)} on a ${String(ratio)}x display`
    : pct(column.zoom);
}

function gatedList(run: Gate3Run): string {
  const facts: [string, boolean][] = [
    [`${String(run.slides)} slides, of the hundred the gate asks for`, run.slides >= 100],
    [
      `${String(run.notDrawn.length)} not drawn, at any zoom or in the strip`,
      run.notDrawn.length === 0,
    ],
    [
      `${String(run.breaks.length)} zooms whose SVG differs from 100 % beyond stroke-width, stroke-dasharray and a line end's numbers`,
      run.breaks.length === 0,
    ],
    [
      `${String(run.ratio.breaks.length)} markup disagreements on a ${String(run.ratio.ratio)}x display, whose 100 % must be the 200 % markup and whose redrawn strip the 25 % markup`,
      run.ratio.breaks.length === 0,
    ],
    [`${String(run.pageErrors.length)} page or console errors`, run.pageErrors.length === 0],
    [
      `${String(run.requests.other.length)} requests the gate does not recognise`,
      run.requests.other.length === 0,
    ],
    [
      `${String(run.requests.afterOffline.length)} requests after the context went offline`,
      run.requests.afterOffline.length === 0,
    ],
    [
      run.baseline === null
        ? 'no baseline: a measurement records nothing'
        : run.baseline.recorded !== null
          ? `baseline recorded at ${run.baseline.recorded}`
          : `${String(run.baseline.changed.length)} rasters changed and ${String(run.baseline.vanished.length)} vanished against the recorded baseline`,
      run.baseline === null ||
        run.baseline.recorded !== null ||
        (run.baseline.changed.length === 0 && run.baseline.vanished.length === 0),
    ],
  ];
  return `<ul class="gated">${facts
    .map(
      ([text, ok]) =>
        `<li><span class="verdict ${ok ? 'pass' : 'fail'}">${ok ? 'ok' : 'no'}</span> ${escapeHtml(text)}</li>`,
    )
    .join('')}</ul>`;
}

function columnTable(run: Gate3Run): string {
  const rows = run.zooms
    .map(
      (column) =>
        `<tr><td>${escapeHtml(columnLabel(column, run.ratio.ratio))} (${column.surface})</td><td class="n">${String(column.width)}</td>` +
        `<td class="n">${String(column.cell)}</td><td class="n">${String(column.slides.length)}</td>` +
        `<td class="n">${run.mode === 'gate' ? String(column.meanBp) : '&mdash;'}</td>` +
        `<td class="n">${column.oracleSelfBp === null ? '&mdash;' : String(column.oracleSelfBp)}</td>` +
        `<td class="n">${mean(column.slides.map((s) => s.mountMs))}</td>` +
        `<td class="n">${mean(column.slides.map((s) => s.screenshotMs))}</td></tr>`,
    )
    .join('');
  return (
    '<table><thead><tr><th>zoom</th><th class="n">px wide</th><th class="n">cell</th><th class="n">drawn</th>' +
    '<th class="n">mean bp</th><th class="n">PowerPoint vs its 960</th><th class="n">mount ms</th><th class="n">screenshot ms</th></tr></thead>' +
    `<tbody>${rows}</tbody></table>`
  );
}

function mean(values: readonly number[]): string {
  if (values.length === 0) return '&mdash;';
  return (values.reduce((a, b) => a + b, 0) / values.length).toFixed(1);
}

function slideRows(run: Gate3Run): string {
  // Keyed by column, not zoom: the 2x column is at 100 % too.
  const byKey = new Map<string, Map<number, ScoredSlide>>();
  run.zooms.forEach((column, at) => {
    for (const slide of column.slides) {
      const row = byKey.get(slide.key) ?? new Map<number, ScoredSlide>();
      row.set(at, slide);
      byKey.set(slide.key, row);
    }
  });
  const worstOf = (row: Map<number, ScoredSlide>): number =>
    Math.min(...[...row.values()].map((slide) => slide.meanBp));
  const rows = [...byKey]
    .sort(([, a], [, b]) => worstOf(a) - worstOf(b))
    .map(
      ([key, row]) =>
        `<tr><td><code>${escapeHtml(key)}</code></td>` +
        run.zooms
          .map((_column, at) => {
            const slide = row.get(at);
            return `<td class="n">${slide === undefined ? '&mdash;' : `${String(slide.meanBp)} <code>${String(slide.maxD)}</code>`}</td>`;
          })
          .join('') +
        '</tr>',
    )
    .join('');
  return (
    `<table><thead><tr><th>slide</th>${run.zooms.map((c) => `<th class="n">${escapeHtml(columnLabel(c, run.ratio.ratio))}</th>`).join('')}</tr></thead>` +
    `<tbody>${rows}</tbody></table>`
  );
}

function worstCards(column: ZoomColumn, ratio: number, out: string): string {
  return column.worst
    .map((worst) => {
      const stem =
        `${worst.key}@${String(column.width)}` +
        (column.surface === 'ratio' ? `@${String(ratio)}x` : '');
      writeFileSync(join(out, `${stem}.ours.png`), worst.oursPng);
      writeFileSync(join(out, `${stem}.ppt.svg`), worst.theirsSvg);
      writeFileSync(join(out, `${stem}.diff.svg`), worst.heatSvg);
      const slide = column.slides.find((entry) => entry.key === worst.key);
      const regions =
        slide === undefined || slide.regions.length === 0
          ? '<code>nothing above the floor</code>'
          : `<code>${slide.regions
              .map(
                (region) =>
                  `${String(region.cells)} cells @ maxD ${String(region.maxD)} [${region.bbox.join(',')}]` +
                  (region.shapes.length === 0
                    ? ''
                    : ` &mdash; ${escapeHtml(region.shapes.join(', '))}`),
              )
              .join('<br>')}</code>`;
      return `<section class="slide">
  <header><h4>${escapeHtml(worst.key)} at ${escapeHtml(columnLabel(column, ratio))}</h4></header>
  <div class="panes">
    <div class="pane"><h5>PowerPoint &mdash; export at ${String(column.width)} px, as a grid</h5>
      <img src="${escapeHtml(stem)}.ppt.svg" alt="PowerPoint's rendering of ${escapeHtml(worst.key)}"></div>
    <div class="pane"><h5>Ours &mdash; the page's own screenshot</h5>
      <img src="${escapeHtml(stem)}.ours.png" alt="our rendering of ${escapeHtml(worst.key)}"></div>
    <div class="pane heat"><h5>Difference &mdash; ${String(column.cell)}px cells</h5>
      <img src="${escapeHtml(stem)}.diff.svg" alt="difference heatmap for ${escapeHtml(worst.key)}"></div>
  </div>
  <div class="figures">
    <div class="figure"><b>${String(worst.meanBp)}</b><span>mean bp</span></div>
    <div class="figure"><b>${slide === undefined ? '&mdash;' : String(slide.maxD)}</b><span>worst cell</span></div>
    <div class="figure" style="flex:1"><span>worst regions</span><div style="margin-top:4px">${regions}</div></div>
  </div>
</section>`;
    })
    .join('\n');
}

export function writeGate3Report(run: Gate3Run, out: string): string {
  mkdirSync(out, { recursive: true });

  const verdict =
    run.mode === 'measurement'
      ? '<span class="verdict note">a measurement, not a gate</span>'
      : run.ok
        ? '<span class="verdict pass">Gate 3 holds</span>'
        : '<span class="verdict fail">Gate 3 does not hold</span>';
  const problems = [
    ...run.notDrawn.map((gap) => `${gap.key} did not draw: ${gap.reason}`),
    ...run.breaks.map(
      (brk) => `${brk.key} differs at ${pct(brk.zoom)} beyond what the stroke rule owns`,
    ),
    ...run.ratio.breaks.map(
      (brk) => `${brk.key} on a ${String(run.ratio.ratio)}x display: ${brk.what} is not the zoom's`,
    ),
    ...run.pageErrors.map((error) => `page error: ${error}`),
    ...run.requests.other.map((url) => `request the gate does not recognise: ${url}`),
    ...run.requests.afterOffline.map((url) => `request after going offline: ${url}`),
    ...(run.baseline?.recorded === null
      ? run.baseline.changed.map((key) => `${key} rasterises differently than recorded`)
      : []),
    ...(run.baseline?.recorded === null
      ? run.baseline.vanished.map((key) => `${key} is recorded and was not drawn`)
      : []),
  ];
  const t = run.timings;

  const page = `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gate 3 &mdash; a hundred slides at any zoom</title>
<style>${STYLE}</style></head><body><main>
<h1>Gate 3 &mdash; a hundred slides, at any zoom, entirely client-side</h1>
<p class="sub">${verdict} &nbsp; <code>${escapeHtml(run.deckPath)}</code>, ${String(run.slides)} slide(s),
${String(run.size.cx)} &times; ${String(run.size.cy)} EMU, drawn by the studio page itself in Chromium and
screenshotted from its DOM; the deck was fetched once and the context went offline before the first zoom.</p>
${problems.length === 0 ? '' : `<p class="sub"><b>Open:</b> ${problems.map(escapeHtml).join('; ')}</p>`}
<h2>What is gated</h2>
${gatedList(run)}
<h2>Timings, reported and never asserted</h2>
<p class="sub">Read in ${t.readMs.toFixed(0)} ms, parsed in ${t.parseMs.toFixed(0)} ms, first slide painted at
${t.firstStageMs.toFixed(0)} ms, every thumbnail at ${t.stripMs.toFixed(0)} ms with ${t.stripCpuMs.toFixed(0)} ms on the
thread${t.heapBytes === null ? '' : `, ${(t.heapBytes / 1048576).toFixed(0)} MB of JS heap after the strip`}.
On a ${String(run.ratio.ratio)}x display, unannounced: the stage in ${run.ratio.stageMs.toFixed(0)} ms, and the strip the page redrew by itself in ${run.ratio.stripCpuMs.toFixed(0)} ms on its thread.
Requests: ${String(run.requests.total)} in all, ${String(run.requests.static)} for the page, ${String(run.requests.deck)} for the deck.</p>
${columnTable(run)}
${
  run.mode === 'gate'
    ? `<h2>Every slide, worst first</h2>
<p class="sub">Mean agreement in basis points against PowerPoint's own export at that width, and the worst cell in levels of 255.
Scores are reported and gate nothing (ADR 0035).</p>
${slideRows(run)}
<h2>The worst three at each zoom</h2>
<p class="sub">Left is PowerPoint, middle is the page, right is where they differ.</p>
${legend()}
${run.zooms.map((column) => `<h3>${escapeHtml(columnLabel(column, run.ratio.ratio))} &mdash; ${String(column.width)} px wide</h3>${worstCards(column, run.ratio.ratio, out)}`).join('\n')}`
    : ''
}
</main></body></html>
`;

  const file = join(out, 'index.html');
  writeFileSync(file, page);
  return file;
}

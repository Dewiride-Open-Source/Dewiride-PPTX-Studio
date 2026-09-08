/**
 * Gate 2's run, as a page a person looks at.
 *
 * Three panels a slide, and the panels are files rather than data URIs so that
 * each one can be opened, diffed and pointed at on its own. The middle panel is
 * the one to read the caption on: it is PowerPoint's PNG when a capture is at
 * hand, and otherwise the committed grid, which is coarse because a grid cell
 * is eight pixels and that is the resolution the score is computed at.
 */

import { copyFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { HISTOGRAM_EDGES } from '../fidelity/metric/score.ts';
import { RAMP } from '../fidelity/heatmap.ts';

import type { GateRun, SlideReport } from './gate2.ts';

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** The SVG our renderer emitted, with any prolog dropped so a file can hold it. */
function svgFile(markup: string): string {
  return markup.replace(/^\s*<\?xml[^?]*\?>\s*/, '');
}

const STYLE = `
:root { color-scheme: dark; --bg:#0d1117; --panel:#161b22; --line:#30363d; --ink:#e6edf3;
        --dim:#8b949e; --ok:#3fb950; --bad:#f85149; }
* { box-sizing: border-box; }
body { margin:0; background:var(--bg); color:var(--ink); font:14px/1.55 ui-sans-serif,system-ui,sans-serif; }
main { max-width:1400px; margin:0 auto; padding:32px 24px 96px; }
h1 { font-size:26px; margin:0 0 4px; letter-spacing:-0.01em; }
h2 { font-size:18px; margin:48px 0 12px; }
p.sub { color:var(--dim); margin:0 0 28px; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th,td { text-align:left; padding:7px 10px; border-bottom:1px solid var(--line); vertical-align:top; }
th { color:var(--dim); font-weight:600; }
td.n, th.n { text-align:right; font-variant-numeric:tabular-nums; }
code { font:12px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace; color:var(--dim); }
.verdict { display:inline-block; padding:4px 12px; border-radius:999px; font-weight:600; font-size:13px; }
.pass { background:rgba(63,185,80,.15); color:var(--ok); }
.fail { background:rgba(248,81,73,.15); color:var(--bad); }
.slide { border:1px solid var(--line); border-radius:10px; background:var(--panel); margin:20px 0; overflow:hidden; }
.slide > header { display:flex; flex-wrap:wrap; gap:12px; align-items:baseline;
                  padding:12px 16px; border-bottom:1px solid var(--line); }
.slide h3 { margin:0; font-size:15px; font-family:ui-monospace,monospace; }
.tags { display:flex; flex-wrap:wrap; gap:6px; }
.tag { font-size:11px; padding:2px 8px; border:1px solid var(--line); border-radius:999px; color:var(--dim); }
.panes { display:grid; grid-template-columns:repeat(3,1fr); gap:1px; background:var(--line); }
.pane { background:var(--bg); padding:10px; min-width:0; }
.pane h4 { margin:0 0 8px; font-size:11px; text-transform:uppercase; letter-spacing:.08em; color:var(--dim); }
.pane img { display:block; width:100%; height:auto; background:#fff; border-radius:3px; }
.pane.heat img { background:var(--bg); }
.figures { display:flex; flex-wrap:wrap; gap:24px; padding:12px 16px; border-top:1px solid var(--line); }
.figure { font-variant-numeric:tabular-nums; }
.figure b { display:block; font-size:18px; font-weight:600; }
.figure span { font-size:11px; color:var(--dim); text-transform:uppercase; letter-spacing:.06em; }
.legend { display:flex; align-items:center; gap:2px; flex-wrap:wrap; margin:8px 0 0; }
.legend i { width:34px; height:12px; display:block; }
.legend small { color:var(--dim); font-size:11px; width:34px; text-align:center; display:block; }
@media (max-width:900px) { .panes { grid-template-columns:1fr; } }
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

function coverageTable(run: GateRun): string {
  const rows = run.coverage
    .map((entry) => {
      const count = entry.slides.length;
      const state =
        count === 0
          ? '<span class="verdict fail">none</span>'
          : `<span class="verdict pass">${String(count)}</span>`;
      return (
        `<tr><td>${escapeHtml(entry.label)}</td><td class="n">${state}</td>` +
        `<td><code>${escapeHtml(entry.slides.join(' ')) || '&mdash;'}</code></td></tr>`
      );
    })
    .join('');
  return (
    '<table><thead><tr><th>Gate 2 names</th><th class="n">slides</th>' +
    `<th>which</th></tr></thead><tbody>${rows}</tbody></table>`
  );
}

function slideCard(slide: SlideReport, theirsSrc: string, theirsNote: string): string {
  const tags = slide.features
    .map((feature) => `<span class="tag">${escapeHtml(feature)}</span>`)
    .join('');
  const regions =
    slide.regions.length === 0
      ? '<td colspan="2"><code>nothing above the floor</code></td>'
      : `<td><code>${slide.regions
          .map(
            (region) =>
              `${String(region.cells)} cells @ maxD ${String(region.maxD)} ` +
              `[${region.bbox.join(',')}]` +
              (region.shapes.length === 0
                ? ''
                : ` &mdash; ${escapeHtml(region.shapes.join(', '))}`),
          )
          .join('<br>')}</code></td>`;
  return `<section class="slide">
  <header><h3>${escapeHtml(slide.key)}</h3><div class="tags">${tags}</div></header>
  <div class="panes">
    <div class="pane"><h4>PowerPoint &mdash; ${escapeHtml(theirsNote)}</h4>
      <img src="${escapeHtml(theirsSrc)}" alt="PowerPoint's rendering of ${escapeHtml(slide.key)}"></div>
    <div class="pane"><h4>Ours &mdash; render-svg</h4>
      <img src="${escapeHtml(slide.key)}.ours.svg" alt="our rendering of ${escapeHtml(slide.key)}"></div>
    <div class="pane heat"><h4>Difference &mdash; ${String(slide.cell)}px cells</h4>
      <img src="${escapeHtml(slide.key)}.diff.svg" alt="difference heatmap for ${escapeHtml(slide.key)}"></div>
  </div>
  <div class="figures">
    <div class="figure"><b>${String(slide.meanBp)}</b><span>mean bp</span></div>
    <div class="figure"><b>${String(slide.maxD)}</b><span>worst cell</span></div>
    <div class="figure"><b>${String(slide.hist.slice(1).reduce((sum, n) => sum + n, 0))}</b><span>cells differing</span></div>
    <div class="figure" style="flex:1"><span>worst regions</span>
      <table style="margin-top:4px"><tr>${regions}</tr></table></div>
  </div>
</section>`;
}

export function writeGate2Report(run: GateRun, out: string): string {
  mkdirSync(out, { recursive: true });

  const cards: string[] = [];
  for (const slide of run.slides) {
    writeFileSync(join(out, `${slide.key}.ours.svg`), svgFile(slide.oursSvg));
    writeFileSync(join(out, `${slide.key}.diff.svg`), slide.heatSvg);

    let src = `${slide.key}.ppt.svg`;
    let note = 'committed oracle grid';
    if (slide.theirsPng === null) {
      writeFileSync(join(out, src), slide.theirsSvg);
    } else {
      src = `${slide.key}.ppt.png`;
      note = 'exported PNG';
      copyFileSync(slide.theirsPng, join(out, src));
    }
    cards.push(slideCard(slide, src, note));
  }

  const uncovered = run.coverage.filter((entry) => entry.slides.length === 0);
  const verdict = run.ok
    ? '<span class="verdict pass">Gate 2 holds</span>'
    : '<span class="verdict fail">Gate 2 does not hold</span>';
  const problems = [
    ...uncovered.map((entry) => `no committed slide carries ${entry.label.toLowerCase()}`),
    ...run.notDrawn.map((gap) => `${gap.key} did not draw: ${gap.reason}`),
  ];

  const page = `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gate 2 &mdash; geometry and paint</title>
<style>${STYLE}</style></head><body><main>
<h1>Gate 2 &mdash; a real deck rendered silent</h1>
<p class="sub">${verdict} &nbsp; ${String(run.slides.length)} slide(s) drawn beside PowerPoint's own,
selected from ${String(run.scanned)} committed slide(s) by the nine features the gate names.
The reference is PowerPoint, not LibreOffice &mdash; see ADR 0038.</p>
${problems.length === 0 ? '' : `<p class="sub"><b>Open:</b> ${problems.map(escapeHtml).join('; ')}</p>`}
<h2>What Gate 2 names, and what carries it</h2>
${coverageTable(run)}
<h2>Slides</h2>
<p class="sub">Left is PowerPoint, middle is ours, right is where they differ. The heatmap is
transparent where the two agree and warmer the further apart they are; the scale is the same
dyadic one the histogram counts in.</p>
${legend()}
${cards.join('\n')}
</main></body></html>
`;

  const file = join(out, 'index.html');
  writeFileSync(file, page);
  return file;
}

'use client';

import { useEffect, useState } from 'react';

import { censusPackage } from '@pptx-studio/census';
import {
  AUTOFIT_LADDER,
  AUTONUMBER_SCHEMES,
  NATURAL_LINE_FACTOR,
  createCanvasMeasurer,
  createFontProbe,
  effectiveFontSize,
  fontReport,
  formatAutonumber,
  naturalLineHeight,
  wrapText,
  type FontFinding,
  type LineBox,
} from '@pptx-studio/text';

import { useDeck } from '@/deck/provider';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `import { createCanvasMeasurer, wrapText, fontReport } from '@pptx-studio/text';

const measurer = createCanvasMeasurer();
const font = { family: 'Calibri', sz: 1800 };   // sz is hundredths of a point

// We break the lines ourselves and place each one absolutely. That is what
// makes measure-equals-render true, and it is what autofit needs.
const lines = wrapText({
  text, widthPt: 320, hyphenWidthPt: 0,
  measure: (start, end) =>
    measurer.measure([...text].slice(start, end).join(''), font).width,
});

// And what this machine will actually draw each typeface in.
fontReport(['Aptos', 'Calibri', 'Wingdings']);`;

const SAMPLE =
  'PowerPoint measures a line box as a font-independent 1.2 times the font size, ' +
  'which ECMA-376 does not mention anywhere. Thirteen faces, seven sizes, one number.';

export default function TextPage() {
  const { deck } = useDeck();
  const [text, setText] = useState(SAMPLE);
  const [widthPt, setWidthPt] = useState(280);
  const [sizePt, setSizePt] = useState(18);
  const [family, setFamily] = useState('Calibri');
  const [findings, setFindings] = useState<readonly FontFinding[] | null>(null);
  const [start, setStart] = useState(1);

  // Measured in an effect rather than during render, because the server has no
  // canvas: breaking lines while rendering would give the first client pass a
  // different answer from the HTML it is hydrating, which React rejects.
  const [lines, setLines] = useState<readonly LineBox[]>([]);

  useEffect(() => {
    try {
      const measurer = createCanvasMeasurer();
      const font = { family, sz: Math.round(sizePt * 100) };
      const points = [...text];
      setLines(
        wrapText({
          text,
          widthPt,
          hyphenWidthPt: measurer.measure('-', font).width,
          measure: (from, to) => measurer.measure(points.slice(from, to).join(''), font).width,
        }),
      );
    } catch {
      // A browser that will not hand over a 2D context measures nothing, which
      // is a page with no preview rather than a page that fails.
      setLines([]);
    }
  }, [text, widthPt, sizePt, family]);

  // The deck's own typefaces, and what this browser would draw each one in.
  useEffect(() => {
    if (deck === null) return;
    try {
      const census = censusPackage(deck.bytes);
      const asked = census.typefaces.flatMap((one) =>
        Array.from({ length: one.count }, () => one.name),
      );
      setFindings(asked.length === 0 ? [] : fontReport(asked, createFontProbe()));
    } catch {
      setFindings(null);
    }
  }, [deck]);

  const points = [...text];
  const substituted = (findings ?? []).filter((one) => one.status !== 'present');

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Text</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/text</Mono> - measurement, line breaking, autofit, bullets and the font
          guard.
        </p>
      </header>

      <Panel
        title="Line breaking"
        hint="We break and place every line ourselves - that is what makes measure equal render."
      >
        <div className="flex flex-col gap-4 p-4">
          <textarea
            value={text}
            onChange={(event) => setText(event.target.value)}
            rows={3}
            className="w-full rounded border border-ink-600 bg-ink-800 px-3 py-2 text-[13px] text-ink-100"
          />
          <div className="flex flex-wrap items-center gap-4 text-[12px] text-ink-400">
            <label className="flex items-center gap-2">
              width
              <input
                type="range"
                min={80}
                max={520}
                value={widthPt}
                onChange={(event) => setWidthPt(Number(event.target.value))}
                className="w-40 accent-chrome"
              />
              <span className="tabular w-14 text-ink-200">{widthPt} pt</span>
            </label>
            <label className="flex items-center gap-2">
              size
              <input
                type="range"
                min={8}
                max={44}
                value={sizePt}
                onChange={(event) => setSizePt(Number(event.target.value))}
                className="w-28 accent-chrome"
              />
              <span className="tabular w-12 text-ink-200">{sizePt} pt</span>
            </label>
            <label className="flex items-center gap-2">
              face
              <select
                value={family}
                onChange={(event) => setFamily(event.target.value)}
                className="rounded border border-ink-600 bg-ink-800 px-2 py-1 font-mono text-[12px] text-ink-100"
              >
                {['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Tahoma', 'Verdana'].map(
                  (one) => (
                    <option key={one}>{one}</option>
                  ),
                )}
              </select>
            </label>
          </div>

          <div
            className="rounded border border-ink-600 bg-white p-3 text-black"
            style={{ width: `${String(widthPt)}pt`, maxWidth: '100%' }}
          >
            {lines.length === 0 ? (
              <p className="text-[13px] text-ink-500">No measuring surface here.</p>
            ) : (
              lines.map((line, at) => (
                <div
                  key={at}
                  style={{
                    fontFamily: family,
                    fontSize: `${String(sizePt)}pt`,
                    lineHeight: `${String(naturalLineHeight(sizePt))}pt`,
                    whiteSpace: 'pre',
                  }}
                >
                  {points.slice(line.start, line.end).join('')}
                </div>
              ))
            )}
          </div>

          <div className="grid grid-cols-3 gap-4">
            <Stat label="lines" value={lines.length} />
            <Stat label="line box" value={`${naturalLineHeight(sizePt).toFixed(1)} pt`} />
            <Stat label="factor" value={`${String(NATURAL_LINE_FACTOR)} x size`} />
          </div>
          <p className="text-[12px] text-ink-500">
            Never <Mono tone="dim">line-height: normal</Mono>: it resolves from{' '}
            <Mono tone="dim">hhea.lineGap</Mono> on macOS and{' '}
            <Mono tone="dim">OS/2.usWinAscent</Mono> on Windows, so the same HTML gives a different
            line count per operating system.
          </p>
        </div>
      </Panel>

      <div className="grid gap-4 lg:grid-cols-2">
        <Panel
          title="Fonts in this deck"
          hint="What it asks for, and what this browser will actually draw it in."
          aside={
            substituted.length === 0 ? (
              <Badge tone="good">all present</Badge>
            ) : (
              <Badge tone="warn">{substituted.length} substituted</Badge>
            )
          }
        >
          {findings === null ? (
            <p className="p-4 text-[13px] text-ink-400">Reading the deck…</p>
          ) : findings.length === 0 ? (
            <p className="p-4 text-[13px] text-ink-400">This deck names no typeface.</p>
          ) : (
            <Table head={['asked for', 'drawn as', 'metrics', 'runs']}>
              {findings.map((one) => (
                <Row key={one.typeface}>
                  <Cell>{one.typeface}</Cell>
                  <Cell>
                    {one.status === 'present' ? (
                      <Badge tone="good">as asked</Badge>
                    ) : one.status === 'substituted' ? (
                      <Badge tone="warn">{one.rendersAs}</Badge>
                    ) : (
                      <Badge tone="bad">missing</Badge>
                    )}
                  </Cell>
                  <Cell className="text-[12px] text-ink-400">
                    {one.metricCompatible === null
                      ? 'unknown'
                      : one.metricCompatible
                        ? 'compatible'
                        : 'approximate'}
                    {one.verified ? '' : ' (from the table)'}
                  </Cell>
                  <Cell className="tabular">{one.runs}</Cell>
                </Row>
              ))}
            </Table>
          )}
        </Panel>

        <Panel
          title="The autofit ladder"
          hint="Discrete rungs. A scale off the ladder is re-snapped by PowerPoint and the text visibly jumps."
        >
          <div className="max-h-80 overflow-y-auto p-4">
            <div className="grid grid-cols-[auto_auto_1fr] items-center gap-x-4 gap-y-1 text-[12px]">
              {AUTOFIT_LADDER.map((rung, at) => (
                <div key={at} className="contents">
                  <span className="tabular text-ink-200">
                    {(rung.fontScale / 1000).toFixed(0)}%
                  </span>
                  <span className="tabular text-ink-400">
                    {rung.lnSpcReduction === 0
                      ? '-'
                      : `-${(rung.lnSpcReduction / 1000).toFixed(0)}%`}
                  </span>
                  <span className="text-ink-500">
                    {sizePt}pt becomes{' '}
                    {(effectiveFontSize(sizePt * 100, rung.fontScale) / 100).toFixed(0)}pt
                  </span>
                </div>
              ))}
            </div>
          </div>
          <p className="border-t border-ink-700 px-4 py-3 text-[12px] text-ink-500">
            {AUTOFIT_LADDER.length} rungs. Font scale descending, and within one scale the smaller
            reduction first - the order is measured, not assumed.
          </p>
        </Panel>
      </div>

      <Panel
        title="Autonumber schemes"
        hint="Every ST_TextAutonumberScheme, at a start value you choose."
        aside={
          <label className="flex items-center gap-2 text-[12px] text-ink-400">
            start at
            <input
              type="number"
              min={1}
              max={999}
              value={start}
              onChange={(event) => setStart(Math.max(1, Number(event.target.value)))}
              className="w-16 rounded border border-ink-600 bg-ink-800 px-2 py-0.5 text-ink-100"
            />
          </label>
        }
      >
        <div className="grid grid-cols-2 gap-x-6 gap-y-1 p-4 sm:grid-cols-3 lg:grid-cols-4">
          {AUTONUMBER_SCHEMES.map((scheme) => {
            let shown: string;
            try {
              shown = formatAutonumber(scheme, start);
            } catch {
              shown = '-';
            }
            return (
              <div key={scheme} className="flex items-baseline justify-between gap-2 text-[12px]">
                <Mono tone="dim">{scheme}</Mono>
                <span className="text-ink-100">{shown}</span>
              </div>
            );
          })}
        </div>
      </Panel>

      <Panel title="How">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

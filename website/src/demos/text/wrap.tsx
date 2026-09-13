'use client';

import { useEffect, useMemo, useState } from 'react';

import {
  APPROXIMATE_FACE_METRICS,
  AUTOFIT_LADDER,
  NATURAL_LINE_FACTOR,
  breakOpportunities,
  createCanvasMeasurer,
  effectiveFontSize,
  faceLineMetrics,
  fitAutofit,
  hasFaceLineMetrics,
  naturalLineHeight,
  wrapText,
  type LineBox,
} from '@pptx-studio/text';

import { Badge } from '@/design/badge';
import { Mono } from '@/design/code';
import { Select, Slider, Switch } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';

const SAMPLE =
  'PowerPoint measures a line box as a font-independent 1.2 times the font size, ' +
  'which ECMA-376 does not mention anywhere. Thirteen faces, seven sizes, one number.';

const FACES = [
  'Calibri',
  'Arial',
  'Times New Roman',
  'Courier New',
  'Tahoma',
  'Verdana',
  'Georgia',
];

interface Wrapped {
  readonly lines: readonly LineBox[];
  readonly breaks: readonly number[];
  readonly fit: {
    fontScale: number;
    lnSpcReduction: number;
    heightPt: number;
    overflows: boolean;
  } | null;
}

/** Wrapping, break opportunities, and the autofit ladder against a box of chosen height. */
export function Wrap() {
  const [text, setText] = useState(SAMPLE);
  const [widthPt, setWidthPt] = useState(280);
  const [heightPt, setHeightPt] = useState(90);
  const [sizePt, setSizePt] = useState(18);
  const [family, setFamily] = useState('Calibri');
  const [showBreaks, setShowBreaks] = useState(true);
  const [wrapped, setWrapped] = useState<Wrapped | null>(null);
  const [noCanvas, setNoCanvas] = useState(false);

  // Measured after a frame: the canvas is a browser surface, and a build has none.
  useEffect(() => {
    const frame = requestAnimationFrame(() => {
      let measurer: ReturnType<typeof createCanvasMeasurer>;
      try {
        measurer = createCanvasMeasurer();
      } catch {
        setNoCanvas(true);
        return;
      }
      const points = [...text];
      const wrapAt = (pt: number): readonly LineBox[] => {
        const font = { family, sz: Math.round(pt * 100) };
        return wrapText({
          text,
          widthPt,
          hyphenWidthPt: measurer.measure('-', font).width,
          measure: (from, to) => measurer.measure(points.slice(from, to).join(''), font).width,
        });
      };
      const metrics = hasFaceLineMetrics(family)
        ? faceLineMetrics(family)
        : APPROXIMATE_FACE_METRICS;
      const fitted = fitAutofit({
        bodyHeightPt: heightPt,
        layout: (scale) => [
          {
            lines: wrapAt(effectiveFontSize(sizePt * 100, scale.fontScale) / 100).length,
            sz: sizePt * 100,
            metrics,
          },
        ],
      });
      setWrapped({
        lines: wrapAt(sizePt),
        breaks: breakOpportunities(text),
        fit: {
          fontScale: fitted.scale.fontScale,
          lnSpcReduction: fitted.scale.lnSpcReduction,
          heightPt: fitted.heightPt,
          overflows: fitted.overflows,
        },
      });
    });
    return () => cancelAnimationFrame(frame);
  }, [text, widthPt, heightPt, sizePt, family]);

  const points = useMemo(() => [...text], [text]);
  const breaks = useMemo(() => new Set(wrapped?.breaks ?? []), [wrapped]);

  return (
    <Panel
      title="Line breaking, and the autofit ladder"
      hint="Every line is broken and placed by the package - that is what makes measure equal render."
    >
      <div className="flex flex-col gap-4 p-4">
        <textarea
          value={text}
          onChange={(event) => setText(event.target.value)}
          rows={3}
          aria-label="Text to wrap"
          className="w-full rounded-control border border-line-strong bg-surface px-3 py-2 text-[13px] text-fg"
        />
        <div className="flex flex-wrap items-center gap-4 text-[12px] text-fg-muted">
          <span className="inline-flex items-center gap-2">
            width{' '}
            <Slider
              ariaLabel="Width in points"
              min={80}
              max={520}
              value={widthPt}
              onChange={setWidthPt}
              format={(v) => `${String(v)} pt`}
            />
          </span>
          <span className="inline-flex items-center gap-2">
            height{' '}
            <Slider
              ariaLabel="Height in points"
              min={30}
              max={300}
              value={heightPt}
              onChange={setHeightPt}
              format={(v) => `${String(v)} pt`}
            />
          </span>
          <span className="inline-flex items-center gap-2">
            size{' '}
            <Slider
              ariaLabel="Font size in points"
              min={8}
              max={44}
              value={sizePt}
              onChange={setSizePt}
              format={(v) => `${String(v)} pt`}
            />
          </span>
          <label className="inline-flex items-center gap-2">
            face
            <Select
              value={family}
              onChange={(event) => setFamily(event.target.value)}
              className="font-mono text-[12px]"
            >
              {FACES.map((one) => (
                <option key={one}>{one}</option>
              ))}
            </Select>
          </label>
          <Switch checked={showBreaks} onChange={setShowBreaks}>
            mark break opportunities
          </Switch>
        </div>

        {noCanvas ? (
          <p className="text-[13px] text-warn">
            No measuring surface here: this browser did not hand over a 2D context, so nothing can
            be wrapped.
          </p>
        ) : wrapped === null ? null : (
          <div className="flex flex-wrap items-start gap-6">
            <div
              className="relative rounded-[2px] border border-line-strong bg-white p-3 text-black"
              style={{ width: `${String(widthPt + 24)}pt`, maxWidth: '100%' }}
            >
              <div
                aria-hidden="true"
                className="pointer-events-none absolute inset-x-3 top-3 border-b border-dashed border-accent"
                style={{ height: `${String(heightPt)}pt` }}
              />
              {wrapped.lines.map((line, at) => (
                <div
                  key={at}
                  style={{
                    fontFamily: family,
                    fontSize: `${String(sizePt)}pt`,
                    lineHeight: `${String(naturalLineHeight(sizePt))}pt`,
                    whiteSpace: 'pre',
                  }}
                >
                  {points.slice(line.start, line.end).map((point, index) => {
                    const offset = line.start + index;
                    return (
                      <span
                        key={offset}
                        className={
                          showBreaks && breaks.has(offset) && index > 0
                            ? 'border-l border-accent/60'
                            : ''
                        }
                      >
                        {point}
                      </span>
                    );
                  })}
                </div>
              ))}
            </div>
            <div className="min-w-56 flex-1">
              <StatRow>
                <Stat label="lines" value={wrapped.lines.length} />
                <Stat label="line box" value={`${naturalLineHeight(sizePt).toFixed(1)} pt`} />
                <Stat label="factor" value={`${String(NATURAL_LINE_FACTOR)} × size`} />
                <Stat label="breaks" value={wrapped.breaks.length} />
              </StatRow>
              {wrapped.fit === null ? null : (
                <div className="mt-2 px-4 text-[13px]">
                  <p className="text-fg">
                    To fit {heightPt} pt the ladder picks{' '}
                    <Badge tone={wrapped.fit.overflows ? 'warn' : 'good'}>
                      {(wrapped.fit.fontScale / 1000).toFixed(0)}% font
                      {wrapped.fit.lnSpcReduction === 0
                        ? ''
                        : `, −${(wrapped.fit.lnSpcReduction / 1000).toFixed(0)}% spacing`}
                    </Badge>{' '}
                    - {(effectiveFontSize(sizePt * 100, wrapped.fit.fontScale) / 100).toFixed(0)}{' '}
                    pt, block {wrapped.fit.heightPt.toFixed(1)} pt
                    {wrapped.fit.overflows ? ', still overflowing at the last rung' : ''}.
                  </p>
                  <p className="mt-1 text-[12px] text-fg-faint">
                    {AUTOFIT_LADDER.length} rungs, font scale descending; a scale off the ladder is
                    re-snapped by PowerPoint and the text visibly jumps. The dashed line is the box.
                  </p>
                </div>
              )}
            </div>
          </div>
        )}
        <p className="text-[12px] text-fg-faint">
          Never <Mono tone="dim">line-height: normal</Mono>: it resolves from{' '}
          <Mono tone="dim">hhea.lineGap</Mono> on macOS and <Mono tone="dim">OS/2.usWinAscent</Mono>{' '}
          on Windows, so the same HTML gives a different line count per operating system.
        </p>
      </div>
    </Panel>
  );
}

'use client';

import { Callout } from '@/design/callout';
import { Code } from '@/design/code';
import { Panel } from '@/design/panel';
import type { DemoProps } from '../registry';
import { Bullets } from './bullets';
import { Fonts } from './fonts';
import { Wrap } from './wrap';

const HOW = `import { createCanvasMeasurer, wrapText, fitAutofit, fontReport, createFontProbe } from '@pptx-studio/text';

const measurer = createCanvasMeasurer();
const font = { family: 'Calibri', sz: 1800 };   // sz is hundredths of a point

// The package breaks the lines and places each one absolutely. That is what
// makes measure-equals-render true, and it is what autofit needs.
const lines = wrapText({
  text, widthPt: 320, hyphenWidthPt: 0,
  measure: (start, end) => measurer.measure([...text].slice(start, end).join(''), font).width,
});

// Edit mode: walk the ladder and take the first rung that fits.
fitAutofit({ bodyHeightPt: 90, layout: (scale) => paragraphsAt(scale) });

// And what this machine will actually draw each typeface in.
fontReport(['Aptos', 'Calibri', 'Wingdings'], createFontProbe());`;

export default function TextDemo({ full }: DemoProps) {
  return (
    <div className="flex flex-col gap-4">
      <Wrap />
      <Fonts />
      <Bullets />
      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Thai does not break: the line breaker is smaller than UAX #14 on purpose, and it says
              so. Date patterns are measured per locale, so a cell nobody measured falls back to the
              cached text and is labelled as such. A face nobody measured uses approximate metrics
              only when named.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

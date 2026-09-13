'use client';

import { Callout } from '@/design/callout';
import { Code } from '@/design/code';
import { Panel } from '@/design/panel';
import type { DemoProps } from '../registry';
import { Colours } from './colours';
import { Fills } from './fills';
import { Strokes } from './strokes';

const HOW = `import { applyTransforms, resolveColor, svgStops, dashArray, toCss } from '@pptx-studio/paint';

// Document order is load-bearing. On #4472C4:
//   lumMod 60% then lumOff 40%  ->  8FAADC   (PowerPoint's "Lighter 40%")
//   lumOff 40% then lumMod 60%  ->  517CC8   (wrong, and plausible)
const out = applyTransforms(base, [
  { op: 'lumMod', val: 60000 },   // hundred-thousandths, as the file writes it
  { op: 'lumOff', val: 40000 },
]);
toCss(out);                        // 'rgb(143 170 220)'

// A scheme colour needs the theme and the colour map the sheet resolves through.
resolveColor({ space: 'scheme', name: 'accent1', transforms: [] }, { scheme, map });

// What the renderer writes into <defs> for a gradient, and onto a stroke for a dash.
svgStops(gradientFill, { scheme, map });
dashArray(segments, widthPx, 'rnd');`;

export default function PaintDemo({ full }: DemoProps) {
  return (
    <div className="flex flex-col gap-4">
      <Colours />
      <Fills />
      <Strokes />
      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Every preview here is the SVG this library emits, drawn by your browser - not
              PowerPoint&apos;s own raster. Effects are on the slides pages; a preset shadow
              (prstShdw) is recorded on the way through and not modelled.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

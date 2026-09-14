'use client';

import { useState } from 'react';

import { Button } from '@/design/button';
import { Code } from '@/design/code';
import { Panel } from '@/design/panel';

const SHOWN = 60_000;

/** The string render-svg emitted for the slide on the stage, to read or to save. */
export function SvgSource({ svg, name }: { svg: string; name: string }) {
  const [shown, setShown] = useState(false);
  const download = () => {
    const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${name}.svg`;
    anchor.click();
    URL.revokeObjectURL(url);
  };
  return (
    <Panel
      title="The SVG render-svg emitted for this slide"
      hint={`${(svg.length / 1024).toFixed(1)} kB · ${String((svg.match(/<defs/g) ?? []).length)} <defs> · real <text>, never foreignObject`}
      aside={
        <div className="flex gap-2">
          <Button size="sm" onClick={() => setShown((on) => !on)} aria-expanded={shown}>
            {shown ? 'Hide' : 'View'}
          </Button>
          <Button size="sm" onClick={download}>
            Download .svg
          </Button>
        </div>
      }
    >
      {shown ? (
        <div className="max-h-96 overflow-auto p-4">
          <Code
            lang="xml"
            code={
              svg.length > SHOWN
                ? `${svg.slice(0, SHOWN)}\n<!-- … ${String(svg.length - SHOWN)} more characters -->`
                : svg
            }
          />
        </div>
      ) : null}
    </Panel>
  );
}

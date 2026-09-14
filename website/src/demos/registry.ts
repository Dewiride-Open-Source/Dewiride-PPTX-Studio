import type { PackageName } from '@/site/packages';

export interface DemoProps {
  /** True on its own page, false when embedded in a docs page. */
  readonly full: boolean;
}

export interface DemoEntry {
  readonly name: PackageName;
  readonly title: string;
  readonly blurb: string;
  /** True where the page shows what the build machine rendered, not the tab. */
  readonly prerendered?: true;
}

export const DEMOS: readonly DemoEntry[] = [
  {
    name: 'render-dom',
    title: 'Viewer',
    blurb: 'Every slide, live: a strip, zoom, a selectable text layer, and the debug overlay.',
  },
  {
    name: 'render-svg',
    title: 'Editor',
    blurb:
      'Click a shape: drag it, colour it, type in it, delete it. Undo. Download a file PowerPoint opens.',
  },
  {
    name: 'model',
    title: 'Inheritance',
    blurb: 'Where every resolved property came from, and what undefined means.',
  },
  {
    name: 'opc',
    title: 'Package',
    blurb: 'Parts, content types and the relationship graph, with limits and digests.',
  },
  {
    name: 'xml',
    title: 'Markup',
    blurb:
      'The tree with its byte offsets, an edit and its exact inverse, byte-identical re-emission.',
  },
  {
    name: 'census',
    title: 'Census',
    blurb: 'What is inside, counted in a Web Worker.',
  },
  {
    name: 'validate',
    title: 'Validate',
    blurb: 'The 29 rules that keep PowerPoint from repairing a file.',
  },
  {
    name: 'writer',
    title: 'Export',
    blurb: 'Export in a Worker, prove nothing untouched moved, and download the file.',
  },
  {
    name: 'geometry',
    title: 'Geometry',
    blurb: 'The 187 presets, their guides, and adjust handles you can drag or key.',
  },
  {
    name: 'paint',
    title: 'Paint',
    blurb: 'Colour transforms, gradients, the 54 pattern tiles, dashes and arrowheads.',
  },
  {
    name: 'text',
    title: 'Text',
    blurb: 'Measurement, line breaking, the autofit ladder, bullets, fields and fonts.',
  },
  {
    name: 'cli',
    title: 'Command line',
    blurb: 'Slides rendered in Node when this site was built, with no LibreOffice and no browser.',
    prerendered: true,
  },
];

export function demoNamed(name: string): DemoEntry | undefined {
  return DEMOS.find((demo) => demo.name === name);
}

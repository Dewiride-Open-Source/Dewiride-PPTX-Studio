import type { PackageName } from '@/site/packages';

export interface DemoProps {
  /** True on its own page, false when embedded in a docs page. */
  readonly full: boolean;
}

export interface DemoEntry {
  readonly name: PackageName;
  readonly title: string;
  /** What the page shows, in one plain sentence. */
  readonly blurb: string;
  /** The first thing to do on it. */
  readonly tryThis: string;
  /** True where the page shows what the build machine rendered, not the tab. */
  readonly prerendered?: true;
}

export const DEMOS: readonly DemoEntry[] = [
  {
    name: 'render-dom',
    title: 'Viewer',
    blurb: 'Every slide of the deck drawn live in the page, with a thumbnail strip and zoom.',
    tryThis: 'Click a thumbnail, zoom in, and select some text - it stays real text at every size.',
  },
  {
    name: 'render-svg',
    title: 'Editor',
    blurb:
      'Click a shape and drag it, colour it, type in it or delete it; undo; download the file.',
    tryThis: 'Click the title, drag it somewhere else, then double-click it and type.',
  },
  {
    name: 'model',
    title: 'Inheritance',
    blurb:
      'Why a shape looks the way it does: which of the slide, its layout, the master or the theme decided each property.',
    tryThis:
      'Pick a shape and read down the chain; a property marked inherited is one the slide never wrote.',
  },
  {
    name: 'opc',
    title: 'Package',
    blurb:
      'A .pptx is a ZIP of XML parts that point at each other; this lists them and draws the links.',
    tryThis: 'Pick a part to see its first bytes, its content type and what it links to.',
  },
  {
    name: 'xml',
    title: 'Markup',
    blurb:
      'The XML of one part, edited in place and written back without touching a byte you did not change.',
    tryThis:
      'Pick a slide part, change one attribute, and read the exact bytes that differ - then undo it.',
  },
  {
    name: 'census',
    title: 'Census',
    blurb:
      'What is inside the deck - slides, shapes, charts, fonts, media - counted off the main thread.',
    tryThis: 'Open a bigger deck and watch the counts arrive while the page stays responsive.',
  },
  {
    name: 'validate',
    title: 'Validate',
    blurb: 'The 29 checks that keep PowerPoint from asking to repair a file, run on this deck.',
    tryThis:
      'Open a rule to read its evidence, then drop a deck of your own and see what it finds.',
  },
  {
    name: 'writer',
    title: 'Export',
    blurb:
      'The deck saved back out: only edited parts are rewritten, everything else streams through as it came in.',
    tryThis:
      'Press Export, read which parts were rewritten and which streamed through untouched, and download it.',
  },
  {
    name: 'geometry',
    title: 'Geometry',
    blurb: 'The 187 preset shapes PowerPoint knows, and the handles that reshape them.',
    tryThis: 'Pick a preset and drag a handle; the value a save would write updates as you drag.',
  },
  {
    name: 'paint',
    title: 'Paint',
    blurb:
      'How colours, gradients, patterns, dashes and arrowheads are computed, each measured against PowerPoint.',
    tryThis: 'Move a gradient stop or a colour transform and watch the output follow.',
  },
  {
    name: 'text',
    title: 'Text',
    blurb: 'How text is measured, wrapped and shrunk to fit, the way PowerPoint does it.',
    tryThis: 'Type into the box and watch the line breaks move and the autofit step change.',
  },
  {
    name: 'cli',
    title: 'Command line',
    blurb: 'Slides rendered in Node when this site was built, with no LibreOffice and no browser.',
    tryThis: 'Copy the command shown and run it on a deck of your own.',
    prerendered: true,
  },
];

export function demoNamed(name: string): DemoEntry | undefined {
  return DEMOS.find((demo) => demo.name === name);
}

/** The eleven published packages, and the tool that shows each one off. */

export interface Tool {
  readonly href: string;
  readonly name: string;
  /** The package the tool is about, without the `@pptx-studio/` scope. */
  readonly package: string;
  readonly blurb: string;
  /** True where the work happens on the server rather than in the tab. */
  readonly server?: boolean;
}

export const TOOLS: readonly Tool[] = [
  {
    href: '/slides',
    name: 'Slides',
    package: 'render-svg',
    blurb: 'Draw the deck, select a shape, move it, recolour it, retype it.',
  },
  {
    href: '/inheritance',
    name: 'Inheritance',
    package: 'model',
    blurb: 'Where every resolved property actually came from.',
  },
  {
    href: '/package',
    name: 'Package',
    package: 'opc',
    blurb: 'Parts, content types and the relationship graph.',
  },
  {
    href: '/markup',
    name: 'Markup',
    package: 'xml',
    blurb: 'The tree, its byte offsets, and byte-identical re-emission.',
  },
  {
    href: '/census',
    name: 'Census',
    package: 'census',
    blurb: 'What is inside, counted in a Web Worker.',
  },
  {
    href: '/validate',
    name: 'Validate',
    package: 'validate',
    blurb: 'The twenty-nine rules that keep PowerPoint from repairing a file.',
  },
  {
    href: '/roundtrip',
    name: 'Round trip',
    package: 'writer',
    blurb: 'Export, and prove nothing that was not edited moved.',
  },
  {
    href: '/geometry',
    name: 'Geometry',
    package: 'geometry',
    blurb: 'The 187 presets, their guides and their adjust handles.',
  },
  {
    href: '/paint',
    name: 'Paint',
    package: 'paint',
    blurb: 'Colour transforms, gradients, patterns, dashes.',
  },
  {
    href: '/text',
    name: 'Text',
    package: 'text',
    blurb: 'Measurement, line breaking, the autofit ladder, bullets.',
  },
  {
    href: '/thumbnails',
    name: 'Thumbnails',
    package: 'cli',
    blurb: 'Server-side rendering with no LibreOffice and no browser.',
    server: true,
  },
];

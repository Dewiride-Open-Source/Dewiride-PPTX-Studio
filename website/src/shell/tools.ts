/** The eleven published packages, and the tool that shows each one off. */

export interface Tool {
  readonly href: string;
  readonly name: string;
  /** The package the tool is about, without the `@pptx-studio/` scope. */
  readonly package: string;
  readonly blurb: string;
  /** True where the page shows what the build machine rendered, not the tab. */
  readonly prerendered?: boolean;
}

export const TOOLS: readonly Tool[] = [
  {
    href: '/demos/render-svg',
    name: 'Slides',
    package: 'render-svg',
    blurb: 'Draw the deck, select a shape, move it, recolour it, retype it.',
  },
  {
    href: '/demos/model',
    name: 'Inheritance',
    package: 'model',
    blurb: 'Where every resolved property actually came from.',
  },
  {
    href: '/demos/opc',
    name: 'Package',
    package: 'opc',
    blurb: 'Parts, content types and the relationship graph.',
  },
  {
    href: '/demos/xml',
    name: 'Markup',
    package: 'xml',
    blurb: 'The tree, its byte offsets, and byte-identical re-emission.',
  },
  {
    href: '/demos/census',
    name: 'Census',
    package: 'census',
    blurb: 'What is inside, counted in a Web Worker.',
  },
  {
    href: '/demos/validate',
    name: 'Validate',
    package: 'validate',
    blurb: 'The twenty-nine rules that keep PowerPoint from repairing a file.',
  },
  {
    href: '/demos/writer',
    name: 'Round trip',
    package: 'writer',
    blurb: 'Export, and prove nothing that was not edited moved.',
  },
  {
    href: '/demos/geometry',
    name: 'Geometry',
    package: 'geometry',
    blurb: 'The 187 presets, their guides and their adjust handles.',
  },
  {
    href: '/demos/paint',
    name: 'Paint',
    package: 'paint',
    blurb: 'Colour transforms, gradients, patterns, dashes.',
  },
  {
    href: '/demos/text',
    name: 'Text',
    package: 'text',
    blurb: 'Measurement, line breaking, the autofit ladder, bullets.',
  },
  {
    href: '/demos/cli',
    name: 'Thumbnails',
    package: 'cli',
    blurb: 'Rendered in Node at build time, with no LibreOffice and no browser.',
    prerendered: true,
  },
];

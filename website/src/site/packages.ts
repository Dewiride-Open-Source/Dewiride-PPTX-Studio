export type PackageName =
  | 'opc'
  | 'xml'
  | 'census'
  | 'geometry'
  | 'paint'
  | 'text'
  | 'model'
  | 'render-svg'
  | 'render-dom'
  | 'validate'
  | 'writer'
  | 'cli';

export type Runtime = 'tab' | 'worker' | 'node';

/** The reading-order groups the site uses; the true layering is drawn on the packages page. */
export type PackageGroup = 'read' | 'draw' | 'write' | 'command-line';

export interface PackageFacts {
  readonly name: PackageName;
  readonly group: PackageGroup;
  readonly runsIn: readonly Runtime[];
  /** One sentence, for a developer who has never heard of the project. */
  readonly blurb: string;
  /** The packages the 30-second example imports beside this one. */
  readonly dependsOn: readonly PackageName[];
}

export const GROUP_TITLES: Record<PackageGroup, string> = {
  read: 'Read the file',
  draw: 'Draw it',
  write: 'Write it back',
  'command-line': 'From the command line',
};

export const PACKAGES: readonly PackageFacts[] = [
  {
    name: 'opc',
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Reads and writes the .pptx ZIP - parts, content types and relationships - with hard limits against hostile files.',
    dependsOn: [],
  },
  {
    name: 'xml',
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Parses OOXML into a tree that remembers its own bytes, so an edit changes only what you touched and undo is exact.',
    dependsOn: [],
  },
  {
    name: 'census',
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Tells you what is inside a deck as plain JSON - parts, features, namespaces, problems - without refusing a broken file.',
    dependsOn: ['opc', 'xml'],
  },
  {
    name: 'model',
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Slides, layouts, masters and themes as one document, with a resolver that reports where every value came from.',
    dependsOn: ['opc', 'xml', 'geometry', 'paint'],
  },
  {
    name: 'geometry',
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'The 187 preset shapes and custom paths, turned into SVG path data; drag an adjust handle and get the number to write back.',
    dependsOn: [],
  },
  {
    name: 'paint',
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Colours, gradients, pattern tiles, strokes and effects, each rule measured against PowerPoint rather than read from the standard.',
    dependsOn: [],
  },
  {
    name: 'text',
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Line spacing, wrapping, autofit, bullets, fields and font substitution, the way PowerPoint actually lays them out.',
    dependsOn: [],
  },
  {
    name: 'render-svg',
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb: 'A slide to an SVG string - in a tab, in a Worker, or in Node.',
    dependsOn: ['opc', 'model'],
  },
  {
    name: 'render-dom',
    group: 'draw',
    runsIn: ['tab'],
    blurb:
      'The same slide as live SVG elements you can hit-test, zoom and overlay, plus an HTML text layer for selection and screen readers.',
    dependsOn: ['opc', 'model', 'render-svg'],
  },
  {
    name: 'validate',
    group: 'write',
    runsIn: ['tab', 'worker', 'node'],
    blurb: 'The 31 rules a .pptx must not break, with the part and XPath of everything that fired.',
    dependsOn: ['opc'],
  },
  {
    name: 'writer',
    group: 'write',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Exports the deck, re-serialising only what changed, checking preservation, and refusing to return a file it broke.',
    dependsOn: ['opc', 'xml'],
  },
  {
    name: 'cli',
    group: 'command-line',
    runsIn: ['node'],
    blurb:
      'The same engine from the command line: render slides without a browser, inspect, validate, round-trip, bisect.',
    dependsOn: [],
  },
];

export function packageFacts(name: PackageName): PackageFacts {
  const found = PACKAGES.find((one) => one.name === name);
  if (found === undefined) throw new Error(`no such package: ${name}`);
  return found;
}

export const RUNTIME_LABELS: Record<Runtime, string> = {
  tab: 'Tab',
  worker: 'Worker',
  node: 'Node',
};

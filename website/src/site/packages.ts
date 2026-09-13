import census from '@pptx-studio/census/package.json';
import cli from '@pptx-studio/cli/package.json';
import geometry from '@pptx-studio/geometry/package.json';
import model from '@pptx-studio/model/package.json';
import opc from '@pptx-studio/opc/package.json';
import paint from '@pptx-studio/paint/package.json';
import renderDom from '@pptx-studio/render-dom/package.json';
import renderSvg from '@pptx-studio/render-svg/package.json';
import text from '@pptx-studio/text/package.json';
import validate from '@pptx-studio/validate/package.json';
import writer from '@pptx-studio/writer/package.json';
import xml from '@pptx-studio/xml/package.json';

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
  /** The version installed from npm when this site was built. */
  readonly version: string;
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
    version: opc.version,
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Reads and writes the .pptx ZIP - parts, content types and relationships - with hard limits against hostile files.',
    dependsOn: [],
  },
  {
    name: 'xml',
    version: xml.version,
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Parses OOXML into a tree that remembers its own bytes, so an edit changes only what you touched and undo is exact.',
    dependsOn: [],
  },
  {
    name: 'census',
    version: census.version,
    group: 'read',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Tells you what is inside a deck as plain JSON - parts, features, namespaces, problems - without refusing a broken file.',
    dependsOn: ['opc', 'xml'],
  },
  {
    name: 'model',
    version: model.version,
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Slides, layouts, masters and themes as one document, with a resolver that reports where every value came from.',
    dependsOn: ['opc', 'xml', 'geometry', 'paint'],
  },
  {
    name: 'geometry',
    version: geometry.version,
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'The 187 preset shapes and custom paths, turned into SVG path data; drag an adjust handle and get the number to write back.',
    dependsOn: [],
  },
  {
    name: 'paint',
    version: paint.version,
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Colours, gradients, pattern tiles, strokes and effects, each rule measured against PowerPoint rather than read from the standard.',
    dependsOn: [],
  },
  {
    name: 'text',
    version: text.version,
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Line spacing, wrapping, autofit, bullets, fields and font substitution, the way PowerPoint actually lays them out.',
    dependsOn: [],
  },
  {
    name: 'render-svg',
    version: renderSvg.version,
    group: 'draw',
    runsIn: ['tab', 'worker', 'node'],
    blurb: 'A slide to an SVG string - in a tab, in a Worker, or in Node.',
    dependsOn: ['opc', 'model'],
  },
  {
    name: 'render-dom',
    version: renderDom.version,
    group: 'draw',
    runsIn: ['tab'],
    blurb:
      'The same slide as live SVG elements you can hit-test, zoom and overlay, plus an HTML text layer for selection and screen readers.',
    dependsOn: ['opc', 'model', 'render-svg'],
  },
  {
    name: 'validate',
    version: validate.version,
    group: 'write',
    runsIn: ['tab', 'worker', 'node'],
    blurb: 'The 29 rules a .pptx must not break, with the part and XPath of everything that fired.',
    dependsOn: ['opc'],
  },
  {
    name: 'writer',
    version: writer.version,
    group: 'write',
    runsIn: ['tab', 'worker', 'node'],
    blurb:
      'Exports the deck, re-serialising only what changed, checking preservation, and refusing to return a file it broke.',
    dependsOn: ['opc', 'xml'],
  },
  {
    name: 'cli',
    version: cli.version,
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

import { readFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { isOpcError } from '@pptx-studio/opc';
import { INSPECT_DEFAULTS, runInspect, type InspectOptions } from './inspect.js';

/**
 * The command line.
 *
 * `node:util.parseArgs` rather than a dependency. A CLI whose entire job in
 * sub-phase 0.8 is one verb with five flags does not need an argument library,
 * and the ones on offer would each be larger than everything this package
 * contains.
 *
 * The verbs the plan gives this package - `roundtrip`, `render`, `validate`,
 * `fidelity`, `resolve`, `bisect` - are listed in the help text with the phase
 * that brings them, and asking for one says so rather than "unknown command".
 * A tool that knows what it will be able to do is more useful than one that
 * pretends the request was nonsense.
 */

interface PlannedVerb {
  readonly name: string;
  readonly summary: string;
  readonly phase: string;
}

const PLANNED: readonly PlannedVerb[] = [
  {
    name: 'roundtrip',
    summary: 'read a deck and write it back, then prove nothing moved',
    phase: '1.4',
  },
  {
    name: 'bisect',
    summary: 'find the smallest change that makes PowerPoint repair a deck',
    phase: '1.5',
  },
  { name: 'validate', summary: 'run the must-not-break rules over a package', phase: '1.2' },
  { name: 'render', summary: 'render slides to SVG or PNG without a browser', phase: '3.10' },
  { name: 'resolve', summary: 'show where a resolved property came from', phase: '7.x' },
  { name: 'fidelity', summary: 'score a render against a reference', phase: '3.9' },
];

function version(): string {
  // Resolves the same from `dist/cli.js` and from `src/main.ts`, so a run
  // straight off the source reports the version the package will publish.
  const manifest: unknown = JSON.parse(
    readFileSync(new URL('../package.json', import.meta.url), 'utf8'),
  );
  return typeof manifest === 'object' && manifest !== null && 'version' in manifest
    ? String(manifest.version)
    : '0.0.0';
}

function usage(): string {
  const planned = PLANNED.map(
    (verb) => '  ' + verb.name.padEnd(12) + verb.summary + '  (sub-phase ' + verb.phase + ')',
  ).join('\n');
  return [
    'pptx-studio ' + version(),
    '',
    'Usage: pptx-studio inspect <deck.pptx> [options]',
    '',
    'Options:',
    '  --json           the census as JSON, for a script or for committing as a fixture',
    '  --parts          include the per-part table',
    '  --namespaces     include the namespace histogram',
    '  --top <n>        rows per histogram before truncating (default 15)',
    '  --out <file>     write to a file instead of stdout',
    '  -h, --help       this text',
    '  -v, --version    print the version',
    '',
    'Exit status is 1 when the census found a structural error, 0 otherwise.',
    'Warnings and notes never fail the command.',
    '',
    'Not built yet:',
    planned,
    '',
  ].join('\n');
}

export interface Streams {
  readonly out: (text: string) => void;
  readonly err: (text: string) => void;
}

const CONSOLE_STREAMS: Streams = {
  out: (text) => process.stdout.write(text),
  err: (text) => process.stderr.write(text),
};

/** Parse, dispatch, and return an exit code. Never calls `process.exit`. */
export function main(argv: readonly string[], streams: Streams = CONSOLE_STREAMS): number {
  let parsed;
  try {
    parsed = parseArgs({
      args: [...argv],
      allowPositionals: true,
      options: {
        json: { type: 'boolean', default: false },
        parts: { type: 'boolean', default: false },
        namespaces: { type: 'boolean', default: false },
        top: { type: 'string' },
        out: { type: 'string' },
        help: { type: 'boolean', short: 'h', default: false },
        version: { type: 'boolean', short: 'v', default: false },
      },
    });
  } catch (error) {
    streams.err(describe(error) + '\n\n' + usage());
    return 2;
  }

  const { values, positionals } = parsed;
  if (values.version === true) {
    streams.out(version() + '\n');
    return 0;
  }
  const command = positionals[0];
  if (values.help === true || command === undefined) {
    streams.out(usage());
    return command === undefined && values.help !== true ? 2 : 0;
  }

  const planned = PLANNED.find((verb) => verb.name === command);
  if (planned !== undefined) {
    streams.err(
      'pptx-studio ' +
        command +
        ' is not built yet: it arrives with sub-phase ' +
        planned.phase +
        '.\n' +
        '  ' +
        planned.summary +
        '\n',
    );
    return 2;
  }

  if (command !== 'inspect') {
    streams.err('unknown command: ' + command + '\n\n' + usage());
    return 2;
  }

  const file = positionals[1];
  if (file === undefined) {
    streams.err('inspect needs a path to a .pptx\n\n' + usage());
    return 2;
  }

  const top = values.top === undefined ? INSPECT_DEFAULTS.top : Number.parseInt(values.top, 10);
  if (!Number.isFinite(top) || top < 1) {
    streams.err('--top wants a positive integer, got ' + String(values.top) + '\n');
    return 2;
  }

  const options: InspectOptions = {
    json: values.json === true,
    parts: values.parts === true,
    namespaces: values.namespaces === true,
    top,
    out: values.out ?? null,
  };

  try {
    return runInspect(file, options, streams.out);
  } catch (error) {
    // A census does not refuse a broken deck; it refuses a file that is not a
    // package at all, and `opc` says which rule it broke. Reporting the code
    // matters more than the sentence - it is what a script can branch on.
    if (isOpcError(error)) {
      streams.err('pptx-studio: ' + error.code + ': ' + error.message + '\n');
      return 1;
    }
    streams.err('pptx-studio: ' + describe(error) + '\n');
    return 1;
  }
}

function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

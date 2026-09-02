import { readFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { isOpcError } from '@pptx-studio/opc';
import { isValidateError } from '@pptx-studio/validate';
import { BISECT_DEFAULTS, runBisect, type BisectOptions } from './bisect.js';
import { INSPECT_DEFAULTS, runInspect, type InspectOptions } from './inspect.js';
import { runRoundTrip, type RoundTripOptions } from './roundtrip.js';
import { runValidate, type ValidateOptions } from './validate.js';

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
    'Usage: pptx-studio <command> <deck.pptx> [options]',
    '',
    'Commands:',
    '  inspect          what is inside a package: parts, relationships, features',
    '  validate         the must-not-break rules, with the part and the XPath',
    '  roundtrip        read a deck, write it back, and prove nothing moved',
    '  bisect           narrow a broken deck to the change that breaks it',
    '',
    'inspect options:',
    '  --json           the census as JSON, for a script or for committing as a fixture',
    '  --parts          include the per-part table',
    '  --namespaces     include the namespace histogram',
    '  --top <n>        rows per histogram before truncating (default 15)',
    '  --out <file>     write to a file instead of stdout',
    '',
    'validate options:',
    '  --json           the report as JSON',
    '  --explain        append the rationale for every rule that fired',
    '  --quiet          fatal findings only',
    '  --out <file>     write to a file instead of stdout',
    '',
    'roundtrip options:',
    '  --json           the comparison as JSON, with a digest per part',
    '  --quiet          the differences only, without the tally',
    '  --write <file>   also save the package that was written, to open in PowerPoint',
    '  --out <file>     write the report to a file instead of stdout',
    '',
    'bisect <deck.pptx> [broken.pptx]',
    '  With one deck the broken package is our own export of it; with two they',
    '  are the original and the broken package, in that order.',
    '',
    '  --oracle <name>  validate (default), powerpoint, or command',
    '  --command <cmd>  for --oracle command; {} becomes the candidate path',
    '  --max-runs <n>   ceiling on oracle runs (default 2000)',
    '  --timeout <ms>   per run, for the oracles that spawn something',
    '  --progress       a line per oracle run; a bisection is not quick',
    '  --write <file>   save the smallest package that still fails',
    '  --json           the result as JSON',
    '  --quiet          the changes that matter, without the tally',
    '  --out <file>     write the report to a file instead of stdout',
    '',
    '  -h, --help       this text',
    '  -v, --version    print the version',
    '',
    'Exit status is 1 when inspect finds a structural error, when validate finds',
    'anything fatal, when roundtrip finds a difference, or when bisect localizes',
    'one. 0 otherwise. Warnings and notes never fail a command.',
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
        explain: { type: 'boolean', default: false },
        quiet: { type: 'boolean', default: false },
        parts: { type: 'boolean', default: false },
        namespaces: { type: 'boolean', default: false },
        top: { type: 'string' },
        out: { type: 'string' },
        write: { type: 'string' },
        oracle: { type: 'string' },
        command: { type: 'string' },
        'max-runs': { type: 'string' },
        timeout: { type: 'string' },
        progress: { type: 'boolean', default: false },
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

  if (
    command !== 'inspect' &&
    command !== 'validate' &&
    command !== 'roundtrip' &&
    command !== 'bisect'
  ) {
    streams.err('unknown command: ' + command + '\n\n' + usage());
    return 2;
  }

  const file = positionals[1];
  if (file === undefined) {
    streams.err(command + ' needs a path to a .pptx\n\n' + usage());
    return 2;
  }

  if (command === 'bisect') {
    const oracle = values.oracle ?? BISECT_DEFAULTS.oracle;
    if (oracle !== 'validate' && oracle !== 'powerpoint' && oracle !== 'command') {
      streams.err('--oracle wants validate, powerpoint or command, got ' + oracle + '\n');
      return 2;
    }
    if (oracle === 'command' && values.command === undefined) {
      streams.err('--oracle command needs --command "<what to run>"\n');
      return 2;
    }
    const maxRuns = positiveInteger(values['max-runs'], BISECT_DEFAULTS.maxRuns);
    const timeout = positiveInteger(values.timeout, BISECT_DEFAULTS.timeout);
    if (maxRuns === null || timeout === null) {
      streams.err('--max-runs and --timeout want positive integers\n');
      return 2;
    }

    const options: BisectOptions = {
      oracle,
      command: values.command ?? null,
      maxRuns,
      timeout,
      write: values.write ?? null,
      json: values.json === true,
      quiet: values.quiet === true,
      out: values.out ?? null,
      progress: values.progress === true,
    };
    try {
      return runBisect(file, positionals[2] ?? null, options, streams.out);
    } catch (error) {
      // The same three-way split `roundtrip` makes, plus one of its own: the
      // one-deck form exports the deck first, so the firewall can refuse before
      // there is anything to bisect.
      if (isValidateError(error)) {
        streams.err(
          'pptx-studio: the export was refused, so there was no broken package to compare.\n  ' +
            error.message +
            '\n',
        );
      } else if (isOpcError(error)) {
        streams.err('pptx-studio: ' + error.code + ': ' + error.message + '\n');
      } else {
        streams.err('pptx-studio: ' + describe(error) + '\n');
      }
      return 1;
    }
  }

  if (command === 'roundtrip') {
    const options: RoundTripOptions = {
      json: values.json === true,
      quiet: values.quiet === true,
      out: values.out ?? null,
      write: values.write ?? null,
    };
    try {
      return runRoundTrip(file, options, streams.out);
    } catch (error) {
      // Three ways to fail and one exit code, so the message has to carry the
      // distinction. A package that will not open, an export the firewall
      // refused, and a comparison that found a difference are three different
      // pieces of work; a script only wants one bit, and a person wants to know
      // which of the three it was.
      if (isValidateError(error)) {
        streams.err(
          'pptx-studio: the export was refused, so there was nothing to compare.\n  ' +
            error.message +
            '\n',
        );
      } else if (isOpcError(error)) {
        streams.err('pptx-studio: ' + error.code + ': ' + error.message + '\n');
      } else {
        streams.err('pptx-studio: ' + describe(error) + '\n');
      }
      return 1;
    }
  }

  if (command === 'validate') {
    const options: ValidateOptions = {
      json: values.json === true,
      explain: values.explain === true,
      quiet: values.quiet === true,
      out: values.out ?? null,
    };
    try {
      return runValidate(file, options, streams.out);
    } catch (error) {
      // The same split `inspect` makes below. A package that will not open at
      // all is a different answer from a package with findings in it, and only
      // the second one is a report.
      if (isOpcError(error)) {
        streams.err('pptx-studio: ' + error.code + ': ' + error.message + '\n');
        return 1;
      }
      streams.err('pptx-studio: ' + describe(error) + '\n');
      return 1;
    }
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

/** A numeric flag, or null when what was given is not one. */
function positiveInteger(value: string | undefined, fallback: number): number | null {
  if (value === undefined) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

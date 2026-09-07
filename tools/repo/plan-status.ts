/**
 * `docs/plan/README.md` - where the build has got to, against every sub-phase.
 *
 * ```
 * node tools/repo/plan-status.ts            # write it
 * node tools/repo/plan-status.ts --check    # fail if it is out of date
 * ```
 *
 * Status is derived, not asserted: a sub-phase is done when `docs/plan/phases.json` names its ADR
 * and that file exists. The working agreement writes the ADR at the end of a sub-phase, so nothing
 * can be marked done here that has no record.
 */

import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { format, resolveConfig } from 'prettier';

import { repoPath } from './root.ts';

const PLAN = repoPath('docs/plan');
const DATA = join(PLAN, 'phases.json');
const INDEX = join(PLAN, 'README.md');
const ADR = repoPath('docs/adr');

/** The one state a sub-phase may claim without an ADR to show for it. */
const IN_PROGRESS = 'in-progress';

interface SubPhase {
  readonly id: string;
  readonly title: string;
  readonly adr?: string;
  readonly result?: string;
  readonly note?: string;
  readonly state?: string;
}

interface Gate {
  readonly title: string;
  readonly adr?: string;
  readonly result?: string;
  readonly note?: string;
}

interface Phase {
  readonly number: number;
  readonly title: string;
  readonly subPhases: readonly SubPhase[];
  readonly gate: Gate;
}

interface Carried {
  readonly title: string;
  readonly detail: string;
  readonly raisedIn: readonly string[];
}

interface Plan {
  readonly phases: readonly Phase[];
  readonly carried: readonly Carried[];
}

function fields(value: unknown, where: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${where} is not an object`);
  }
  return value as Record<string, unknown>;
}

function text(value: unknown, where: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${where} is not a non-empty string`);
  }
  return value;
}

function optionalText(value: unknown, where: string): string | undefined {
  return value === undefined ? undefined : text(value, where);
}

function list(value: unknown, where: string): readonly unknown[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`${where} is not a non-empty array`);
  }
  return value as readonly unknown[];
}

/** Every ADR a sub-phase names has to be a file, or the table would link to nothing. */
function adrOf(value: unknown, where: string): string | undefined {
  const path = optionalText(value, `${where}.adr`);
  if (path === undefined) return undefined;
  if (!existsSync(join(ADR, path)))
    throw new Error(`${where}.adr names ${path}, which is not a file`);
  return path;
}

function readSubPhase(value: unknown, phase: number, at: number): SubPhase {
  const where = `phase ${String(phase)} sub-phase ${String(at)}`;
  const raw = fields(value, where);
  const id = text(raw['id'], `${where}.id`);
  if (!new RegExp(String.raw`^${String(phase)}\.\d+$`).test(id)) {
    throw new Error(`${where} is called ${id}, which is not a sub-phase of phase ${String(phase)}`);
  }
  const adr = adrOf(raw['adr'], where);
  const state = optionalText(raw['state'], `${where}.state`);
  if (state !== undefined && state !== IN_PROGRESS) {
    throw new Error(`${id} claims state ${state}; the only state is ${IN_PROGRESS}`);
  }
  if (state !== undefined && adr !== undefined) {
    throw new Error(`${id} is ${IN_PROGRESS} and has an ADR; a recorded sub-phase is done`);
  }
  const result = optionalText(raw['result'], `${where}.result`);
  const note = optionalText(raw['note'], `${where}.note`);
  return {
    id,
    title: text(raw['title'], `${where}.title`),
    ...(adr === undefined ? {} : { adr }),
    ...(result === undefined ? {} : { result }),
    ...(note === undefined ? {} : { note }),
    ...(state === undefined ? {} : { state }),
  };
}

function readPhase(value: unknown, at: number): Phase {
  const where = `phase at index ${String(at)}`;
  const raw = fields(value, where);
  const number = raw['number'];
  if (typeof number !== 'number' || !Number.isInteger(number) || number < 0) {
    throw new Error(`${where} has no whole phase number`);
  }
  const subPhases = list(raw['subPhases'], `${where}.subPhases`).map((entry, index) =>
    readSubPhase(entry, number, index + 1),
  );
  const seen = new Set<string>();
  for (const subPhase of subPhases) {
    if (seen.has(subPhase.id)) throw new Error(`${subPhase.id} appears twice`);
    seen.add(subPhase.id);
  }
  const gateWhere = `phase ${String(number)}.gate`;
  const gate = fields(raw['gate'], gateWhere);
  const gateAdr = adrOf(gate['adr'], gateWhere);
  const gateResult = optionalText(gate['result'], `${gateWhere}.result`);
  const gateNote = optionalText(gate['note'], `${gateWhere}.note`);
  return {
    number,
    title: text(raw['title'], `${where}.title`),
    subPhases,
    gate: {
      title: text(gate['title'], `${gateWhere}.title`),
      ...(gateAdr === undefined ? {} : { adr: gateAdr }),
      ...(gateResult === undefined ? {} : { result: gateResult }),
      ...(gateNote === undefined ? {} : { note: gateNote }),
    },
  };
}

function readPlan(): Plan {
  const raw = fields(JSON.parse(readFileSync(DATA, 'utf8')), 'phases.json');
  const phases = list(raw['phases'], 'phases').map((entry, index) => readPhase(entry, index));
  for (let at = 1; at < phases.length; at++) {
    const previous = phases[at - 1]?.number ?? -1;
    const current = phases[at]?.number ?? -1;
    if (current <= previous)
      throw new Error(`phase ${String(current)} does not follow ${String(previous)}`);
  }
  const carried = list(raw['carried'], 'carried').map((entry, index) => {
    const where = `carried[${String(index)}]`;
    const item = fields(entry, where);
    return {
      title: text(item['title'], `${where}.title`),
      detail: text(item['detail'], `${where}.detail`),
      raisedIn: list(item['raisedIn'], `${where}.raisedIn`).map((source, sourceAt) =>
        text(source, `${where}.raisedIn[${String(sourceAt)}]`),
      ),
    };
  });
  return { phases, carried };
}

const DONE = '✅ done';
const CAVEAT = '⚠️ done';
const WORKING = '🟡 in progress';
const TODO = '⬜ not started';

interface Entry {
  readonly adr?: string;
  readonly result?: string;
  readonly note?: string;
  readonly state?: string;
}

function statusOf(entry: Entry): string {
  if (entry.adr === undefined) return entry.state === IN_PROGRESS ? WORKING : TODO;
  return entry.note === undefined ? DONE : CAVEAT;
}

/** The ADR as a link reading `0035`, then what it measured, then what is still open. */
function record(entry: Entry): string {
  const parts: string[] = [];
  if (entry.adr !== undefined) {
    const number = /\/(\d{4})-/.exec(entry.adr)?.[1] ?? entry.adr;
    parts.push(`[${number}](../adr/${entry.adr})`);
  }
  if (entry.result !== undefined) parts.push(entry.result);
  if (entry.note !== undefined) parts.push(entry.note);
  return parts.join(' — ');
}

function done(phase: Phase): number {
  return phase.subPhases.filter((subPhase) => subPhase.adr !== undefined).length;
}

function render(plan: Plan): string {
  const total = plan.phases.reduce((sum, phase) => sum + phase.subPhases.length, 0);
  const built = plan.phases.reduce((sum, phase) => sum + done(phase), 0);
  const gates = plan.phases.filter((phase) => phase.gate.adr !== undefined).length;

  const lines: string[] = [
    '# Where the build has got to',
    '',
    `**${String(built)} of ${String(total)} sub-phases recorded, ${String(gates)} of ${String(plan.phases.length)} gates closed.**`,
    '',
    'A sub-phase is done when it has an ADR, because that is when the working agreement says the',
    'record gets written. ⚠️ means done with something still on the record — the note says what.',
    '',
    '**This file is generated** from `docs/plan/phases.json` by `tools/repo/plan-status.ts`;',
    '`pnpm references` fails if it is out of date. Update the JSON at the end of a sub-phase, in the',
    'same commit as its ADR.',
    '',
    '| phase | sub-phases | gate |',
    '| --- | --- | --- |',
  ];
  for (const phase of plan.phases) {
    const count = `${String(done(phase))}/${String(phase.subPhases.length)}`;
    const gate = phase.gate.adr === undefined ? TODO : statusOf(phase.gate);
    lines.push(
      `| [${String(phase.number)} — ${phase.title}](#phase-${String(phase.number)}) | ${count} | ${gate} |`,
    );
  }

  for (const phase of plan.phases) {
    lines.push(
      '',
      `<a id="phase-${String(phase.number)}"></a>`,
      '',
      `## Phase ${String(phase.number)} — ${phase.title}`,
      '',
      '| # | sub-phase | status | record |',
      '| --- | --- | --- | --- |',
    );
    for (const subPhase of phase.subPhases) {
      lines.push(
        `| ${subPhase.id} | ${subPhase.title} | ${statusOf(subPhase)} | ${record(subPhase)} |`,
      );
    }
    lines.push(
      `| **Gate ${String(phase.number)}** | ${phase.gate.title} | ${statusOf(phase.gate)} | ${record(phase.gate)} |`,
    );
  }

  lines.push(
    '',
    '## Carried debt',
    '',
    'Open across sub-phases, and not owned by any of them.',
    '',
  );
  for (const item of plan.carried) {
    lines.push(
      `### ${item.title}`,
      '',
      item.detail,
      '',
      `Raised in: ${item.raisedIn.join(', ')}.`,
      '',
    );
  }
  return lines.join('\n');
}

/** Formatted through Prettier so that `pnpm format:check` and this agree. */
const config = await resolveConfig(INDEX);
const wanted = await format(render(readPlan()), { ...config, filepath: INDEX });

if (process.argv.includes('--check')) {
  let found: string;
  try {
    found = readFileSync(INDEX, 'utf8');
  } catch {
    found = '';
  }
  if (found !== wanted) {
    console.error('docs/plan/README.md is out of date - run: node tools/repo/plan-status.ts');
    process.exitCode = 1;
  }
} else {
  writeFileSync(INDEX, wanted);
  console.log(`wrote ${INDEX}`);
}

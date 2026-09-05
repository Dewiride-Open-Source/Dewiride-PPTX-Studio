import { PartStore, readZip, type ReadZipOptions, type ZipArchive } from '@pptx-studio/opc';
import { createContext, type Context, type RuntimeContext } from './context.js';
import { ValidateError } from './errors.js';
import {
  v001ContentTypeCoverage,
  v002ContentTypeMap,
  v003ArchiveShape,
  v004PartNames,
  v005MainPart,
} from './rules/package.js';
import {
  v006ReferencesResolve,
  v007RelationshipIds,
  v008TargetsResolve,
  v009RequiredEdges,
} from './rules/relationships.js';
import { v010SchemaOrder, v011ExtLstLast, v012UnexpectedChild } from './rules/order.js';
import {
  v013NotesSize,
  v014ColorMap,
  v015ShapeTreePrologue,
  v016TextBody,
  v017GraphicFrame,
} from './rules/required.js';
import { v018SlideIds, v019SheetIds, v020ShapeIds, v021PlaceholderIndices } from './rules/id.js';
import {
  v022PlaceholderType,
  v023GeometryGuides,
  v024SeriesText,
  v025Control,
  v026ChartStyleComplete,
} from './rules/refusal.js';
import {
  v027UneditedPartsUnchanged,
  v028OpaqueContainersUnchanged,
  v029TextAndFieldIdentity,
} from './rules/preservation.js';
import {
  buildReport,
  findingKey,
  type Finding,
  type Report,
  type SkippedRule,
} from './report/report.js';
import { ruleById, RULE_IDS, type RuleId } from './rules/rules.js';

/**
 * The twenty-nine, wired to their implementations.
 *
 * A plain table, so that "is every rule reachable" is a thing a test can ask
 * rather than a thing a reader has to believe. `rules.test.ts` asserts that the
 * keys here are exactly `RULE_IDS` - a rule with a definition and no function
 * would otherwise sit in the table looking enforced.
 */
const IMPLEMENTATIONS: Readonly<Record<RuleId, (ctx: Context) => void>> = {
  V001: v001ContentTypeCoverage,
  V002: v002ContentTypeMap,
  V003: v003ArchiveShape,
  V004: v004PartNames,
  V005: v005MainPart,
  V006: v006ReferencesResolve,
  V007: v007RelationshipIds,
  V008: v008TargetsResolve,
  V009: v009RequiredEdges,
  V010: v010SchemaOrder,
  V011: v011ExtLstLast,
  V012: v012UnexpectedChild,
  V013: v013NotesSize,
  V014: v014ColorMap,
  V015: v015ShapeTreePrologue,
  V016: v016TextBody,
  V017: v017GraphicFrame,
  V018: v018SlideIds,
  V019: v019SheetIds,
  V020: v020ShapeIds,
  V021: v021PlaceholderIndices,
  V022: v022PlaceholderType,
  V023: v023GeometryGuides,
  V024: v024SeriesText,
  V025: v025Control,
  V026: v026ChartStyleComplete,
  V027: v027UneditedPartsUnchanged,
  V028: v028OpaqueContainersUnchanged,
  V029: v029TextAndFieldIdentity,
};

/** Rules that cannot run without the archive bytes. See `Context.archive`. */
const ARCHIVE_RULES: readonly RuleId[] = ['V003'];

export interface ValidateOptions {
  /**
   * The package about to be handed over.
   *
   * On an export this is the store the writer emitted from, **not** a store
   * re-opened from the written bytes. The difference is `V027`: a re-opened
   * store thinks every part came from the archive, so the one rule that asks
   * "did anything change that nobody edited" would have no history to ask
   * about and would answer yes about every deliberate edit.
   */
  readonly store?: PartStore;
  /**
   * The archive, when the caller has it.
   *
   * Required for `V003` and for the duplicate half of `V002` - a duplicate
   * `<Default>` is visible only in the markup, because the parsed content-type
   * map collapses one that agrees with itself. Given a `store` and no `bytes`,
   * those are skipped and the report says so.
   */
  readonly bytes?: Uint8Array;
  /** The package as it was opened. The six preservation rules need it. */
  readonly baseline?: PartStore;
  /** The baseline's archive, so an inherited `V003` can be recognised as inherited. */
  readonly baselineBytes?: Uint8Array;
  /** A subset to run. Defaults to all twenty-nine. */
  readonly rules?: readonly RuleId[];
  /** Passed to `readZip` when `bytes` are given. */
  readonly zip?: ReadZipOptions;
}

function openArchive(
  bytes: Uint8Array | undefined,
  zip: ReadZipOptions | undefined,
): ZipArchive | null {
  if (bytes === undefined) return null;
  try {
    return readZip(bytes, zip ?? {});
  } catch {
    // The store below will fail on the same bytes and say so with a code. A
    // second, worse-worded copy of that message helps nobody.
    return null;
  }
}

/**
 * Check a package against the rules.
 *
 * Never throws for anything it finds - a broken package produces a report with
 * findings in it, which is the shape a caller can render. It throws only when
 * it was handed something it cannot open at all.
 */
export function validatePackage(options: ValidateOptions): Report {
  const archive = openArchive(options.bytes, options.zip);
  const store = options.store ?? openStore(options);

  const requested = resolveRules(options.rules);
  const skipped: SkippedRule[] = [];
  const running: RuleId[] = [];

  for (const id of requested) {
    const rule = ruleById(id)!;
    if (rule.needsBaseline === true && options.baseline === undefined) {
      skipped.push({
        rule: id,
        why:
          'it compares against the package as it was opened, and none was supplied. ' +
          'A preservation rule with nothing to compare against reports nothing, which looks ' +
          'exactly like finding nothing.',
      });
      continue;
    }
    if (ARCHIVE_RULES.includes(id) && archive === null) {
      skipped.push({
        rule: id,
        why: 'it reads the archive itself, and only a PartStore was supplied.',
      });
      continue;
    }
    running.push(id);
  }

  const ctx = createContext({
    store,
    bytes: options.bytes ?? null,
    archive,
    baseline: options.baseline ?? null,
  });
  run(ctx, running);

  const findings = attributeOrigins(ctx, options, running);

  return buildReport({
    findings,
    checked: running,
    skipped,
    problems: ctx.problems,
  });
}

function openStore(options: ValidateOptions): PartStore {
  if (options.bytes === undefined) {
    throw new ValidateError(
      'ERR_UNVALIDATABLE',
      'validatePackage needs either a `store` or the archive `bytes`; it was given neither.',
    );
  }
  try {
    return PartStore.open(options.bytes, options.zip);
  } catch (error) {
    throw new ValidateError(
      'ERR_UNVALIDATABLE',
      'the package could not be opened, so there is nothing to validate: ' +
        (error instanceof Error ? error.message : String(error)),
    );
  }
}

function resolveRules(requested: readonly RuleId[] | undefined): readonly RuleId[] {
  if (requested === undefined) return RULE_IDS;
  for (const id of requested) {
    if (ruleById(id) === undefined) {
      throw new ValidateError('ERR_UNKNOWN_RULE', 'there is no rule ' + id, { rule: id });
    }
  }
  // Run in table order whatever order the caller asked in, so a report reads
  // the same way however it was produced.
  return RULE_IDS.filter((id) => requested.includes(id));
}

function run(ctx: RuntimeContext, rules: readonly RuleId[]): void {
  for (const id of rules) IMPLEMENTATIONS[id](ctx);
}

/**
 * Decide, for each finding, whether we introduced it.
 *
 * The second pass is the whole of it: run the same rules against the package as
 * it was opened and difference the two sets. A finding in both was already
 * there. See `report.ts` for why that question is the one that decides whether
 * an export is refused, and why answering it by differencing beats answering it
 * per rule.
 *
 * Three properties make it cheap enough to be unconditional in the only case
 * that matters. It runs **only when something fatal was found**, so a clean
 * export - which is nearly all of them - pays nothing. It runs **only the rules
 * that fired**, not all twenty-nine. And it skips the preservation rules, which
 * would be comparing the baseline against itself and would find nothing by
 * construction.
 */
function attributeOrigins(
  ctx: RuntimeContext,
  options: ValidateOptions,
  running: readonly RuleId[],
): readonly Finding[] {
  const baseline = options.baseline;
  if (baseline === undefined) return ctx.findings;

  const fatal = ctx.findings.filter((finding) => finding.severity === 'fatal');
  if (fatal.length === 0) {
    return ctx.findings.map((finding) => ({ ...finding, origin: 'introduced' as const }));
  }

  const firedRules = new Set(fatal.map((finding) => finding.rule));
  const replay = running.filter((id) => firedRules.has(id) && ruleById(id)!.needsBaseline !== true);

  const before = createContext({
    store: baseline,
    bytes: options.baselineBytes ?? null,
    archive: openArchive(options.baselineBytes, options.zip),
    baseline: null,
  });
  run(before, replay);
  const inherited = new Set(before.findings.map(findingKey));

  return ctx.findings.map((finding) => ({
    ...finding,
    origin: inherited.has(findingKey(finding)) ? ('inherited' as const) : ('introduced' as const),
  }));
}

/**
 * Validate, and refuse to go on if we broke something.
 *
 * The one call an export path makes. It throws rather than returning a boolean
 * because the failure has to be impossible to ignore: PowerPoint emits no
 * diagnostic log, its refusal message names no part and no line, and a user who
 * gets a repair prompt has no way at all to find out why. This is the only
 * feedback loop that exists, and a caller who forgot to check a return value
 * would have removed it.
 */
export function assertValid(options: ValidateOptions): Report {
  const report = validatePackage(options);
  if (report.ok) return report;
  const blocking = report.findings.filter(
    (finding) => finding.severity === 'fatal' && finding.origin !== 'inherited',
  );
  throw new ValidateError(
    'ERR_VALIDATION_FAILED',
    'refusing to hand over the package: ' +
      String(report.blocking) +
      ' fatal finding(s) this session introduced. First: ' +
      blocking[0]!.rule +
      ' at ' +
      blocking[0]!.where.part +
      (blocking[0]!.where.xpath === null ? '' : ' ' + blocking[0]!.where.xpath) +
      ' - ' +
      blocking[0]!.message,
    { part: blocking[0]!.where.part, report },
  );
}

/**
 * `@pptx-studio/validate` - the repair firewall.
 *
 * PowerPoint emits no diagnostic log. When it refuses a file the message names
 * no part, no element and no line: "PowerPoint could not open the file", or
 * `0x80070570`, "the file or directory is corrupted and unreadable". When it
 * *repairs* one it says even less, and the user is left with a deck that has
 * quietly lost something.
 *
 * That is the whole reason this package exists. It runs on every export, in
 * development and in production, and refuses to hand over bytes when a rule
 * this session broke would produce one of those messages. There is no other
 * feedback loop.
 *
 * ## Twenty-nine rules, in two halves
 *
 * Roughly half come from ECMA-376: content-type coverage, `xsd:sequence` child
 * order, `minOccurs`, the four identifier ranges. Those could be derived from
 * the schemas, and one of them is - the ordering table is generated.
 *
 * The other half cannot be derived from anything. They are the record of a
 * package built with **one** change in it, opened in PowerPoint 16.0.20326, and
 * declined: a `p:ph type="hdr"` on a slide, a geometry guide referenced but
 * never defined, a `c:strLit` inside `c:tx`, a `p:control` in any of eight
 * forms, a `cs:chartStyle` with thirty of its thirty-one entries, a master id
 * that collides with a layout id. Every one of those is schema-legal. Each
 * rule's `why` carries what was tried and what opened, so that whoever
 * eventually contradicts one knows what they are contradicting.
 *
 * ## It answers "did *we* break it", not "is this valid"
 *
 * A fatal finding in a file the user just imported is information. The same
 * finding in a file they just edited is a bug in us. Refusing both would mean a
 * deck with one pre-existing defect could be opened here and never saved again
 * - the editor declining to give somebody back their own file over a problem it
 * did not cause. So findings carry an `origin`, computed by running the rules a
 * second time against the package as it was opened and differencing the two
 * reports, and only the ones we introduced refuse an export.
 *
 * ## What passing does not mean
 *
 * Necessary, not sufficient. PowerPoint rejects some schema-legal markup for
 * reasons nobody has enumerated, and this package knows the ones we have found.
 * The strongest defence is not this: it is architectural - *never synthesize
 * markup we did not read*. The second strongest is `cli bisect`, which reduces
 * a deck PowerPoint refused until the smallest reproducing change is left, and
 * whose output is how the `measured` rules below got here.
 *
 * Everything here runs in a browser tab and in a Web Worker. There is no Node.
 */

export {
  ValidateError,
  VALIDATE_ERROR_CODES,
  isValidateError,
  type ValidateErrorCode,
  type ValidateErrorDetail,
} from './errors.js';

export {
  RULES,
  RULE_IDS,
  BASELINE_RULES,
  ruleById,
  type Evidence,
  type Rule,
  type RuleCategory,
  type RuleId,
  type Severity,
} from './rules.js';

export {
  buildReport,
  findingKey,
  formatReport,
  isReport,
  type Finding,
  type FormatOptions,
  type Origin,
  type ReadProblem,
  type Report,
  type SkippedRule,
} from './report.js';

export {
  attributeLocation,
  elementLocation,
  inDocumentOrder,
  lineColumn,
  PACKAGE_LOCATION,
  partLocation,
  rootName,
  xpathOf,
  xpathOfAttribute,
  type Location,
} from './location.js';

export { createContext, type Context, type ContextInput } from './context.js';

export { assertValid, validatePackage, type ValidateOptions } from './validate.js';

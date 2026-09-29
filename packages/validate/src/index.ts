/**
 * `@pptx-studio/validate` - the repair firewall, run on every export: it refuses to hand over
 * bytes a rule this session broke. Half the rules are ECMA-376, half measured refusals and
 * repairs; findings inherited from the opened file report and do not refuse. See the README.
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
} from './rules/rules.js';

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
} from './report/report.js';

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
} from './report/location.js';

export { createContext, type Context, type ContextInput } from './context.js';

export { assertValid, validatePackage, type ValidateOptions } from './validate.js';

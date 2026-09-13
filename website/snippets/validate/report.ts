import { formatReport, validatePackage } from '@pptx-studio/validate';

declare const bytes: Uint8Array;

//#region example
const report = validatePackage({ bytes }); // never throws for a finding
report.ok; // false when anything fatal fired
report.findings; // [{ rule: 'V022', severity: 'fatal', where: { part, xpath }, message, origin }]
report.skipped; // the baseline rules, with why
console.log(formatReport(report, { explain: true }));
//#endregion

'use client';

import { useMemo, useState } from 'react';

import { RULES, validatePackage, type RuleCategory } from '@pptx-studio/validate';

import { useDeck } from '@/deck/provider';
import { useDocument } from '@/deck/document';
import { Badge } from '@/shell/badge';
import { Mono, Snippet } from '@/shell/code';
import { Panel, Stat } from '@/shell/panel';
import { Cell, Row, Table } from '@/shell/table';

const HOW = `import { validatePackage, assertValid } from '@pptx-studio/validate';

// Reading someone else's file: findings are information.
const report = validatePackage({ store, bytes });
report.findings;   // rule, severity, part, xpath, message
report.skipped;    // rules that could not run, and why - never silent

// Writing one: assertValid throws rather than returning a boolean,
// because a repair prompt names no part and no line.
assertValid({ store, bytes, baseline, baselineBytes });`;

const CATEGORIES: readonly RuleCategory[] = [...new Set(RULES.map((rule) => rule.category))];

export default function ValidatePage() {
  const { deck } = useDeck();
  const { parsed, error, loading } = useDocument();
  const [open, setOpen] = useState<string | null>(null);

  const report = useMemo(() => {
    if (parsed === null || deck === null) return null;
    try {
      return validatePackage({ store: parsed.store, bytes: deck.bytes });
    } catch {
      return null;
    }
  }, [parsed, deck]);

  if (loading) return <p className="text-sm text-ink-400">Reading the deck…</p>;
  if (error !== null) return <p className="text-sm text-handle">{error}</p>;
  if (parsed === null || report === null) return null;

  const fired = new Set(report.findings.map((finding) => finding.rule));
  const skipped = new Set(report.skipped.map((one) => one.rule));

  return (
    <div className="flex flex-col gap-4">
      <header>
        <h1 className="text-lg font-semibold text-ink-100">Validate</h1>
        <p className="text-[13px] text-ink-400">
          <Mono>@pptx-studio/validate</Mono> - the twenty-nine rules that stand between an export
          and a repair prompt.
        </p>
      </header>

      <Panel
        aside={
          report.ok ? (
            <Badge tone="good">nothing blocking</Badge>
          ) : (
            <Badge tone="bad">blocked</Badge>
          )
        }
      >
        <div className="grid grid-cols-2 gap-4 p-4 sm:grid-cols-5">
          <Stat label="rules" value={RULES.length} />
          <Stat label="ran" value={report.checked.length} />
          <Stat
            label="skipped"
            value={report.skipped.length}
            tone={report.skipped.length > 0 ? 'warn' : 'plain'}
          />
          <Stat
            label="findings"
            value={report.findings.length}
            tone={report.findings.length > 0 ? 'warn' : 'good'}
          />
          <Stat
            label="blocking"
            value={report.blocking}
            tone={report.blocking > 0 ? 'bad' : 'good'}
          />
        </div>
      </Panel>

      {report.skipped.length === 0 ? null : (
        <Panel
          title="Rules that did not run"
          hint="Six need the package as it was opened. Reported nothing and found nothing are opposite meanings."
        >
          <Table head={['rule', 'why']}>
            {report.skipped.map((one) => (
              <Row key={one.rule}>
                <Cell className="whitespace-nowrap">
                  <Mono>{one.rule}</Mono>
                </Cell>
                <Cell className="text-ink-400">{one.why}</Cell>
              </Row>
            ))}
          </Table>
        </Panel>
      )}

      {report.findings.length === 0 ? null : (
        <Panel title="Findings">
          <Table head={['rule', 'severity', 'origin', 'where', 'message']}>
            {report.findings.map((finding, at) => (
              <Row key={at}>
                <Cell className="whitespace-nowrap">
                  <Mono>{finding.rule}</Mono>
                </Cell>
                <Cell>
                  <Badge tone={finding.severity === 'fatal' ? 'bad' : 'warn'}>
                    {finding.severity}
                  </Badge>
                </Cell>
                <Cell>
                  <Badge tone={finding.origin === 'introduced' ? 'bad' : 'plain'}>
                    {finding.origin}
                  </Badge>
                </Cell>
                <Cell>
                  <Mono tone="dim">{finding.where.part}</Mono>
                  {finding.where.xpath === null ? null : (
                    <div>
                      <Mono tone="dim">{finding.where.xpath}</Mono>
                    </div>
                  )}
                </Cell>
                <Cell>{finding.message}</Cell>
              </Row>
            ))}
          </Table>
        </Panel>
      )}

      <Panel
        title="The rules"
        hint="Half come from ECMA-376. The other half are decks PowerPoint declined, one change at a time."
      >
        <div className="divide-y divide-ink-800">
          {CATEGORIES.map((category) => (
            <div key={category} className="p-4">
              <h3 className="mb-2 text-[11px] tracking-wider text-ink-400 uppercase">{category}</h3>
              <div className="flex flex-col gap-1">
                {RULES.filter((rule) => rule.category === category).map((rule) => (
                  <div key={rule.id}>
                    <button
                      type="button"
                      onClick={() => setOpen((was) => (was === rule.id ? null : rule.id))}
                      className="flex w-full items-baseline gap-2 rounded px-2 py-1.5 text-left hover:bg-ink-800"
                    >
                      <Mono>{rule.id}</Mono>
                      {fired.has(rule.id) ? (
                        <Badge tone="bad">fired</Badge>
                      ) : skipped.has(rule.id) ? (
                        <Badge tone="warn">skipped</Badge>
                      ) : (
                        <Badge tone="good">passed</Badge>
                      )}
                      <span className="min-w-0 flex-1 text-[13px] text-ink-200">{rule.title}</span>
                      <Badge tone={rule.evidence === 'measured' ? 'info' : 'plain'}>
                        {rule.evidence}
                      </Badge>
                    </button>
                    {open === rule.id ? (
                      <p className="px-2 pb-2 pl-14 text-[12px] leading-relaxed text-ink-400">
                        {rule.why}
                      </p>
                    ) : null}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </Panel>

      <Panel title="How">
        <div className="p-4">
          <Snippet code={HOW} />
        </div>
      </Panel>
    </div>
  );
}

'use client';

import { useEffect, useState } from 'react';

import { RULES, formatReport, type Report, type RuleCategory } from '@pptx-studio/validate';

import { useDeck } from '@/deck/provider';
import { DeckWorkerError, useDeckWorker } from '@/deck/worker/client';
import { Badge, Chip } from '@/design/badge';
import { Button } from '@/design/button';
import { Callout } from '@/design/callout';
import { Code, Mono } from '@/design/code';
import { Switch } from '@/design/field';
import { Panel, Stat, StatRow } from '@/design/panel';
import { ErrorState, Status, type Failure } from '@/design/state';
import { Cell, Row, Table } from '@/design/table';
import type { DemoProps } from '../registry';

const HOW = `import { validatePackage, assertValid, formatReport } from '@pptx-studio/validate';

// Reading someone else's file: findings are information.
const report = validatePackage({ bytes });
report.findings;   // rule, severity, part, xpath, message, origin
report.skipped;    // rules that could not run, and why - never silent
formatReport(report, { explain: true });

// Writing one: assertValid throws rather than returning a boolean,
// because a repair prompt names no part and no line.
assertValid({ store, bytes, baseline, baselineBytes });`;

const CATEGORIES: readonly RuleCategory[] = [...new Set(RULES.map((rule) => rule.category))];

export default function ValidateDemo({ full }: DemoProps) {
  const { deck } = useDeck();
  const worker = useDeckWorker();
  const [result, setResult] = useState<{ report: Report; ms: number } | null>(null);
  const [failure, setFailure] = useState<Failure | null>(null);
  const [open, setOpen] = useState<string | null>(null);
  const [category, setCategory] = useState<RuleCategory | null>(null);
  const [onlyFired, setOnlyFired] = useState(false);
  const [asText, setAsText] = useState(false);

  useEffect(() => {
    const current = worker.current;
    if (deck === null || current === null) return;
    let live = true;
    current
      .validate(deck.bytes)
      .then((done) => {
        if (!live) return;
        setResult(done);
        setFailure(null);
      })
      .catch((cause: unknown) => {
        if (live)
          setFailure(cause instanceof DeckWorkerError ? cause.failure : { message: String(cause) });
      });
    return () => {
      live = false;
    };
  }, [deck, worker]);

  if (failure !== null) return <ErrorState {...failure} />;
  if (result === null) return <Status busy>Running the 31 rules in the Worker…</Status>;

  const { report } = result;
  const fired = new Set(report.findings.map((finding) => finding.rule));
  const skipped = new Set(report.skipped.map((one) => one.rule));
  const shown = RULES.filter(
    (rule) =>
      (category === null || rule.category === category) && (!onlyFired || fired.has(rule.id)),
  );

  return (
    <div className="flex flex-col gap-4">
      <Panel
        aside={
          <div className="flex items-center gap-2">
            {report.ok ? (
              <Badge tone="good">nothing blocking</Badge>
            ) : (
              <Badge tone="bad">blocked</Badge>
            )}
            <Button size="sm" onClick={() => setAsText((on) => !on)} aria-pressed={asText}>
              {asText ? 'As tables' : 'As the CLI prints it'}
            </Button>
          </div>
        }
      >
        <StatRow>
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
        </StatRow>
        <p className="border-t border-line px-4 py-3 text-[12px] text-fg-muted">
          Ran in the Worker in {result.ms.toFixed(0)} ms. Passing is necessary, not sufficient: the
          rules are the ways PowerPoint is known to refuse a file, not every way it could.
        </p>
      </Panel>

      {asText ? (
        <Code
          lang="txt"
          code={formatReport(report, { explain: true })}
          title="pptx-studio validate"
        />
      ) : (
        <>
          {report.skipped.length === 0 ? null : (
            <Panel
              title="Rules that did not run"
              hint="Three need the package as it was opened. Reported nothing and found nothing are opposite meanings."
            >
              <Table head={['rule', 'why']} caption="Rules that were skipped">
                {report.skipped.map((one) => (
                  <Row key={one.rule}>
                    <Cell className="whitespace-nowrap">
                      <Mono>{one.rule}</Mono>
                    </Cell>
                    <Cell className="text-fg-muted">{one.why}</Cell>
                  </Row>
                ))}
              </Table>
            </Panel>
          )}

          {report.findings.length === 0 ? null : (
            <Panel title="Findings">
              <Table head={['rule', 'severity', 'origin', 'where', 'message']} caption="Findings">
                {report.findings.map((finding, at) => (
                  <Row key={at}>
                    <Cell className="whitespace-nowrap">
                      <a href={`#${finding.rule}`}>
                        <Mono>{finding.rule}</Mono>
                      </a>
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
            aside={
              <Switch checked={onlyFired} onChange={setOnlyFired}>
                only fired
              </Switch>
            }
          >
            <div className="flex flex-wrap gap-1.5 border-b border-line px-4 py-3">
              <Chip selected={category === null} onClick={() => setCategory(null)}>
                all
              </Chip>
              {CATEGORIES.map((one) => (
                <Chip key={one} selected={category === one} onClick={() => setCategory(one)}>
                  {one}
                </Chip>
              ))}
            </div>
            <div className="flex flex-col gap-1 p-2">
              {shown.length === 0 ? (
                <p className="p-2 text-[13px] text-fg-muted">No rule fired in this category.</p>
              ) : null}
              {shown.map((rule) => (
                <div key={rule.id} id={rule.id}>
                  <button
                    type="button"
                    onClick={() => setOpen((was) => (was === rule.id ? null : rule.id))}
                    aria-expanded={open === rule.id}
                    className="flex w-full flex-wrap items-baseline gap-2 rounded-control px-2 py-1.5 text-left hover:bg-sunken"
                  >
                    <Mono>{rule.id}</Mono>
                    {fired.has(rule.id) ? (
                      <Badge tone="bad">fired</Badge>
                    ) : skipped.has(rule.id) ? (
                      <Badge tone="warn">skipped</Badge>
                    ) : (
                      <Badge tone="good">passed</Badge>
                    )}
                    <span className="min-w-0 flex-1 text-[13px] text-fg">{rule.title}</span>
                    <Badge tone={rule.evidence === 'schema' ? 'plain' : 'info'}>
                      {rule.evidence}
                    </Badge>
                    <Badge>{rule.category}</Badge>
                  </button>
                  {open === rule.id ? (
                    <p className="px-2 pb-2 pl-14 text-[12px] leading-relaxed text-fg-muted">
                      {rule.why}
                    </p>
                  ) : null}
                </div>
              ))}
            </div>
          </Panel>
        </>
      )}

      {full ? (
        <>
          <Panel title="How this page does it">
            <div className="p-4">
              <Code code={HOW} />
            </div>
          </Panel>
          <Callout kind="honest">
            <p>
              Three rules compare a package against the one it was opened from, so they run only on
              an export; the writer demo shows them. Passing here does not prove PowerPoint opens
              the file: no hosted runner has Office, and Gate 1 was measured by hand.
            </p>
          </Callout>
        </>
      ) : null}
    </div>
  );
}

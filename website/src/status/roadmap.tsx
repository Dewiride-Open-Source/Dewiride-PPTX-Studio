import { Badge, type Tone } from '@/design/badge';
import { REPOSITORY } from '@/site/navigation';
import { plan, type SubPhaseState } from './plan';

const STATE: Record<SubPhaseState, { label: string; tone: Tone }> = {
  done: { label: 'done', tone: 'good' },
  caveat: { label: 'done, with a note', tone: 'warn' },
  'in-progress': { label: 'in progress', tone: 'accent' },
  todo: { label: 'not started', tone: 'plain' },
};

function adrLink(adr: string | null) {
  if (adr === null) return null;
  const number = /\/(\d{4})-/.exec(adr)?.[1] ?? adr;
  return (
    <a
      href={`${REPOSITORY}/blob/main/docs/adr/${adr}`}
      target="_blank"
      rel="noreferrer"
      className="font-mono text-[12px] text-accent hover:underline"
    >
      ADR {number}
    </a>
  );
}

/** Every phase of the plan, as the plan file records it; no dates, because the plan has none. */
export function Roadmap() {
  return (
    <div className="not-prose space-y-8">
      <p className="text-[14px] text-fg-muted">
        <span className="tabular font-mono text-fg">
          {plan.recorded} of {plan.total}
        </span>{' '}
        sub-phases recorded,{' '}
        <span className="tabular font-mono text-fg">
          {plan.gatesClosed} of {plan.gates}
        </span>{' '}
        gates closed. A sub-phase is done when it has an architecture decision record, because that
        is when the working agreement says the record gets written.
      </p>
      {plan.phases.map((phase) => (
        <section key={phase.number} id={`phase-${String(phase.number)}`}>
          <h3 className="flex flex-wrap items-baseline gap-x-3 text-base font-semibold">
            <span>
              Phase {phase.number} - {phase.title}
            </span>
            <span className="tabular font-mono text-[12px] font-normal text-fg-muted">
              {phase.done}/{phase.total}
            </span>
          </h3>
          <div className="mt-2 overflow-x-auto rounded-panel border border-line">
            <table className="w-full border-collapse text-left text-[13px]">
              <caption className="sr-only">Sub-phases of phase {phase.number}</caption>
              <thead>
                <tr className="border-b border-line bg-sunken">
                  <th scope="col" className="px-3 py-2 font-medium text-fg-muted">
                    #
                  </th>
                  <th scope="col" className="px-3 py-2 font-medium text-fg-muted">
                    Sub-phase
                  </th>
                  <th scope="col" className="px-3 py-2 font-medium text-fg-muted">
                    Status
                  </th>
                  <th scope="col" className="px-3 py-2 font-medium text-fg-muted">
                    Record
                  </th>
                </tr>
              </thead>
              <tbody>
                {phase.subPhases.map((subPhase) => (
                  <tr key={subPhase.id} className="border-b border-line/60 last:border-0">
                    <td className="px-3 py-2 font-mono text-[12px] whitespace-nowrap text-fg-muted">
                      {subPhase.id}
                    </td>
                    <td className="px-3 py-2 text-fg">{subPhase.title}</td>
                    <td className="px-3 py-2">
                      <Badge tone={STATE[subPhase.state].tone}>{STATE[subPhase.state].label}</Badge>
                    </td>
                    <td className="px-3 py-2 text-fg-muted">
                      {adrLink(subPhase.adr)}
                      {subPhase.note === null ? null : <p className="mt-1">{subPhase.note}</p>}
                    </td>
                  </tr>
                ))}
                <tr className="bg-sunken/60">
                  <td className="px-3 py-2 font-mono text-[12px] whitespace-nowrap text-fg-muted">
                    Gate
                  </td>
                  <td className="px-3 py-2 font-medium text-fg">{phase.gate.title}</td>
                  <td className="px-3 py-2">
                    <Badge tone={STATE[phase.gate.state].tone}>
                      {STATE[phase.gate.state].label}
                    </Badge>
                  </td>
                  <td className="px-3 py-2">{adrLink(phase.gate.adr)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>
      ))}
      <section id="carried">
        <h3 className="text-base font-semibold">Carried debt</h3>
        <p className="mt-1 text-[13px] text-fg-muted">
          Open across sub-phases, and owned by none of them.
        </p>
        <ul className="mt-3 space-y-3">
          {plan.carried.map((item) => (
            <li
              key={item.title}
              className="rounded-panel border border-line bg-surface p-4 text-[13px]"
            >
              <p className="font-medium text-fg">{item.title}</p>
              <p className="mt-1 text-fg-muted">{item.detail}</p>
              <p className="mt-1 text-[12px] text-fg-faint">
                Raised in {item.raisedIn.join(', ')}.
              </p>
            </li>
          ))}
        </ul>
      </section>
    </div>
  );
}

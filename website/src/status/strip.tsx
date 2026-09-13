import Link from 'next/link';

import { plan } from './plan';

const COLUMNS: readonly { title: string; tone: string; items: readonly string[]; note: string }[] =
  [
    {
      title: 'Drawn today',
      tone: 'text-good',
      items: [
        'Shapes, groups and pictures',
        'Fills, gradients, patterns',
        'Strokes, arrowheads, effects',
        'Text with bullets, fields and autofit',
        'Vertical and rotated text',
        'A hundred slides at any zoom',
      ],
      note: 'Phases 0 to 3, gates 0 to 3 closed.',
    },
    {
      title: 'Preserved, not drawn',
      tone: 'text-warn',
      items: [
        'Tables',
        'Charts',
        'SmartArt',
        'OLE objects',
        'Ink and 3-D models',
        'Video, audio, animations, macros',
      ],
      note: 'Carried through an export byte for byte; shown as frames until their phase draws them.',
    },
    {
      title: 'Not yet',
      tone: 'text-fg-muted',
      items: [
        'Select, resize, rotate',
        'Text editing',
        'Layout switching and theme verbs',
        'Font embedding',
        'Charts, drawn',
        '1.0',
      ],
      note: 'Phases 4 to 12 of the plan, in that order.',
    },
  ];

/** What works, what is carried, what is not built - from the same file the plan table is made of. */
export function StatusStrip() {
  return (
    <section aria-labelledby="status" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <div className="flex flex-wrap items-end justify-between gap-4">
        <div>
          <h2 id="status" className="text-2xl font-semibold tracking-tight">
            What works today
          </h2>
          <p className="mt-2 text-[14px] text-fg-muted">
            <span className="tabular font-mono text-fg">
              {plan.gatesClosed} of {plan.gates}
            </span>{' '}
            gates closed ·{' '}
            <span className="tabular font-mono text-fg">
              {plan.recorded} of {plan.total}
            </span>{' '}
            sub-phases recorded ·{' '}
            <span className="tabular font-mono text-fg">
              {plan.roundTrip.decks}/{plan.roundTrip.total}
            </span>{' '}
            decks round-trip in CI
          </p>
        </div>
        <Link href="/docs/status" className="text-[13px] text-accent hover:underline">
          Full status and roadmap →
        </Link>
      </div>
      <div className="mt-6 grid gap-4 md:grid-cols-3">
        {COLUMNS.map((column) => (
          <div
            key={column.title}
            className="rounded-panel border border-line bg-surface p-5 shadow-panel"
          >
            <h3 className={`text-[12px] font-semibold tracking-wider uppercase ${column.tone}`}>
              {column.title}
            </h3>
            <ul className="mt-3 space-y-1.5 text-[13px] text-fg">
              {column.items.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
            <p className="mt-3 text-[12px] text-fg-faint">{column.note}</p>
          </div>
        ))}
      </div>
    </section>
  );
}

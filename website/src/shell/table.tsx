import type { ReactNode } from 'react';

/** Every table here is wide, so each one scrolls inside its own box. */
export function Table({ head, children }: { head: readonly ReactNode[]; children: ReactNode }) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-left text-[13px]">
        <thead>
          <tr className="border-b border-ink-700">
            {head.map((cell, at) => (
              <th
                key={at}
                className="px-4 py-2 text-[11px] font-medium tracking-wide text-ink-400 uppercase whitespace-nowrap"
              >
                {cell}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}

export function Row({
  children,
  onClick,
  selected = false,
}: {
  children: ReactNode;
  onClick?: () => void;
  selected?: boolean;
}) {
  return (
    <tr
      onClick={onClick}
      className={`border-b border-ink-800 last:border-0 ${
        onClick === undefined ? '' : 'cursor-pointer'
      } ${selected ? 'bg-chrome/10' : 'hover:bg-ink-800/60'}`}
    >
      {children}
    </tr>
  );
}

export function Cell({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <td className={`px-4 py-2 align-top text-ink-200 ${className}`}>{children}</td>;
}

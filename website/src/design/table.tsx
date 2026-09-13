import type { ReactNode } from 'react';

interface TableProps {
  readonly head: readonly ReactNode[];
  readonly caption: string;
  readonly children: ReactNode;
}

/** Every table here is wide, so each one scrolls inside its own box. */
export function Table({ head, caption, children }: TableProps) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-left text-[13px]">
        <caption className="sr-only">{caption}</caption>
        <thead>
          <tr className="border-b border-line">
            {head.map((cell, at) => (
              <th
                key={at}
                scope="col"
                className="px-4 py-2 text-[11px] font-medium tracking-wide text-fg-muted uppercase whitespace-nowrap"
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

interface RowProps {
  readonly children: ReactNode;
  readonly onClick?: () => void;
  readonly selected?: boolean;
}

/** A row that can be chosen is a real button for the keyboard, not a clickable `<tr>`. */
export function Row({ children, onClick, selected = false }: RowProps) {
  const chosen = selected ? 'bg-accent-soft' : onClick === undefined ? '' : 'hover:bg-sunken';
  return (
    <tr
      onClick={onClick}
      onKeyDown={
        onClick === undefined
          ? undefined
          : (event) => {
              if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                onClick();
              }
            }
      }
      tabIndex={onClick === undefined ? undefined : 0}
      aria-selected={onClick === undefined ? undefined : selected}
      className={`border-b border-line/60 last:border-0 ${onClick === undefined ? '' : 'cursor-pointer'} ${chosen}`}
    >
      {children}
    </tr>
  );
}

export function Cell({ children, className = '' }: { children: ReactNode; className?: string }) {
  return <td className={`px-4 py-2 align-top text-fg ${className}`}>{children}</td>;
}

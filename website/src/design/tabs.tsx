'use client';

import { useId, useState, type ReactNode } from 'react';

export interface Tab<T extends string> {
  readonly value: T;
  readonly label: string;
}

interface TabsProps<T extends string> {
  readonly label: string;
  readonly tabs: readonly Tab<T>[];
  readonly initial: T;
  readonly children: (active: T) => ReactNode;
}

/** A tab list with the keyboard behaviour a tab list has; only the active panel is rendered. */
export function Tabs<T extends string>({ label, tabs, initial, children }: TabsProps<T>) {
  const [active, setActive] = useState<T>(initial);
  const id = useId();
  const move = (from: number, by: number) => {
    const next = tabs[(from + by + tabs.length) % tabs.length];
    if (next !== undefined) {
      setActive(next.value);
      document.getElementById(`${id}-tab-${next.value}`)?.focus();
    }
  };
  return (
    <div>
      <div role="tablist" aria-label={label} className="flex flex-wrap gap-1 border-b border-line">
        {tabs.map((tab, at) => {
          const on = tab.value === active;
          return (
            <button
              key={tab.value}
              id={`${id}-tab-${tab.value}`}
              type="button"
              role="tab"
              aria-selected={on}
              aria-controls={`${id}-panel`}
              tabIndex={on ? 0 : -1}
              onClick={() => setActive(tab.value)}
              onKeyDown={(event) => {
                if (event.key === 'ArrowRight') move(at, 1);
                else if (event.key === 'ArrowLeft') move(at, -1);
              }}
              className={`-mb-px border-b-2 px-3 py-2 text-[13px] font-medium transition-colors ${
                on ? 'border-accent text-fg' : 'border-transparent text-fg-muted hover:text-fg'
              }`}
            >
              {tab.label}
            </button>
          );
        })}
      </div>
      <div id={`${id}-panel`} role="tabpanel" className="pt-4">
        {children(active)}
      </div>
    </div>
  );
}

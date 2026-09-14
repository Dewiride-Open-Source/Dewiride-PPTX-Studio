'use client';

import { useId } from 'react';

export interface Segment<T extends string> {
  readonly value: T;
  readonly label: string;
}

interface SegmentedProps<T extends string> {
  readonly label: string;
  readonly options: readonly Segment<T>[];
  readonly value: T;
  readonly onChange: (value: T) => void;
  readonly size?: 'sm' | 'md';
}

/** One choice out of a few, as a radio group that looks like a row of buttons. */
export function Segmented<T extends string>({
  label,
  options,
  value,
  onChange,
  size = 'md',
}: SegmentedProps<T>) {
  const name = useId();
  const pad = size === 'sm' ? 'h-7 px-2 text-[12px]' : 'h-8 px-3 text-[13px]';
  return (
    <fieldset className="inline-flex rounded-control border border-line-strong bg-surface p-0.5">
      <legend className="sr-only">{label}</legend>
      {options.map((option) => {
        const on = option.value === value;
        return (
          <label
            key={option.value}
            className={`inline-flex cursor-pointer items-center rounded-[4px] font-medium transition-colors duration-150 has-focus-visible:outline-2 has-focus-visible:outline-offset-1 has-focus-visible:outline-accent ${pad} ${
              on ? 'bg-accent text-on-accent' : 'text-fg-muted hover:bg-sunken hover:text-fg'
            }`}
          >
            <input
              type="radio"
              name={name}
              value={option.value}
              checked={on}
              onChange={() => onChange(option.value)}
              className="sr-only"
            />
            {option.label}
          </label>
        );
      })}
    </fieldset>
  );
}

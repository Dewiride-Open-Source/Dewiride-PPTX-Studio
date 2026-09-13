import { useId, type ComponentProps, type ReactNode } from 'react';

const CONTROL =
  'rounded-control border border-line-strong bg-surface text-fg text-[13px] transition-colors duration-150 hover:border-accent focus:border-accent';

interface LabelledProps {
  readonly label: string;
  readonly hint?: string;
  readonly children: (id: string) => ReactNode;
  readonly inline?: boolean;
}

/** A visible label bound to one control by id; every input on the site has one. */
export function Labelled({ label, hint, children, inline = false }: LabelledProps) {
  const id = useId();
  return (
    <div className={inline ? 'flex items-center gap-2' : 'flex flex-col gap-1'}>
      <label htmlFor={id} className="text-[12px] font-medium text-fg-muted">
        {label}
      </label>
      {children(id)}
      {hint === undefined ? null : <span className="text-[11px] text-fg-faint">{hint}</span>}
    </div>
  );
}

interface SelectProps extends Omit<ComponentProps<'select'>, 'className'> {
  readonly className?: string;
}

export function Select({ className = '', ...rest }: SelectProps) {
  return <select className={`${CONTROL} h-8 px-2 ${className}`} {...rest} />;
}

interface TextInputProps extends Omit<ComponentProps<'input'>, 'className' | 'type'> {
  readonly className?: string;
  readonly mono?: boolean;
}

export function TextInput({ className = '', mono = false, ...rest }: TextInputProps) {
  return (
    <input
      type="text"
      className={`${CONTROL} h-8 px-2 ${mono ? 'font-mono text-[12px]' : ''} ${className}`}
      {...rest}
    />
  );
}

interface SliderProps {
  readonly id?: string;
  readonly value: number;
  readonly min: number;
  readonly max: number;
  readonly step?: number;
  readonly onChange: (value: number) => void;
  readonly format?: (value: number) => string;
  readonly ariaLabel?: string;
}

/** A range input with its value always in view. */
export function Slider({
  id,
  value,
  min,
  max,
  step = 1,
  onChange,
  format,
  ariaLabel,
}: SliderProps) {
  return (
    <span className="inline-flex items-center gap-2">
      <input
        id={id}
        type="range"
        min={min}
        max={max}
        step={step}
        value={value}
        aria-label={ariaLabel}
        onChange={(event) => onChange(Number(event.target.value))}
        className="h-1.5 w-36 cursor-pointer accent-accent"
      />
      <span className="tabular min-w-12 font-mono text-[12px] text-fg-muted">
        {format === undefined ? String(value) : format(value)}
      </span>
    </span>
  );
}

interface SwitchProps {
  readonly id?: string;
  readonly checked: boolean;
  readonly onChange: (checked: boolean) => void;
  readonly children: ReactNode;
}

/** An on/off control that says what it switches. */
export function Switch({ id, checked, onChange, children }: SwitchProps) {
  return (
    <label
      htmlFor={id}
      className="inline-flex cursor-pointer items-center gap-2 text-[13px] text-fg"
    >
      <input
        id={id}
        type="checkbox"
        role="switch"
        checked={checked}
        aria-checked={checked}
        onChange={(event) => onChange(event.target.checked)}
        className="peer sr-only"
      />
      <span
        aria-hidden="true"
        className="relative h-5 w-9 rounded-full border border-line-strong bg-sunken transition-colors peer-checked:border-accent peer-checked:bg-accent peer-focus-visible:outline-2 peer-focus-visible:outline-offset-2 peer-focus-visible:outline-accent after:absolute after:top-0.5 after:left-0.5 after:h-3.5 after:w-3.5 after:rounded-full after:bg-surface after:transition-transform peer-checked:after:translate-x-4"
      />
      {children}
    </label>
  );
}

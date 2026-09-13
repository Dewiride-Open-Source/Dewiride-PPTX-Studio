import Link from 'next/link';
import type { ComponentProps, ReactNode } from 'react';

export type ButtonVariant = 'primary' | 'secondary' | 'ghost' | 'danger';
export type ButtonSize = 'sm' | 'md';

const BASE =
  'inline-flex items-center justify-center gap-1.5 rounded-control font-medium whitespace-nowrap transition-colors duration-150 ease-out-quart disabled:cursor-not-allowed disabled:opacity-40';

const VARIANT: Record<ButtonVariant, string> = {
  primary: 'bg-accent text-white hover:brightness-110 border border-transparent',
  secondary: 'border border-line-strong bg-surface text-fg hover:border-accent hover:text-accent',
  ghost: 'border border-transparent text-fg-muted hover:bg-sunken hover:text-fg',
  danger: 'border border-bad/40 bg-bad/10 text-bad hover:bg-bad/20',
};

const SIZE: Record<ButtonSize, string> = {
  sm: 'h-7 px-2.5 text-[12px]',
  md: 'h-9 px-3.5 text-[13px]',
};

export function buttonClass(variant: ButtonVariant = 'secondary', size: ButtonSize = 'md'): string {
  return `${BASE} ${VARIANT[variant]} ${SIZE[size]}`;
}

interface ButtonProps extends Omit<ComponentProps<'button'>, 'className'> {
  readonly variant?: ButtonVariant;
  readonly size?: ButtonSize;
  readonly className?: string;
}

export function Button({ variant, size, className = '', type = 'button', ...rest }: ButtonProps) {
  return <button type={type} className={`${buttonClass(variant, size)} ${className}`} {...rest} />;
}

interface LinkButtonProps {
  readonly href: string;
  readonly variant?: ButtonVariant;
  readonly size?: ButtonSize;
  readonly className?: string;
  readonly children: ReactNode;
}

/** A link that reads as a button; external hrefs open in a new tab. */
export function LinkButton({ href, variant, size, className = '', children }: LinkButtonProps) {
  const external = /^https?:/.test(href);
  return (
    <Link
      href={href}
      className={`${buttonClass(variant, size)} ${className}`}
      {...(external ? { target: '_blank', rel: 'noreferrer' } : {})}
    >
      {children}
    </Link>
  );
}

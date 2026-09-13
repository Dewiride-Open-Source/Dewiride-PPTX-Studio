import type { PackageName } from '@/site/packages';
import { readReference } from './read';

/** The error class a package throws and every code it carries, from the installed d.ts. */
export function ErrorCodes({ pkg }: { pkg: PackageName }) {
  const reference = readReference(pkg);
  if (reference.errors === null) {
    return (
      <p className="text-[13px] text-fg-muted">
        This package declares no error class of its own; what it throws comes from the packages
        beneath it.
      </p>
    );
  }
  return (
    <div className="not-prose">
      <p className="text-[13px] text-fg-muted">
        Throws <code className="font-mono text-fg">{reference.errors.className}</code>, whose{' '}
        <code className="font-mono">code</code> is one of:
      </p>
      <ul className="mt-2 grid gap-1 sm:grid-cols-2 lg:grid-cols-3">
        {reference.errors.codes.map((code) => (
          <li
            key={code}
            className="rounded-control bg-sunken px-2 py-1 font-mono text-[12px] text-fg"
          >
            {code}
          </li>
        ))}
      </ul>
    </div>
  );
}

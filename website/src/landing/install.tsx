'use client';

import { useState } from 'react';

import { Chip } from '@/design/badge';
import { Code } from '@/design/code';
import { Segmented } from '@/design/segmented';
import { PACKAGES, type PackageName } from '@/site/packages';

type Manager = 'npm' | 'pnpm' | 'yarn' | 'bun';

const COMMAND: Record<Manager, string> = {
  npm: 'npm install',
  pnpm: 'pnpm add',
  yarn: 'yarn add',
  bun: 'bun add',
};

const PRESETS: readonly { label: string; packages: readonly PackageName[] }[] = [
  { label: 'Draw slides', packages: ['opc', 'model', 'render-svg'] },
  { label: 'Show a live slide', packages: ['opc', 'model', 'render-svg', 'render-dom'] },
  { label: 'Read what is inside', packages: ['census'] },
  { label: 'Edit and save', packages: ['opc', 'xml', 'model', 'writer'] },
  { label: 'Server thumbnails', packages: ['cli'] },
  { label: 'Everything', packages: PACKAGES.map((one) => one.name) },
];

const same = (a: readonly PackageName[], b: readonly PackageName[]): boolean =>
  a.length === b.length && a.every((name) => b.includes(name));

/** The install line, built from what the visitor wants to do. */
export function Install() {
  const [manager, setManager] = useState<Manager>('npm');
  const [chosen, setChosen] = useState<readonly PackageName[]>(PRESETS[0]?.packages ?? []);

  const ordered = PACKAGES.map((one) => one.name).filter((name) => chosen.includes(name));
  const line =
    ordered.length === 0
      ? `${COMMAND[manager]} @pptx-studio/<package>`
      : `${COMMAND[manager]} ${ordered.map((name) => `@pptx-studio/${name}`).join(' ')}`;

  return (
    <section aria-labelledby="install" className="mx-auto max-w-6xl px-4 py-12 sm:px-6">
      <div className="flex flex-wrap items-end justify-between gap-4">
        <div>
          <h2 id="install" className="text-2xl font-semibold tracking-tight">
            Install
          </h2>
          <p className="mt-2 text-[14px] text-fg-muted">
            Pick what you want to do, or tick the packages yourself.
          </p>
        </div>
        <Segmented
          label="Package manager"
          size="sm"
          options={(['npm', 'pnpm', 'yarn', 'bun'] as const).map((value) => ({
            value,
            label: value,
          }))}
          value={manager}
          onChange={setManager}
        />
      </div>
      <div className="mt-4 flex flex-wrap gap-2">
        {PRESETS.map((preset) => (
          <Chip
            key={preset.label}
            selected={same(preset.packages, chosen)}
            onClick={() => setChosen(preset.packages)}
          >
            {preset.label}
          </Chip>
        ))}
      </div>
      <div className="mt-4">
        <Code lang="sh" code={line} />
      </div>
      <fieldset className="mt-4 flex flex-wrap gap-x-4 gap-y-2">
        <legend className="mb-2 text-[12px] text-fg-muted">Or choose the packages:</legend>
        {PACKAGES.map((one) => (
          <label key={one.name} className="inline-flex items-center gap-1.5 text-[13px]">
            <input
              type="checkbox"
              checked={chosen.includes(one.name)}
              onChange={(event) =>
                setChosen(
                  event.target.checked
                    ? [...chosen, one.name]
                    : chosen.filter((name) => name !== one.name),
                )
              }
              className="accent-accent"
            />
            <code className="font-mono text-[12px]">{one.name}</code>
          </label>
        ))}
      </fieldset>
      <p className="mt-3 text-[12px] text-fg-faint">
        ESM only. TypeScript types are shipped. Node 24 or newer for the CLI.
      </p>
    </section>
  );
}

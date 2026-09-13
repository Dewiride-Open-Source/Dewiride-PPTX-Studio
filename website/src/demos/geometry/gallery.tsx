'use client';

import { useMemo, useState } from 'react';

import {
  BUCKET_MEMBERS,
  getPreset,
  presetNames,
  resolveGeometry,
  type PresetBucketName,
} from '@pptx-studio/geometry';

import { Chip } from '@/design/badge';
import { TextInput } from '@/design/field';

const SWATCH = { w: 96, h: 64 };
const BUCKETS = Object.keys(BUCKET_MEMBERS) as readonly PresetBucketName[];

function Swatch({ name }: { name: string }) {
  const paths = useMemo(() => {
    const preset = getPreset(name);
    return preset === undefined ? [] : resolveGeometry(preset, SWATCH).paths;
  }, [name]);
  return (
    <svg
      viewBox={`0 0 ${String(SWATCH.w)} ${String(SWATCH.h)}`}
      className="h-12 w-full"
      aria-hidden="true"
    >
      {paths.map((path, at) => (
        <path
          key={at}
          d={path.d}
          fill={path.fill === 'none' ? 'none' : 'rgb(59 143 212 / 0.25)'}
          stroke="#3b8fd4"
          strokeWidth={1.2}
        />
      ))}
    </svg>
  );
}

interface GalleryProps {
  readonly chosen: string;
  readonly onChoose: (name: string) => void;
}

/** The 187 presets, searchable by name and by bucket. */
export function Gallery({ chosen, onChoose }: GalleryProps) {
  const names = useMemo(() => presetNames(), []);
  const [query, setQuery] = useState('');
  const [bucket, setBucket] = useState<PresetBucketName | null>(null);

  const shown = useMemo(() => {
    const needle = query.trim().toLowerCase();
    const members: readonly string[] = bucket === null ? names : BUCKET_MEMBERS[bucket];
    return needle === '' ? members : members.filter((one) => one.toLowerCase().includes(needle));
  }, [names, query, bucket]);

  return (
    <div>
      <div className="flex flex-col gap-2 border-b border-line p-3">
        <TextInput
          aria-label="Search presets"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder={`Search ${String(names.length)} shapes`}
          className="w-full"
        />
        <div className="flex flex-wrap gap-1">
          <Chip selected={bucket === null} onClick={() => setBucket(null)}>
            all
          </Chip>
          {BUCKETS.map((one) => (
            <Chip key={one} selected={bucket === one} onClick={() => setBucket(one)}>
              {one} · {BUCKET_MEMBERS[one].length}
            </Chip>
          ))}
        </div>
      </div>
      <div
        role="listbox"
        aria-label="Presets"
        className="grid max-h-[30rem] grid-cols-3 gap-1 overflow-y-auto p-2"
      >
        {shown.map((name) => (
          <button
            key={name}
            type="button"
            role="option"
            aria-selected={name === chosen}
            onClick={() => onChoose(name)}
            title={name}
            className={`rounded-control border p-1.5 ${
              name === chosen
                ? 'border-accent bg-accent-soft'
                : 'border-transparent hover:border-line-strong'
            }`}
          >
            <Swatch name={name} />
            <span className="mt-1 block truncate font-mono text-[10px] text-fg-muted">{name}</span>
          </button>
        ))}
        {shown.length === 0 ? (
          <p className="col-span-3 p-4 text-center text-[13px] text-fg-muted">No preset matches.</p>
        ) : null}
      </div>
    </div>
  );
}

'use client';

import { useRef, useState } from 'react';

import { Badge } from '@/shell/badge';
import { useDeck } from './provider';
import { SAMPLES } from './samples';

/** The deck chooser in the header: the six shipped decks, or one of your own. */
export function DeckPicker() {
  const { deck, loading, load, loadSample } = useDeck();
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  return (
    <div
      onDragOver={(event) => {
        event.preventDefault();
        setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(event) => {
        event.preventDefault();
        setOver(false);
        const file = event.dataTransfer.files[0];
        if (file !== undefined) load(file);
      }}
      className={`flex flex-wrap items-center gap-2 rounded-lg border px-3 py-2 transition-colors ${
        over ? 'border-chrome bg-chrome/10' : 'border-ink-700 bg-ink-850'
      }`}
    >
      <span className="text-xs text-ink-400">Deck</span>
      <select
        value={SAMPLES.some((s) => s.file === deck?.name) ? deck?.name : ''}
        onChange={(event) => {
          if (event.target.value !== '') loadSample(event.target.value);
        }}
        className="rounded border border-ink-600 bg-ink-800 px-2 py-1 font-mono text-[12px] text-ink-100"
      >
        {SAMPLES.some((s) => s.file === deck?.name) ? null : (
          <option value="">{deck?.name ?? 'none'}</option>
        )}
        {SAMPLES.map((sample) => (
          <option key={sample.file} value={sample.file}>
            {sample.name} - {sample.file}
          </option>
        ))}
      </select>

      <button
        type="button"
        onClick={() => input.current?.click()}
        className="rounded border border-ink-600 bg-ink-800 px-2 py-1 text-[12px] text-ink-200 hover:border-chrome hover:text-ink-100"
      >
        Open a .pptx
      </button>
      <input
        ref={input}
        type="file"
        accept=".pptx,.pptm,.ppsx,.potx"
        hidden
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file !== undefined) load(file);
        }}
      />

      {loading ? <Badge tone="info">reading</Badge> : null}
      <span className="text-[11px] text-ink-500">or drop one here - it never leaves the tab</span>
    </div>
  );
}

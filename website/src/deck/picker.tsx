'use client';

import { useRef, useState, type DragEvent } from 'react';

import { Badge } from '@/design/badge';
import { Button } from '@/design/button';
import { Select } from '@/design/field';
import { Status } from '@/design/state';
import { useDeck } from './provider';
import { SAMPLES, sampleNamed } from './samples';

const ACCEPT = '.pptx,.pptm,.ppsx,.potx';

/** The deck chooser: a shipped sample, or one of the visitor's own, by button or by drop. */
export function DeckPicker({ compact = false }: { compact?: boolean }) {
  const { deck, loading, failure, load, loadSample } = useDeck();
  const input = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);

  const sample = deck === null ? undefined : sampleNamed(deck.name);
  const drop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setOver(false);
    const file = event.dataTransfer.files[0];
    if (file !== undefined) load(file);
  };

  return (
    <div
      onDragOver={(event) => {
        event.preventDefault();
        setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={drop}
      className={`flex flex-wrap items-center gap-2 rounded-panel border px-3 py-2 transition-colors ${
        over ? 'border-accent bg-accent-soft' : 'border-line bg-surface'
      }`}
    >
      <label htmlFor="deck-picker" className="text-[12px] font-medium text-fg-muted">
        Deck
      </label>
      <Select
        id="deck-picker"
        value={sample?.file ?? ''}
        onChange={(event) => {
          if (event.target.value !== '') loadSample(event.target.value);
        }}
        className="max-w-[24rem] font-mono text-[12px]"
      >
        {sample === undefined ? <option value="">{deck?.name ?? 'none'}</option> : null}
        {SAMPLES.map((one) => (
          <option key={one.file} value={one.file}>
            {one.file} · {one.slides} slide{one.slides === 1 ? '' : 's'}
          </option>
        ))}
      </Select>
      <Button size="sm" onClick={() => input.current?.click()}>
        Open a .pptx
      </Button>
      <input
        ref={input}
        type="file"
        accept={ACCEPT}
        hidden
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file !== undefined) load(file);
        }}
      />
      {deck?.own === true ? <Badge tone="accent">your file</Badge> : null}
      {loading ? (
        <Status busy>Reading the deck…</Status>
      ) : failure !== null ? (
        <Status>
          <span className="text-bad">{failure.message}</span>
        </Status>
      ) : compact ? null : (
        <span className="text-[11px] text-fg-faint">
          or drop one here · it never leaves your browser
        </span>
      )}
    </div>
  );
}

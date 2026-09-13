'use client';

import { useMemo, useState } from 'react';

import {
  AUTONUMBER_SCHEMES,
  RESERVED_FIELD_TYPES,
  UNMEASURED_SCHEMES,
  formatAutonumber,
  renderField,
  symbolBulletChar,
} from '@pptx-studio/text';

import { Badge } from '@/design/badge';
import { Mono } from '@/design/code';
import { Select, TextInput } from '@/design/field';
import { Panel } from '@/design/panel';
import { Cell, Row, Table } from '@/design/table';

const LANGS = ['en-US', 'en-GB', 'de-DE', 'fr-FR', 'ja-JP', 'es-ES'];

/** Every autonumber scheme at a chosen start, and every reserved field type in a chosen locale. */
export function Bullets() {
  const [start, setStart] = useState(1);
  const [lang, setLang] = useState('en-US');
  const now = useMemo(() => new Date(), []);
  const unmeasured = new Set<string>(UNMEASURED_SCHEMES);

  return (
    <div className="grid gap-4 lg:grid-cols-2">
      <Panel
        title="Autonumber schemes"
        hint="Every ST_TextAutonumberScheme; the numbers are not in the file, they are counted."
        aside={
          <label className="flex items-center gap-2 text-[12px] text-fg-muted">
            start at
            <TextInput
              aria-label="Start at"
              value={String(start)}
              onChange={(event) =>
                setStart(Math.max(1, Math.min(32767, Number(event.target.value) || 1)))
              }
              className="w-16"
            />
          </label>
        }
      >
        <div className="grid grid-cols-2 gap-x-6 gap-y-1 p-4 sm:grid-cols-3">
          {AUTONUMBER_SCHEMES.map((scheme) => (
            <div key={scheme} className="flex items-baseline justify-between gap-2 text-[12px]">
              <Mono tone="dim">{scheme}</Mono>
              {unmeasured.has(scheme) ? (
                <Badge tone="warn">refused</Badge>
              ) : (
                <span className="tabular text-fg">{formatAutonumber(scheme, start)}</span>
              )}
            </div>
          ))}
        </div>
        <p className="border-t border-line px-4 py-3 text-[12px] text-fg-faint">
          Two schemes throw <Mono tone="dim">TEXT_AUTONUMBER_UNMEASURED</Mono>: PowerPoint shapes
          them into glyph ids before drawing, so the measurement holds indices, not characters.
          Symbol bullets map through the face:{' '}
          <Mono tone="dim">symbolBulletChar(&apos;•&apos;, &apos;Wingdings&apos;)</Mono> is{' '}
          <Mono>
            U+{symbolBulletChar('•', 'Wingdings').codePointAt(0)?.toString(16).toUpperCase() ?? '?'}
          </Mono>
          .
        </p>
      </Panel>

      <Panel
        title="Reserved fields"
        hint="slidenum is the position; a measured date pattern is recomputed; everything else shows the cached text."
        aside={
          <Select
            aria-label="Locale"
            value={lang}
            onChange={(event) => setLang(event.target.value)}
            className="font-mono text-[12px]"
          >
            {LANGS.map((one) => (
              <option key={one}>{one}</option>
            ))}
          </Select>
        }
      >
        <div className="max-h-96 overflow-y-auto">
          <Table head={['type', 'shows', 'source']} caption="Reserved field types">
            {RESERVED_FIELD_TYPES.map((type) => {
              const rendered = renderField({
                type,
                lang,
                cachedText: '(cached text)',
                slideNumber: 3,
                now,
              });
              return (
                <Row key={type}>
                  <Cell>
                    <Mono>{type}</Mono>
                  </Cell>
                  <Cell className="tabular">{rendered.text}</Cell>
                  <Cell>
                    <Badge tone={rendered.source === 'computed' ? 'good' : 'warn'}>
                      {rendered.source}
                    </Badge>
                    {rendered.reason === '' ? null : (
                      <span className="ml-2 text-[11px] text-fg-faint">{rendered.reason}</span>
                    )}
                  </Cell>
                </Row>
              );
            })}
          </Table>
        </div>
      </Panel>
    </div>
  );
}

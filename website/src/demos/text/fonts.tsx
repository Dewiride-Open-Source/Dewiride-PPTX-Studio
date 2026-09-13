'use client';

import { useEffect, useState } from 'react';

import { censusPackage } from '@pptx-studio/census';
import { createFontProbe, fontReport, type FontFinding } from '@pptx-studio/text';

import { useDeck } from '@/deck/provider';
import { Badge } from '@/design/badge';
import { Panel } from '@/design/panel';
import { Cell, Row, Table } from '@/design/table';

/** The deck's own typefaces, and what this browser would draw each one in. */
export function Fonts() {
  const { deck } = useDeck();
  const [findings, setFindings] = useState<readonly FontFinding[] | null>(null);

  useEffect(() => {
    if (deck === null) return;
    let live = true;
    // The probe measures through a canvas, so it waits for a frame rather than blocking the paint.
    const frame = requestAnimationFrame(() => {
      const census = censusPackage(deck.bytes);
      const asked = census.typefaces.flatMap((one) =>
        Array.from({ length: one.count }, () => one.name),
      );
      if (live) setFindings(asked.length === 0 ? [] : fontReport(asked, createFontProbe()));
    });
    return () => {
      live = false;
      cancelAnimationFrame(frame);
    };
  }, [deck]);

  const substituted = (findings ?? []).filter((one) => one.status !== 'present');

  return (
    <Panel
      title="Fonts in this deck"
      hint="What it asks for, and what this browser will actually draw it in."
      aside={
        findings === null ? null : substituted.length === 0 ? (
          <Badge tone="good">all present</Badge>
        ) : (
          <Badge tone="warn">{substituted.length} substituted</Badge>
        )
      }
    >
      {findings === null ? (
        <p className="p-4 text-[13px] text-fg-muted">Reading the deck…</p>
      ) : findings.length === 0 ? (
        <p className="p-4 text-[13px] text-fg-muted">This deck names no typeface.</p>
      ) : (
        <Table
          head={['asked for', 'drawn as', 'metrics', 'runs']}
          caption="Typefaces the deck asks for"
        >
          {findings.map((one) => (
            <Row key={one.typeface}>
              <Cell>{one.typeface}</Cell>
              <Cell>
                {one.status === 'present' ? (
                  <Badge tone="good">as asked</Badge>
                ) : one.status === 'substituted' ? (
                  <Badge tone="warn">{one.rendersAs}</Badge>
                ) : (
                  <Badge tone="bad">missing</Badge>
                )}
              </Cell>
              <Cell className="text-[12px] text-fg-muted">
                {one.metricCompatible === null
                  ? 'unknown'
                  : one.metricCompatible
                    ? 'compatible'
                    : 'approximate'}
                {one.verified ? '' : ' (from the table)'}
              </Cell>
              <Cell className="tabular">{one.runs}</Cell>
            </Row>
          ))}
        </Table>
      )}
    </Panel>
  );
}

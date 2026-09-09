'use client';

import { useCallback, useMemo, useRef, useState } from 'react';

import { PartStore } from '@pptx-studio/opc';
import { loadDocument, type Document, type Sheet } from '@pptx-studio/model';
import {
  applyEdit,
  parseXml,
  serializeXml,
  type XDocument,
  type XElement,
  type XmlEdit,
} from '@pptx-studio/xml';

import { NotEditable, applyAll, planMove, planRecolour, planRetext } from './edits';
import { shapeElement } from './locate';

interface Step {
  readonly label: string;
  readonly partName: string;
  readonly inverses: readonly XmlEdit[];
}

interface Session {
  readonly store: PartStore;
  /** The parsed tree per part, kept so edits address the same nodes each time. */
  readonly trees: Map<string, XDocument>;
  past: Step[];
  future: Step[];
}

export interface Editor {
  readonly document: Document;
  readonly store: PartStore;
  /** Bumped on every edit, so memoised views recompute. */
  readonly version: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly lastStep: string | null;
  readonly refused: string | null;
  readonly editedParts: readonly string[];
  move(sheet: Sheet, cNvPrId: number, dx: number, dy: number): void;
  recolour(sheet: Sheet, cNvPrId: number, hex: string): void;
  retext(sheet: Sheet, cNvPrId: number, value: string): void;
  undo(): void;
  redo(): void;
}

/**
 * An editing session over one deck.
 *
 * The `PartStore` is the document; the parsed trees are the edit surface. After
 * every change the touched part is serialised back into the store and the model
 * is loaded again, because a `Sheet` is a projection and a stale one would draw
 * the shape where it used to be.
 */
export function useEditor(bytes: Uint8Array | null): Editor | null {
  const [version, setVersion] = useState(0);
  const [refused, setRefused] = useState<string | null>(null);
  const [lastStep, setLastStep] = useState<string | null>(null);
  const session = useRef<{ key: Uint8Array; value: Session } | null>(null);

  if (bytes !== null && session.current?.key !== bytes) {
    session.current = {
      key: bytes,
      value: { store: PartStore.open(bytes), trees: new Map(), past: [], future: [] },
    };
  }
  const current = bytes === null ? null : (session.current?.value ?? null);

  const treeFor = useCallback((live: Session, partName: string): XDocument => {
    const found = live.trees.get(partName);
    if (found !== undefined) return found;
    const parsed = parseXml(live.store.read(partName));
    live.trees.set(partName, parsed);
    return parsed;
  }, []);

  const commit = useCallback((live: Session, partName: string) => {
    const tree = live.trees.get(partName);
    if (tree !== undefined) live.store.replacePart(partName, serializeXml(tree));
  }, []);

  const run = useCallback(
    (
      label: string,
      sheet: Sheet,
      cNvPrId: number,
      plan: (shape: XElement) => readonly XmlEdit[],
    ) => {
      const live = current;
      if (live === null) return;
      setRefused(null);
      try {
        const tree = treeFor(live, sheet.partName);
        const shape = shapeElement(tree.root, cNvPrId);
        if (shape === undefined) {
          setRefused(`No shape with cNvPr id ${String(cNvPrId)} in ${sheet.partName}.`);
          return;
        }
        const inverses = applyAll(plan(shape));
        commit(live, sheet.partName);
        live.past.push({ label, partName: sheet.partName, inverses });
        live.future = [];
        setLastStep(label);
        setVersion((at) => at + 1);
      } catch (cause) {
        setRefused(
          cause instanceof NotEditable || cause instanceof Error ? cause.message : String(cause),
        );
      }
    },
    [current, treeFor, commit],
  );

  const step = useCallback(
    (from: 'past' | 'future') => {
      const live = current;
      if (live === null) return;
      const taken = from === 'past' ? live.past.pop() : live.future.pop();
      if (taken === undefined) return;
      const back = [...taken.inverses].map((edit) => applyEdit(edit)).reverse();
      commit(live, taken.partName);
      const onto = from === 'past' ? live.future : live.past;
      onto.push({ label: taken.label, partName: taken.partName, inverses: back });
      setLastStep(`${from === 'past' ? 'undo' : 'redo'} ${taken.label}`);
      setVersion((at) => at + 1);
    },
    [current, commit],
  );

  const undo = useCallback(() => step('past'), [step]);
  const redo = useCallback(() => step('future'), [step]);

  return useMemo(() => {
    if (current === null) return null;
    return {
      document: loadDocument(current.store),
      store: current.store,
      version,
      canUndo: current.past.length > 0,
      canRedo: current.future.length > 0,
      lastStep,
      refused,
      editedParts: [...current.trees.keys()],
      move: (sheet, id, dx, dy) => run('move', sheet, id, (shape) => planMove(shape, dx, dy)),
      recolour: (sheet, id, hex) => run('recolour', sheet, id, (shape) => planRecolour(shape, hex)),
      retext: (sheet, id, value) => run('retext', sheet, id, (shape) => planRetext(shape, value)),
      undo,
      redo,
    };
  }, [current, version, lastStep, refused, run, undo, redo]);
}

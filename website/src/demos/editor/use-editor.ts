'use client';

import { useCallback, useRef, useState } from 'react';

import { loadDocument, type Document, type Sheet } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import {
  applyEdit,
  parseXml,
  serializeXml,
  type XDocument,
  type XElement,
  type XmlEdit,
} from '@pptx-studio/xml';

import type { ReplacedPart } from '@/deck/worker/protocol';
import { NotEditable, applyAll, planMove, planRecolour, planRetext } from './edits';
import { shapeElement } from './locate';

interface Step {
  readonly label: string;
  readonly partName: string;
  readonly inverses: readonly XmlEdit[];
}

/** The mutable side: one store, the parsed tree per touched part, and the two stacks. */
interface Session {
  readonly key: Uint8Array;
  readonly store: PartStore;
  readonly trees: Map<string, XDocument>;
  past: Step[];
  future: Step[];
}

/** What a render reads: derived from the session after every operation, never from it directly. */
export interface EditorView {
  readonly key: Uint8Array;
  readonly document: Document;
  readonly version: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly lastStep: string | null;
  readonly refused: string | null;
  readonly editedParts: readonly string[];
}

export interface Editor {
  /** Null until the first edit on this deck; the unedited document is the deck's own. */
  readonly view: EditorView | null;
  readonly move: (sheet: Sheet, cNvPrId: number, dx: number, dy: number) => void;
  readonly recolour: (sheet: Sheet, cNvPrId: number, hex: string) => void;
  readonly retext: (sheet: Sheet, cNvPrId: number, value: string) => void;
  readonly undo: () => void;
  readonly redo: () => void;
  /** Every part rewritten so far, as the Worker's validate and export take them. */
  readonly replaced: () => readonly ReplacedPart[];
}

function treeFor(live: Session, partName: string): XDocument {
  const found = live.trees.get(partName);
  if (found !== undefined) return found;
  const parsed = parseXml(live.store.read(partName));
  live.trees.set(partName, parsed);
  return parsed;
}

function commit(live: Session, partName: string): void {
  const tree = live.trees.get(partName);
  if (tree !== undefined) live.store.replacePart(partName, serializeXml(tree));
}

/**
 * An editing session over one deck.
 *
 * The `PartStore` is the document; the parsed trees are the edit surface. After
 * every change the touched part is serialised back into the store and the model
 * is loaded again, because a `Sheet` is a projection and a stale one would draw
 * the shape where it used to be.
 */
export function useEditor(bytes: Uint8Array | null): Editor {
  const session = useRef<Session | null>(null);
  const [made, setMade] = useState<EditorView | null>(null);
  const view = made !== null && made.key === bytes ? made : null;

  const live = useCallback((): Session | null => {
    if (bytes === null) return null;
    if (session.current === null || session.current.key !== bytes) {
      session.current = {
        key: bytes,
        store: PartStore.open(bytes),
        trees: new Map(),
        past: [],
        future: [],
      };
    }
    return session.current;
  }, [bytes]);

  const publish = useCallback(
    (current: Session, lastStep: string | null, refused: string | null) => {
      setMade((was) => ({
        key: current.key,
        document: loadDocument(current.store),
        version: (was?.key === current.key ? was.version : 0) + 1,
        canUndo: current.past.length > 0,
        canRedo: current.future.length > 0,
        lastStep,
        refused,
        editedParts: [...current.trees.keys()],
      }));
    },
    [],
  );

  const run = useCallback(
    (
      label: string,
      sheet: Sheet,
      cNvPrId: number,
      plan: (shape: XElement) => readonly XmlEdit[],
    ) => {
      const current = live();
      if (current === null) return;
      try {
        const tree = treeFor(current, sheet.partName);
        const shape = shapeElement(tree.root, cNvPrId);
        if (shape === undefined) {
          publish(current, null, `No shape with cNvPr id ${String(cNvPrId)} in ${sheet.partName}.`);
          return;
        }
        const inverses = applyAll(plan(shape));
        commit(current, sheet.partName);
        current.past.push({ label, partName: sheet.partName, inverses });
        current.future = [];
        publish(current, label, null);
      } catch (cause) {
        publish(
          current,
          null,
          cause instanceof NotEditable || cause instanceof Error ? cause.message : String(cause),
        );
      }
    },
    [live, publish],
  );

  const step = useCallback(
    (from: 'past' | 'future') => {
      const current = live();
      if (current === null) return;
      const taken = from === 'past' ? current.past.pop() : current.future.pop();
      if (taken === undefined) return;
      const back = [...taken.inverses].map((edit) => applyEdit(edit)).reverse();
      commit(current, taken.partName);
      const onto = from === 'past' ? current.future : current.past;
      onto.push({ label: taken.label, partName: taken.partName, inverses: back });
      publish(current, `${from === 'past' ? 'undo' : 'redo'} ${taken.label}`, null);
    },
    [live, publish],
  );

  const replaced = useCallback((): readonly ReplacedPart[] => {
    const current = live();
    if (current === null) return [];
    return [...current.trees.keys()].map((part) => {
      const content = current.store.read(part);
      return {
        part,
        bytes: content.buffer.slice(
          content.byteOffset,
          content.byteOffset + content.byteLength,
        ) as ArrayBuffer,
      };
    });
  }, [live]);

  return {
    view,
    move: (sheet, id, dx, dy) => run('move', sheet, id, (shape) => planMove(shape, dx, dy)),
    recolour: (sheet, id, hex) => run('recolour', sheet, id, (shape) => planRecolour(shape, hex)),
    retext: (sheet, id, value) => run('retext', sheet, id, (shape) => planRetext(shape, value)),
    undo: () => step('past'),
    redo: () => step('future'),
    replaced,
  };
}

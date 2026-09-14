'use client';

import { useCallback, useMemo, useRef, useState } from 'react';

import { parseSheet, type Document, type Sheet } from '@pptx-studio/model';
import { PartStore } from '@pptx-studio/opc';
import type { Frame } from '@pptx-studio/render-svg';
import {
  applyEdits,
  parseXml,
  serializeXml,
  type XDocument,
  type XElement,
  type XmlEdit,
} from '@pptx-studio/xml';

import type { ReplacedPart } from '@/deck/worker/protocol';
import { applyAll, planDelete, planMove, planRecolour, planRetext, type Plan } from './edits.ts';
import { nameOf, shapeElement } from './locate.ts';

interface Step {
  readonly label: string;
  partName: string;
  inverses: XmlEdit[];
}

/** The mutable side: one store, the parsed tree per touched part, the document, the two stacks. */
interface Session {
  readonly key: Uint8Array;
  readonly store: PartStore;
  readonly trees: Map<string, XDocument>;
  document: Document;
  past: Step[];
  future: Step[];
  /** The step a drag or a typing session accumulates into until it ends. */
  gesture: Step | null;
}

/** What a render reads: derived from the session after every operation, never from it directly. */
export interface EditorView {
  readonly key: Uint8Array;
  readonly document: Document;
  readonly version: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /** What just happened, for the status line. */
  readonly said: string | null;
  /** Why the last gesture could not be made, in words a visitor can act on. */
  readonly refused: string | null;
  readonly editedParts: readonly string[];
}

export interface Editor {
  /** Null until the first gesture on this deck; the unedited document is the deck's own. */
  readonly view: EditorView | null;
  /** Open a step that the gestures until `end` accumulate into. */
  readonly begin: (label: string) => void;
  readonly end: () => void;
  /** Undo the open step and close it. */
  readonly cancel: () => void;
  readonly move: (sheet: Sheet, cNvPrId: number, frame: Frame, dx: number, dy: number) => boolean;
  readonly recolour: (sheet: Sheet, cNvPrId: number, hex: string) => boolean;
  readonly retext: (
    sheet: Sheet,
    cNvPrId: number,
    paragraphs: readonly (readonly string[])[],
  ) => boolean;
  readonly remove: (sheet: Sheet, cNvPrId: number) => boolean;
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

/** Serialise the touched part back into the store and re-read that one sheet into the document. */
function commit(live: Session, partName: string): void {
  const tree = live.trees.get(partName);
  if (tree === undefined) return;
  live.store.replacePart(partName, serializeXml(tree));
  const index = live.document.slides.findIndex((one) => one.partName === partName);
  const was = live.document.slides[index];
  if (was === undefined) return;
  const sheet: Sheet = { ...parseSheet(tree.root, partName), parent: was.parent, theme: was.theme };
  live.document = { ...live.document, slides: live.document.slides.with(index, sheet) };
}

const said = (cause: unknown): string => (cause instanceof Error ? cause.message : String(cause));

/**
 * An editing session over one deck.
 *
 * The `PartStore` is the document; the parsed trees are the edit surface. After
 * every change the touched part is serialised back into the store and that
 * slide alone is parsed again, because a `Sheet` is a projection and a stale
 * one would draw the shape where it used to be.
 */
export function useEditor(bytes: Uint8Array | null, base: Document | null): Editor {
  const session = useRef<Session | null>(null);
  const [made, setMade] = useState<EditorView | null>(null);
  const view = made !== null && made.key === bytes ? made : null;

  const live = useCallback((): Session | null => {
    if (bytes === null || base === null) return null;
    if (session.current === null || session.current.key !== bytes) {
      session.current = {
        key: bytes,
        store: PartStore.open(bytes),
        trees: new Map(),
        document: base,
        past: [],
        future: [],
        gesture: null,
      };
    }
    return session.current;
  }, [bytes, base]);

  const publish = useCallback((current: Session, what: string | null, refused: string | null) => {
    setMade((was) => ({
      key: current.key,
      document: current.document,
      version: (was?.key === current.key ? was.version : 0) + 1,
      canUndo: current.past.length > 0,
      canRedo: current.future.length > 0,
      said: what,
      refused,
      editedParts: [...current.trees.keys()],
    }));
  }, []);

  const run = useCallback(
    (
      label: (name: string) => string,
      sheet: Sheet,
      cNvPrId: number,
      plan: (shape: XElement) => Plan,
    ): boolean => {
      const current = live();
      if (current === null) return false;
      const tree = treeFor(current, sheet.partName);
      const shape = shapeElement(tree.root, cNvPrId);
      if (shape === undefined) {
        publish(current, null, `No shape #${String(cNvPrId)} on ${sheet.partName}.`);
        return false;
      }
      const what = label(nameOf(shape));
      try {
        const inverses = applyAll(plan(shape));
        if (inverses.length === 0) return true;
        if (current.gesture !== null) {
          current.gesture.partName = sheet.partName;
          current.gesture.inverses = [...inverses, ...current.gesture.inverses];
        } else {
          current.past.push({ label: what, partName: sheet.partName, inverses });
          current.future = [];
        }
        commit(current, sheet.partName);
        publish(current, current.gesture?.label ?? what, null);
        return true;
      } catch (cause) {
        publish(current, null, said(cause));
        return false;
      }
    },
    [live, publish],
  );

  const begin = useCallback(
    (label: string) => {
      const current = live();
      if (current === null) return;
      current.gesture = { label, partName: '', inverses: [] };
    },
    [live],
  );

  const end = useCallback(() => {
    const current = live();
    if (current === null || current.gesture === null) return;
    const taken = current.gesture;
    current.gesture = null;
    if (taken.inverses.length === 0) return;
    current.past.push(taken);
    current.future = [];
    publish(current, taken.label, null);
  }, [live, publish]);

  const cancel = useCallback(() => {
    const current = live();
    if (current === null || current.gesture === null) return;
    const taken = current.gesture;
    current.gesture = null;
    if (taken.inverses.length === 0) return;
    applyEdits(taken.inverses);
    commit(current, taken.partName);
    publish(current, `Cancelled ${taken.label}`, null);
  }, [live, publish]);

  const step = useCallback(
    (from: 'past' | 'future') => {
      const current = live();
      if (current === null) return;
      const taken = from === 'past' ? current.past.pop() : current.future.pop();
      if (taken === undefined) return;
      const back = applyEdits(taken.inverses);
      commit(current, taken.partName);
      const onto = from === 'past' ? current.future : current.past;
      onto.push({ label: taken.label, partName: taken.partName, inverses: back });
      publish(current, `${from === 'past' ? 'Undid' : 'Redid'}: ${taken.label}`, null);
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

  return useMemo(
    () => ({
      view,
      begin,
      end,
      cancel,
      move: (sheet, id, frame, dx, dy) =>
        run(
          (name) => `Moved ${name}`,
          sheet,
          id,
          (shape) => planMove(shape, frame, dx, dy),
        ),
      recolour: (sheet, id, hex) =>
        run(
          (name) => `Coloured ${name} ${hex.toUpperCase()}`,
          sheet,
          id,
          (shape) => planRecolour(shape, hex),
        ),
      retext: (sheet, id, paragraphs) =>
        run(
          (name) => `Retyped ${name}`,
          sheet,
          id,
          (shape) => planRetext(shape, paragraphs),
        ),
      remove: (sheet, id) =>
        run(
          (name) => `Deleted ${name}`,
          sheet,
          id,
          (shape) => planDelete(shape),
        ),
      undo: () => step('past'),
      redo: () => step('future'),
      replaced,
    }),
    [view, begin, end, cancel, run, step, replaced],
  );
}

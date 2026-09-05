/**
 * The package into sheets, and the sheets into a chain.
 *
 * ## Every binding comes from the part's own `.rels`, and from nothing else
 *
 * A slide part does not name its layout. A layout does not name its master. A
 * master does not name its theme. Each of those bindings lives in exactly one
 * place - `ppt/slides/_rels/slideN.xml.rels`, and its two counterparts - and
 * every other candidate is a trap:
 *
 * - **`p:sldLayoutIdLst` is not the binding.** It gives each layout an id and an
 *   order for the UI. A two-master deck saved by PowerPoint puts all 22 layouts
 *   in one flat `ppt/slideLayouts/` folder, numbered 1 to 22, and which master
 *   owns which is recorded only in the two masters' relationship parts.
 * - **The part number is not the binding.** Same reason.
 * - **`ppt/presentation.xml.rels` has a `theme` relationship**, it always points
 *   at `theme1.xml`, and reading it gives every slide the first master's palette.
 *   On a one-master deck - which is every deck most people ever test with - that
 *   is indistinguishable from correct.
 * - **The order of a `.rels` file is not `rId` order.** PowerPoint writes them
 *   in whatever order it likes; `rId8` routinely comes first. Anything that
 *   reads a relationship part positionally is wrong on the files Office writes.
 *
 * ## A broken binding is not an error
 *
 * A slide with no layout relationship, a layout with no master, a master with no
 * theme: PowerPoint repairs all three rather than refusing them, and after the
 * repair the binding is simply gone. So this loads them with `parent` or `theme`
 * as `null` and records the problem, rather than throwing. Refusing to *write*
 * one is `validate`'s job, and it is a different question.
 */

import { REL_TYPE, type PartStore } from '@pptx-studio/opc';
import { parseXml, type XElement } from '@pptx-studio/xml';

import { ModelError } from './errors.js';
import { parseSheet, parseTheme } from './parse/sheet.js';
import { parseDefaultTextStyle } from './parse/text.js';
import type { ListStyle } from './text.js';
import type { Sheet, Theme } from './types.js';

/** One thing wrong with the package that did not stop the load. */
export interface DocumentProblem {
  readonly code: 'MODEL_NO_PARENT' | 'MODEL_NO_THEME' | 'MODEL_PART_MISSING' | 'MODEL_PART_KIND';
  readonly partName: string;
  readonly message: string;
}

const REL_SLIDE = REL_TYPE.slide;
const REL_LAYOUT = REL_TYPE.slideLayout;
const REL_MASTER = REL_TYPE.slideMaster;
const REL_THEME = REL_TYPE.theme;

interface Loading {
  readonly problems: DocumentProblem[];
  readonly sheets: Map<string, Sheet>;
  readonly themes: Map<string, Theme>;
  readonly loading: Set<string>;
  readonly store: PartStore;
}

function rootOf(store: PartStore, partName: string): XElement {
  const document = parseXml(store.read(partName));
  const root = document.root;
  if (root === null) {
    throw new ModelError('MODEL_PART_KIND', `${partName} has no root element`, partName);
  }
  return root;
}

function themeFor(state: Loading, masterPartName: string): Theme | null {
  const rels = state.store.relationships(masterPartName);
  const rel = rels.firstOfType(REL_THEME);
  if (rel === undefined) {
    state.problems.push({
      code: 'MODEL_NO_THEME',
      partName: masterPartName,
      message: 'the master has no theme relationship',
    });
    return null;
  }
  const target = rels.resolve(rel);
  const cached = state.themes.get(target);
  if (cached !== undefined) return cached;
  if (!state.store.has(target)) {
    state.problems.push({
      code: 'MODEL_PART_MISSING',
      partName: target,
      message: `${masterPartName} names a theme part that is not in the package`,
    });
    return null;
  }
  const theme = parseTheme(rootOf(state.store, target), target);
  state.themes.set(target, theme);
  return theme;
}

/**
 * Load one sheet and, recursively, its parent.
 *
 * The `loading` set is a cycle guard rather than an optimisation. A layout whose
 * `slideMaster` relationship points back at a layout is a package a hostile
 * input can produce, and following it without a guard is an unbounded recursion
 * in a Web Worker.
 */
function loadSheet(state: Loading, partName: string): Sheet | null {
  const cached = state.sheets.get(partName);
  if (cached !== undefined) return cached;
  if (state.loading.has(partName)) {
    state.problems.push({
      code: 'MODEL_PART_KIND',
      partName,
      message: 'the sheet chain returns to this part',
    });
    return null;
  }
  if (!state.store.has(partName)) {
    state.problems.push({
      code: 'MODEL_PART_MISSING',
      partName,
      message: 'the part is not in the package',
    });
    return null;
  }

  state.loading.add(partName);
  let parsed;
  try {
    parsed = parseSheet(rootOf(state.store, partName), partName);
  } catch (error) {
    state.loading.delete(partName);
    if (error instanceof ModelError && error.code === 'MODEL_PART_KIND') {
      state.problems.push({ code: 'MODEL_PART_KIND', partName, message: error.message });
      return null;
    }
    throw error;
  }

  let parent: Sheet | null = null;
  let theme: Theme | null = null;

  if (parsed.kind === 'master') {
    theme = themeFor(state, partName);
  } else {
    const wantedType = parsed.kind === 'slide' ? REL_LAYOUT : REL_MASTER;
    const rels = state.store.relationships(partName);
    const rel = rels.firstOfType(wantedType);
    if (rel === undefined) {
      state.problems.push({
        code: 'MODEL_NO_PARENT',
        partName,
        message: `no ${parsed.kind === 'slide' ? 'slideLayout' : 'slideMaster'} relationship`,
      });
    } else {
      parent = loadSheet(state, rels.resolve(rel));
      if (parent !== null && parent.kind === parsed.kind) {
        // A layout whose `slideMaster` relationship points at another layout.
        // The chain would still terminate, and it would be the wrong chain.
        state.problems.push({
          code: 'MODEL_PART_KIND',
          partName,
          message: `the parent of a ${parsed.kind} is another ${parsed.kind}`,
        });
        parent = null;
      }
    }
  }

  state.loading.delete(partName);
  const sheet: Sheet = { ...parsed, parent, theme };
  state.sheets.set(partName, sheet);
  return sheet;
}

/**
 * A loaded presentation: its slides, in `p:sldIdLst` order, each with its chain.
 *
 * Nothing here is a class with behaviour. It is the result of a load, and every
 * question about it is a free function in `resolve.ts`, `style.ts` or
 * `background.ts` - so a caller can hold a `Sheet` from anywhere and still ask
 * everything, and so the whole surface stays tree-shakable.
 */
export interface Document {
  /** The presentation part, for the parts of it 2.9 does not model. */
  readonly presentation: XElement;
  readonly presentationPartName: string;
  /** In `p:sldIdLst` order, which *is* the deck's order. */
  readonly slides: readonly Sheet[];
  /** In `p:sldMasterIdLst` order. */
  readonly masters: readonly Sheet[];
  /** Every layout reachable from a master, in the order its master lists it. */
  readonly layouts: readonly Sheet[];
  readonly themes: readonly Theme[];
  readonly problems: readonly DocumentProblem[];
  /**
   * `p:defaultTextStyle`, the terminus of the text cascade for every shape that
   * reaches no `p:txStyles` bucket - which is every shape that is not a
   * placeholder, and the `dt`, `ftr`, `sldNum` and `hdr` placeholders besides.
   *
   * It lives on the presentation part, so it belongs to the package rather than
   * to any sheet, and `undefined` means the package declared none.
   */
  readonly defaultTextStyle: ListStyle | undefined;
  /** Slide size in EMU, from `p:sldSz`. */
  readonly slideSize: { readonly cx: number; readonly cy: number };
  sheet(partName: string): Sheet | undefined;
}

function officeDocumentOf(store: PartStore): string {
  const rel = store.rootRelationships().firstOfType(REL_TYPE.officeDocument);
  if (rel === undefined) {
    throw new ModelError('MODEL_NO_PRESENTATION', 'the package root names no officeDocument');
  }
  return store.rootRelationships().resolve(rel);
}

export function loadDocument(store: PartStore): Document {
  const presentationPartName = officeDocumentOf(store);
  if (!store.has(presentationPartName)) {
    throw new ModelError(
      'MODEL_NO_PRESENTATION',
      `${presentationPartName} is not in the package`,
      presentationPartName,
    );
  }
  const presentation = rootOf(store, presentationPartName);

  const state: Loading = {
    problems: [],
    sheets: new Map(),
    themes: new Map(),
    loading: new Set(),
    store,
  };

  const rels = store.relationships(presentationPartName);

  // Slides in `p:sldIdLst` order. The list gives the order; the `r:id` on each
  // entry gives the part, through this part's own rels.
  const slides: Sheet[] = [];
  const slideList = presentation.children.find(
    (child) => child.type === 'element' && child.qname === 'p:sldIdLst',
  );
  if (slideList !== undefined && slideList.type === 'element') {
    for (const entry of slideList.children) {
      if (entry.type !== 'element' || entry.qname !== 'p:sldId') continue;
      const id = entry.attributes.find((attr) => attr.qname === 'r:id')?.value;
      if (id === undefined) continue;
      const target = rels.targetOf(id);
      if (target === undefined) continue;
      const sheet = loadSheet(state, target);
      if (sheet !== null) slides.push(sheet);
    }
  }

  // Any slide the list forgot is still a slide, and a package that lost its
  // `p:sldIdLst` should still show something.
  for (const rel of rels.byType(REL_SLIDE)) {
    const target = rels.resolve(rel);
    if (state.sheets.has(target)) continue;
    const sheet = loadSheet(state, target);
    if (sheet !== null) slides.push(sheet);
  }

  const masters: Sheet[] = [];
  for (const rel of rels.byType(REL_MASTER)) {
    const sheet = loadSheet(state, rels.resolve(rel));
    if (sheet !== null) masters.push(sheet);
  }

  // Layouts, through each master's own rels - which is the only record of which
  // master owns which layout.
  const layouts: Sheet[] = [];
  for (const master of masters) {
    const masterRels = store.relationships(master.partName);
    for (const rel of masterRels.byType(REL_LAYOUT)) {
      const sheet = loadSheet(state, masterRels.resolve(rel));
      if (sheet !== null && !layouts.includes(sheet)) layouts.push(sheet);
    }
  }

  const size = presentation.children.find(
    (child) => child.type === 'element' && child.qname === 'p:sldSz',
  );
  const dimension = (name: string): number => {
    if (size === undefined || size.type !== 'element') return 0;
    const raw = size.attributes.find((attr) => attr.qname === name)?.value;
    return raw === undefined ? 0 : Number(raw);
  };

  const sheets = state.sheets;
  return {
    presentation,
    presentationPartName,
    slides,
    masters,
    layouts,
    themes: [...state.themes.values()],
    problems: state.problems,
    defaultTextStyle: parseDefaultTextStyle(presentation, presentationPartName),
    slideSize: { cx: dimension('cx'), cy: dimension('cy') },
    sheet(partName: string): Sheet | undefined {
      return sheets.get(partName);
    },
  };
}

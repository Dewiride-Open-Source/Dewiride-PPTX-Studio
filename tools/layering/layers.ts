/**
 * The dependency graph is one-way, and this file is the only place that says so.
 *
 * The plan's rule in prose: `opc`/`xml` depend on nothing but their externals;
 * `geometry`/`paint`/`text`/`fonts` depend on nothing above them; `model`
 * depends on all of the above; renderers depend on `model`; `react` depends on
 * renderers; **nothing depends on `react`**.
 *
 * Encoded as integer layers. A package may depend on any package in a strictly
 * lower layer, and on same-layer packages only so long as the resulting graph
 * stays acyclic (checked separately) — that concession exists because e.g.
 * `render-dom` legitimately builds on `render-svg` without either being a tier
 * apart.
 *
 * Packages are listed here *before* they exist. That is the point: the guard
 * has to predate the code it guards, or the first violation lands unnoticed.
 */

export type Runtime = 'browser' | 'node';

export interface PackageSpec {
  /** Lower numbers are closer to the leaves. May only depend on <= its own layer. */
  readonly layer: number;
  /**
   * `browser` packages must not import `node:*` builtins and must not declare a
   * `"node"` condition in their exports map. Enforced here at the manifest
   * level and again in eslint.config.mjs at the import level.
   */
  readonly runtime: Runtime;
  /** One line, for the violation message and for docs generation later. */
  readonly role: string;
}

export const SCOPE = '@pptx-studio';

export const PACKAGES: Readonly<Record<string, PackageSpec>> = {
  // Layer 0 — the package format. Depends on nothing but fflate.
  opc: { layer: 0, runtime: 'browser', role: 'OPC container: zip, parts, content types, rels' },
  xml: { layer: 0, runtime: 'browser', role: 'byte-preserving XML tokenizer, XNode, serializer' },
  'fonts-metric-compat': {
    layer: 0,
    runtime: 'browser',
    role: 'OFL metric-compatible substitute faces (separate licence boundary)',
  },

  // Layer 1 — pure computation over OOXML primitives. No document model.
  geometry: {
    layer: 1,
    runtime: 'browser',
    role: 'preset + custom geometry, fmla evaluator, arcs',
  },
  paint: {
    layer: 1,
    runtime: 'browser',
    role: 'colour transforms, gradients, patterns, dashes, effects',
  },
  text: {
    layer: 1,
    runtime: 'browser',
    role: 'measurement, line breaking, layout, autofit, bullets',
  },
  fonts: {
    layer: 1,
    runtime: 'browser',
    role: 'SFNT/EOT read+write, rights gate, FontFace registry',
  },

  // Layer 2 — the document model and the resolver everything else reads.
  model: { layer: 2, runtime: 'browser', role: 'parse, inheritance resolver, commands, history' },

  // Layer 3 — consumers of the model.
  'render-svg': {
    layer: 3,
    runtime: 'browser',
    role: 'string-only SVG emitter (thumbnails, export)',
  },
  'render-dom': {
    layer: 3,
    runtime: 'browser',
    role: 'live DOM renderer: SVG + HTML text + overlay canvas',
  },
  validate: { layer: 3, runtime: 'browser', role: 'the repair firewall — must-not-break rules' },
  writer: {
    layer: 3,
    runtime: 'browser',
    role: 'export: dirty-part serialization, font embed, media GC',
  },

  // Layer 4 — interaction.
  'editor-core': {
    layer: 4,
    runtime: 'browser',
    role: 'selection, OBB resize/rotate, snapping, gestures',
  },
  'text-editor': { layer: 4, runtime: 'browser', role: 'ProseMirror bridge, IME, live autofit' },
  'hard-content': {
    layer: 4,
    runtime: 'browser',
    role: 'tables, SmartArt, charts, images, OLE, ink, 3D',
  },

  // Layer 5 — the only package allowed to know React exists.
  react: { layer: 5, runtime: 'browser', role: 'PptxViewer, PptxEditor, hooks, panels' },

  // Layer 6 — Node entry point. The one place `node:*` is legal.
  cli: {
    layer: 6,
    runtime: 'node',
    role: 'inspect, roundtrip, render, validate, fidelity, bisect',
  },
} as const;

/** The package that owns React. Nothing at a lower layer may reference it. */
export const REACT_OWNER = 'react';

/** Fields every publishable package manifest must carry, with expected values. */
export const REQUIRED_FIELDS = {
  type: 'module',
  license: 'Apache-2.0',
  sideEffects: false,
} as const;

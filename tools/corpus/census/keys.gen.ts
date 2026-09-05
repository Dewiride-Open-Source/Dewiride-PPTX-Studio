/* eslint-disable */
// Generated from @pptx-studio/census. Do not edit.
//
// The Node-side corpus checker runs on a bare clone with no build, so it cannot
// import the census package - nothing links workspace packages into tools/, and
// the resolution that lets tools/bench/bench.test.ts do it comes from Vitest's
// alias table, not from Node. So the key list is committed, and
// census-drift.test.ts asserts it still matches the census's own table.
//
// Regenerate: pnpm build && node tools/corpus/write-census-keys.ts

export const CENSUS_FEATURE_KEYS: readonly string[] = [
  'alternateContent',
  'animation',
  'audio',
  'blipFill',
  'chart',
  'chartEx',
  'comment',
  'connector',
  'contentPart',
  'customGeom',
  'customShow',
  'decorative',
  'embeddedFont',
  'embeddedPackage',
  'field',
  'glow',
  'gradientFill',
  'graphicFrame',
  'group',
  'groupFill',
  'hyperlink',
  'ink',
  'innerShadow',
  'macros',
  'math',
  'media',
  'model3d',
  'notesSlide',
  'oleObject',
  'patternFill',
  'picture',
  'placeholder',
  'presetGeom',
  'reflection',
  'scene3d',
  'section',
  'shadow',
  'shape',
  'smartArt',
  'smartArtDrawing',
  'softEdge',
  'svgBlip',
  'table',
  'thumbnail',
  'transition',
  'video',
  'vml',
];

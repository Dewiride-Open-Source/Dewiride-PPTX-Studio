/**
 * `@pptx-studio/cli` - the Node entry point.
 *
 * The one package in the repository where `node:*` is legal, and the reason the
 * ban holds everywhere else: a CLI has to read paths and write streams, and
 * putting those two capabilities behind a package boundary is what keeps them
 * out of the core.
 *
 * There is almost nothing here. `pptx-studio inspect` reads a file and hands
 * the bytes to `@pptx-studio/census`, which is a browser package - the same
 * code the explorer in a tab runs in its Web Worker. Two front ends, one
 * answer, and no way for them to disagree about what is in a deck.
 */

export { main, type Streams } from './main.js';

export {
  inspectFile,
  runInspect,
  INSPECT_DEFAULTS,
  type InspectOptions,
  type InspectResult,
} from './inspect.js';

export {
  renderDeck,
  runRender,
  DEFAULT_WIDTH,
  type RenderDeckOptions,
  type RenderOptions,
  type RenderResult,
  type RenderedSlide,
} from './render/render.js';

export { indexFonts, systemFontDirectories, type FontLibrary } from './render/faces.js';

export { createFontMeasurer, type FaceUse, type FontMeasurer } from './render/measure.js';

export { facesIn, type Face } from './render/sfnt.js';

export {
  RenderError,
  RENDER_ERROR_CODES,
  isRenderError,
  type RenderErrorCode,
} from './render/errors.js';

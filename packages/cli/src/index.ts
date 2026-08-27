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

/**
 * `@pptx-studio/render-dom` - the live renderer.
 *
 * One `<svg>` per slide, built from the identical node tree
 * `@pptx-studio/render-svg` serialises, so the picture in the editor and the
 * picture in a thumbnail cannot drift apart. Every fact about DrawingML - what
 * a group does to its children, which way round a flipped and rotated shape
 * goes, what `a:grpFill` reaches for - lives there and was measured in
 * experiment C6.
 *
 * Sub-phase 2.10 draws geometry. Text is 3.8, the overlay canvas and the
 * adjust-handle chrome are 2.11, and virtualization is 12.1.
 */

export {
  RenderDomError,
  RENDER_DOM_ERROR_CODES,
  isRenderDomError,
  type RenderDomErrorCode,
} from './errors.js';

export {
  createNode,
  mountSlide,
  renderSlideMarkup,
  type MountOptions,
  type MountedSlide,
} from './mount.js';

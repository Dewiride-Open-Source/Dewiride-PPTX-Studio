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
 * Sub-phase 3.8 adds `mountTextLayer`: an HTML layer over that `<svg>`, built
 * from the identical `TextBlock` the SVG emitter draws, so the two cannot
 * disagree about a line box. Real text nodes, for the selection, the IME and the
 * screen reader an SVG `<text>` cannot give.
 *
 * The overlay canvas and the adjust-handle chrome are 2.11, and virtualization
 * is 12.1.
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

export {
  mergeAdjust,
  mountOverlay,
  slidePoint,
  type HandleEdit,
  type MountedOverlay,
  type OverlayMountOptions,
} from './overlay.js';

export {
  mountTextLayer,
  type LayerBlock,
  type LayerSize,
  type MountedTextLayer,
  type TextLayerOptions,
} from './text/layer.js';

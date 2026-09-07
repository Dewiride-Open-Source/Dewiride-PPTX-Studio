/**
 * `@pptx-studio/render-svg` - slides to SVG, geometry first.
 *
 * Sub-phase 2.10. What this package knows that a naive renderer does not is
 * almost entirely about transforms, and all of it was measured against
 * Microsoft PowerPoint in experiment C6 rather than argued from the standard:
 *
 * - **A flip happens before a rotation.** PowerPoint, asked to mirror a shape
 *   already turned 30 degrees, writes `rot="19800000" flipH="1"` - minus thirty
 *   - which is what a renderer that flips first has to write and what one that
 *   rotates first must not. Four angles, both axes, and again in the positions
 *   of a mirrored group's children.
 * - **A group's children are written in their own coordinate system**, and the
 *   scale factor `ext/chExt` is the part everybody drops. Dropping it is right
 *   on three quarters of the probes here and on every group nobody has resized.
 * - **A `chExt` of zero means no scaling**; an `ext` of zero really is zero; and
 *   a group with no `a:chOff`/`a:chExt` at all uses `chOff = 0`, not `chOff =
 *   off`.
 * - **A rotated child in a non-uniformly scaled group** has no honest answer -
 *   the true image is a parallelogram and `a:xfrm` cannot hold one - and what
 *   PowerPoint does is snap the child's angle to the nearest quadrant and swap
 *   which scale factor reaches which extent. Eighteen angles, and the rounding
 *   at 45 and 135 is the part worth having.
 * - **A group's own fill is never painted.** It exists to be asked for.
 * - **`a:grpFill` is a slice**, and of a rectangle nobody writes down: the
 *   bounding box of everything the group contains, not the group's `off`/`ext`.
 * - **A group does not scale the strokes inside it.** A 4pt outline in a group
 *   stretched to double width is still 4pt.
 * - **`@showMasterSp` hides everything inherited**, not the master's
 *   contribution specifically - on a slide it takes the layout's furniture too.
 *
 * The fixture is `corpus/ground-truth/transforms.json` and it carries the raw
 * cases, so `render.test.ts` re-derives every one of these rather than
 * comparing against a sentence.
 *
 * ## Text, added in 3.8
 *
 * Real `<text>` and `<tspan>`, never `foreignObject`, in a group that is a
 * **sibling** of the shape's own - because `flipH` mirrors the outline and
 * leaves the glyphs alone, while `flipV` turns the whole block half a
 * revolution. Measured 20 of 20 in experiment T8, where counter-flipping both
 * axes fits 10. The baseline sits at the typeface's own ascent-to-descent share
 * of the line box, which the CSS model every browser implements gets wrong on
 * 27 of 36 rows. `corpus/ground-truth/text-rendering.json`.
 *
 * ## What it does not do yet
 *
 * Bullets are resolved and not drawn, four of the seven `@vert` values throw
 * rather than draw approximately, `a:blipFill` paints nothing,
 * `p:graphicFrame` draws nothing, and a compound stroke draws as a single rail.
 */

export { RenderError, RENDER_ERROR_CODES, isRenderError, type RenderErrorCode } from './errors.js';

export {
  MAX_GROUP_DEPTH,
  flatten,
  inheritedSheets,
  layoutSheet,
  layoutSlide,
  type GeometrySource,
  type Placed,
} from './layout.js';

export {
  DEFAULT_UNIT,
  LOCUS_SAMPLES,
  PATH_COLORS,
  shapeOverlay,
  type OverlayGuide,
  type OverlayHandle,
  type OverlayOptions,
  type ShapeOverlay,
} from './overlay.js';

export {
  element,
  escapeXml,
  num,
  serializeSvg,
  text,
  type AttributeValue,
  type SvgElement,
  type SvgNode,
  type SvgText,
} from './node.js';

export {
  Defs,
  effectFilterAttribute,
  fillAttributes,
  strokeAttributes,
  type Attrs,
  type StrokePaint,
} from './paint.js';

export { shapeNodes } from './shape.js';

export {
  approximateRules,
  createTextEngine,
  shapeTextNodes,
  textBlockOf,
  type TextEngine,
  type TextOptions,
} from './text/draw.js';

export { textNodes } from './text/emit.js';

export {
  SHIFT_SIZE_RATIO,
  SMALL_CAPS_RATIO,
  alignOffset,
  capStretches,
  layoutText,
  stretches,
  strutHeight,
  textTurn,
  type LayoutTextOptions,
  type PieceRule,
  type TextBlock,
  type TextLine,
  type TextPiece,
} from './text/layout.js';

export {
  resolveText,
  type ResolvedFrame,
  type ResolvedParagraph,
  type ResolvedRun,
  type ResolvedText,
} from './text/resolve.js';

export {
  renderSlide,
  slideNode,
  type RenderOptions,
  type SlideRender,
  type SlideSize,
} from './slide.js';

export {
  ANGLE_UNITS_PER_DEGREE,
  EMU_PER_POINT,
  IDENTITY_FRAME,
  UNIT_CHILD_SPACE,
  childSpace,
  composeTurn,
  frameOf,
  framePoint,
  frameTransform,
  inverseFramePoint,
  placeChild,
  swapsExtents,
  turnVector,
  unionBox,
  type Box,
  type ChildSpace,
  type Frame,
  type Vec,
} from './transform.js';

export { blipPaint, type MediaImage, type MediaResolver } from './image/blip.js';
export { dataUri, imageSize } from './image/header.js';

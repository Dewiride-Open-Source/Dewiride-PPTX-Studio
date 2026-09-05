import type { ProbeDeck } from '../types.ts';
import { a01Minimal } from './a01-minimal.ts';
import { a02Placeholders } from './a02-placeholders.ts';
import { a03Fills } from './a03-fills.ts';
import { a04Effects } from './a04-effects.ts';
import { a05Geometry } from './a05-geometry.ts';
import { a06Lines } from './a06-lines.ts';
import { a07TextCascade } from './a07-text-cascade.ts';
import { a08Bullets } from './a08-bullets.ts';
import { a09Fields } from './a09-fields.ts';
import { a10RtlCjk } from './a10-rtl-cjk.ts';
import { a11Autofit } from './a11-autofit.ts';
import { a12Masters } from './a12-masters.ts';
import { a13Sections } from './a13-sections.ts';
import { a14Notes } from './a14-notes.ts';
import { a15Comments } from './a15-comments.ts';
import { a16Transitions } from './a16-transitions.ts';
import { a17Animations } from './a17-animations.ts';
import { a18SlideSizes } from './a18-slide-sizes.ts';
import { a19Decorative } from './a19-decorative.ts';
import { a20Tables } from './a20-tables.ts';
import { a21Charts } from './a21-charts.ts';
import { a22ChartEx } from './a22-chartex.ts';
import { a23SmartArt } from './a23-smartart.ts';
import { a24Media } from './a24-media.ts';
import { a25SvgBlips } from './a25-svg-blips.ts';
import { a26Ole } from './a26-ole.ts';
import { a27Ink } from './a27-ink.ts';
import { a29Math } from './a29-math.ts';
import { a30Vml } from './a30-vml.ts';
import { a31EmbeddedFonts } from './a31-embedded-fonts.ts';
import { a32Macros } from './a32-macros.ts';
import { a33Thumbnail } from './a33-thumbnail.ts';
import { a34ExtLst } from './a34-extlst.ts';
import { a35ZipShapes } from './a35-zip-shapes.ts';
import { a36SpacedTags } from './a36-spaced-tags.ts';
import { a37Mce } from './a37-mce.ts';
import { a38Degenerate } from './a38-degenerate.ts';
import { a39LargeIds } from './a39-large-ids.ts';
import { a40Unicode } from './a40-unicode.ts';
import { a41A4 } from './a41-a4.ts';
import { a42CustomSize } from './a42-custom-size.ts';
import { a43KitchenSink } from './a43-kitchen-sink.ts';

/**
 * The Tier A roster, in id order.
 *
 * `tools/corpus/ROSTER.md` is the plan: forty-one decks, what each is the
 * probe for, and which experiment gates the ones that are not yet buildable.
 * This array is what exists, and the two agree: the forty-second slot,
 * `a28-model3d`, is cut rather than pending, and its census key is declared in
 * the manifest's `uncovered` array. A slot named in the roster and missing here
 * would be a deck still to write rather than an omission, and `C-COV` (`C019`)
 * is what refuses to let the difference go unnoticed.
 */
export const PROBE_DECKS: readonly ProbeDeck[] = [
  a01Minimal,
  a02Placeholders,
  a03Fills,
  a04Effects,
  a05Geometry,
  a06Lines,
  a07TextCascade,
  a08Bullets,
  a09Fields,
  a10RtlCjk,
  a11Autofit,
  a12Masters,
  a13Sections,
  a14Notes,
  a15Comments,
  a16Transitions,
  a17Animations,
  a18SlideSizes,
  a19Decorative,
  a20Tables,
  a21Charts,
  a22ChartEx,
  a23SmartArt,
  a24Media,
  a25SvgBlips,
  a26Ole,
  a27Ink,
  a29Math,
  a30Vml,
  a31EmbeddedFonts,
  a32Macros,
  a33Thumbnail,
  a34ExtLst,
  a35ZipShapes,
  a36SpacedTags,
  a37Mce,
  a38Degenerate,
  a39LargeIds,
  a40Unicode,
  a41A4,
  a42CustomSize,
  a43KitchenSink,
];

import type { ProbeDeck } from '../markup/types.ts';
import { a01Minimal } from './basics/a01-minimal.ts';
import { a02Placeholders } from './basics/a02-placeholders.ts';
import { a03Fills } from './basics/a03-fills.ts';
import { a04Effects } from './basics/a04-effects.ts';
import { a05Geometry } from './basics/a05-geometry.ts';
import { a06Lines } from './basics/a06-lines.ts';
import { a07TextCascade } from './text/a07-text-cascade.ts';
import { a08Bullets } from './text/a08-bullets.ts';
import { a09Fields } from './text/a09-fields.ts';
import { a10RtlCjk } from './text/a10-rtl-cjk.ts';
import { a11Autofit } from './text/a11-autofit.ts';
import { a12Masters } from './deck/a12-masters.ts';
import { a13Sections } from './deck/a13-sections.ts';
import { a14Notes } from './deck/a14-notes.ts';
import { a15Comments } from './deck/a15-comments.ts';
import { a16Transitions } from './deck/a16-transitions.ts';
import { a17Animations } from './deck/a17-animations.ts';
import { a18SlideSizes } from './deck/a18-slide-sizes.ts';
import { a19Decorative } from './deck/a19-decorative.ts';
import { a20Tables } from './content/a20-tables.ts';
import { a21Charts } from './content/a21-charts.ts';
import { a22ChartEx } from './content/a22-chartex.ts';
import { a23SmartArt } from './content/a23-smartart.ts';
import { a24Media } from './content/a24-media.ts';
import { a25SvgBlips } from './content/a25-svg-blips.ts';
import { a26Ole } from './content/a26-ole.ts';
import { a27Ink } from './content/a27-ink.ts';
import { a29Math } from './content/a29-math.ts';
import { a30Vml } from './content/a30-vml.ts';
import { a31EmbeddedFonts } from './packaging/a31-embedded-fonts.ts';
import { a32Macros } from './packaging/a32-macros.ts';
import { a33Thumbnail } from './packaging/a33-thumbnail.ts';
import { a34ExtLst } from './packaging/a34-extlst.ts';
import { a35ZipShapes } from './packaging/a35-zip-shapes.ts';
import { a36SpacedTags } from './packaging/a36-spaced-tags.ts';
import { a37Mce } from './packaging/a37-mce.ts';
import { a38Degenerate } from './packaging/a38-degenerate.ts';
import { a39LargeIds } from './packaging/a39-large-ids.ts';
import { a40Unicode } from './packaging/a40-unicode.ts';
import { a41A4 } from './deck/a41-a4.ts';
import { a42CustomSize } from './deck/a42-custom-size.ts';
import { a43KitchenSink } from './packaging/a43-kitchen-sink.ts';
import { a44Transforms } from './basics/a44-transforms.ts';
import { a45Backgrounds } from './deck/a45-backgrounds.ts';

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
  a44Transforms,
  a45Backgrounds,
];

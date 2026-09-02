/**
 * The two images `b09-picture` inserts, written from this repository's own
 * encoders.
 *
 *   node tools/corpus/authored/make-assets.ts <dir>
 *
 * `b09` is the only Tier B deck that needs a file to exist before PowerPoint
 * runs, and this is the whole of why it needs one: `Shapes.AddPicture` takes a
 * path, not bytes. The alternative - reaching for a photograph on this machine,
 * or for stock imagery through the Designer pane - is exactly what the roster's
 * third Tier B constraint forbids, so the image is generated instead.
 *
 * Both files land in the same directory as the decks and are deleted by the
 * build afterwards; neither is committed. What is committed is `b09-picture.pptx`
 * with the bytes embedded in it, which is CC0 for the same reason every Tier A
 * deck is: this repository wrote every pixel.
 *
 * The PNG is deliberately not a flat colour. A crop has to be visible for the
 * deck to be worth opening, and a uniform image makes `a:srcRect` unfalsifiable
 * by eye - four sides cropped off a plain blue square looks identical to no crop
 * at all.
 */

import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { png } from '../gen/png.ts';
import { probeJpeg } from '../gen/jpeg.ts';

const dir = process.argv[2];
if (dir === undefined) {
  throw new Error('usage: node tools/corpus/authored/make-assets.ts <dir>');
}

/**
 * A 160 x 120 target: concentric rings on a diagonal wash, with a one-pixel
 * border and corner blocks. Every one of those reads differently after a crop,
 * which is the point - the border tells you a side was cut, the rings tell you
 * how much, and the corner blocks tell you which side.
 */
const WIDTH = 160;
const HEIGHT = 120;

const pngBytes = png(WIDTH, HEIGHT, (x, y) => {
  const edge = x === 0 || y === 0 || x === WIDTH - 1 || y === HEIGHT - 1;
  if (edge) return 0x1f3864;

  // Corner blocks, 12 x 12, one accent each, so a crop is orientable.
  const left = x < 12;
  const right = x >= WIDTH - 12;
  const top = y < 12;
  const bottom = y >= HEIGHT - 12;
  if (top && left) return 0xc00000;
  if (top && right) return 0x00b050;
  if (bottom && left) return 0xffc000;
  if (bottom && right) return 0x7030a0;

  const dx = x - WIDTH / 2;
  const dy = y - HEIGHT / 2;
  const ring = Math.sqrt(dx * dx + dy * dy);
  if (Math.floor(ring / 9) % 2 === 0) {
    const wash = Math.round(((x / WIDTH) * 0.5 + (y / HEIGHT) * 0.5) * 255);
    return (wash << 16) | (0x88 << 8) | (255 - wash);
  }
  return 0xf2f2f2;
});

/**
 * The JPEG is greyscale blocks, which is all `probeJpeg` writes - it exists so
 * `b09` carries both of the raster formats PowerPoint round-trips without
 * re-encoding, and so the effects on slide 1 sit on a `blipFill` whose blip is
 * not a PNG. A gradient of 8 x 8 blocks with a darker frame.
 */
const jpegBytes = probeJpeg({
  width: 320,
  height: 240,
  density: 96,
  level: (bx, by) => {
    const cols = 40;
    const rows = 30;
    if (bx === 0 || by === 0 || bx === cols - 1 || by === rows - 1) return 32;
    return Math.min(255, 40 + Math.round((bx / cols) * 120 + (by / rows) * 90));
  },
});

const pngPath = join(dir, 'b09-source.png');
const jpegPath = join(dir, 'b09-source.jpeg');
writeFileSync(pngPath, pngBytes);
writeFileSync(jpegPath, jpegBytes);
console.log('wrote ' + pngPath + '  (' + pngBytes.length + ' bytes)');
console.log('wrote ' + jpegPath + '  (' + jpegBytes.length + ' bytes)');

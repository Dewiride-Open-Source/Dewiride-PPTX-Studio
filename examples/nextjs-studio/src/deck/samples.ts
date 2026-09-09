/**
 * The decks shipped with this example.
 *
 * All six are copied from the project's own corpus, all are CC0-1.0 and
 * self-authored, and the `b` ones were authored in real PowerPoint rather than
 * generated - which is why they are the ones worth looking at first.
 */

export interface Sample {
  readonly file: string;
  readonly name: string;
  readonly about: string;
}

export const SAMPLES: readonly Sample[] = [
  {
    file: 'b02-layouts.pptx',
    name: 'Layouts',
    about: 'The eleven built-in layouts, one slide each. The Blank one really is empty.',
  },
  { file: 'b03-text.pptx', name: 'Text', about: 'Placeholders, bullets and the text cascade.' },
  {
    file: 'b09-picture.pptx',
    name: 'Picture',
    about: 'Cropped image fills and a picture placeholder.',
  },
  {
    file: 'a05-geometry.pptx',
    name: 'Geometry',
    about: 'Preset shapes, adjust handles and custom geometry.',
  },
  {
    file: 'a44-transforms.pptx',
    name: 'Transforms',
    about: 'Rotated and flipped shapes, and nested groups.',
  },
  {
    file: 'a45-backgrounds.pptx',
    name: 'Backgrounds',
    about: 'Themed backgrounds through the style matrix.',
  },
];

export const DEFAULT_SAMPLE = SAMPLES[0]!;

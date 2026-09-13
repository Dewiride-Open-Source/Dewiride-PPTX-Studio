import manifest from '../../public/decks/manifest.json';

/** A deck the site serves: a byte-identical copy of a corpus deck, as `public/decks/manifest.json` claims it. */
export interface Sample {
  readonly file: string;
  readonly slides: number;
  readonly bytes: number;
  readonly about: string;
}

export const SAMPLES: readonly Sample[] = manifest.decks.map((deck) => ({
  file: deck.file,
  slides: deck.slides,
  bytes: deck.bytes,
  about: deck.about,
}));

/** The first deck in the manifest is the one every page opens on. */
export const DEFAULT_SAMPLE: Sample = SAMPLES[0]!;

export function sampleNamed(file: string): Sample | undefined {
  return SAMPLES.find((sample) => sample.file === file);
}

/** `a46-hundred-slides.pptx` -> `a46 hundred slides`. */
export function sampleTitle(sample: Sample): string {
  return sample.file.replace(/\.ppt[xm]$/, '').replace(/-/g, ' ');
}

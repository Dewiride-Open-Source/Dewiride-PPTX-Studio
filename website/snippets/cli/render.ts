import { renderDeck } from '@pptx-studio/cli';

declare const bytes: Uint8Array;

//#region example
const { slides, fonts, missing } = renderDeck(bytes, { width: 1280 });
slides[0]?.svg; // stands alone: pictures as data: URIs, defs inline
fonts.filter((face) => face.substituted); // [{ asked: 'Aptos', drawn: 'Carlito', file, substituted: true }]
missing; // code points no indexed face could draw
//#endregion

/** A deck opened for drawing: its slides, its pictures, and one text engine for all of them. */

import { PartStore } from '@pptx-studio/opc';
import { loadDocument, type ListStyle, type Sheet } from '@pptx-studio/model';
import {
  createTextEngine,
  mediaFromStore,
  type MediaResolver,
  type SlideSize,
  type TextEngine,
} from '@pptx-studio/render-svg';

export interface Deck {
  readonly slides: readonly Sheet[];
  readonly size: SlideSize;
  /** The text cascade's seventh source, which lives on `ppt/presentation.xml`. */
  readonly defaultTextStyle: ListStyle | undefined;
  readonly problems: readonly { message: string }[];
  readonly media: MediaResolver;
  /** One measurer and one face-box probe for every mount, not one per slide. */
  readonly text: TextEngine;
}

export function openDeck(bytes: Uint8Array): Deck {
  const store = PartStore.open(bytes);
  const document = loadDocument(store);
  return {
    slides: document.slides,
    size: document.slideSize,
    defaultTextStyle: document.defaultTextStyle,
    problems: document.problems,
    media: mediaFromStore(store),
    text: createTextEngine({ defaultTextStyle: document.defaultTextStyle }),
  };
}

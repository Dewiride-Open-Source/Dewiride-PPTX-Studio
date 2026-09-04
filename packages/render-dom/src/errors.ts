/**
 * Every failure in this package is a `RenderDomError` with a code.
 *
 * There are only two, and both mean the *caller* asked for something
 * impossible: a host that is not an element, or an element with no document to
 * build in. Everything about the slide itself is `render-svg`'s to refuse, and
 * this package never re-decides it.
 */

export type RenderDomErrorCode =
  /** `mountSlide` was handed something that is not an element. */
  | 'RENDER_DOM_NO_HOST'
  /** The host is detached and no `document` was passed. */
  | 'RENDER_DOM_NO_DOCUMENT';

export const RENDER_DOM_ERROR_CODES: readonly RenderDomErrorCode[] = [
  'RENDER_DOM_NO_HOST',
  'RENDER_DOM_NO_DOCUMENT',
];

export class RenderDomError extends Error {
  readonly code: RenderDomErrorCode;

  constructor(code: RenderDomErrorCode, message: string) {
    super(message);
    this.name = 'RenderDomError';
    this.code = code;
  }
}

export function isRenderDomError(value: unknown): value is RenderDomError {
  return value instanceof RenderDomError;
}

export type SiteErrorCode =
  'SITE_PUBLIC_PATH' | 'SITE_RENDERED_MISSING' | 'SITE_REFERENCE_MISSING' | 'SITE_UNKNOWN_DEMO';

/** A failure in the site's own wiring, as opposed to one a package reports. */
export class SiteError extends Error {
  readonly code: SiteErrorCode;

  constructor(code: SiteErrorCode, message: string) {
    super(message);
    this.name = 'SiteError';
    this.code = code;
  }
}

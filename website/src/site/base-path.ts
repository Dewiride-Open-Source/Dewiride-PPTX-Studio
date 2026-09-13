import { SiteError } from './errors';

const BASE_PATH = process.env['NEXT_PUBLIC_BASE_PATH'] ?? '';

/** A file under `public/`, at the path the site is served from. */
export function publicUrl(path: string): string {
  if (!path.startsWith('/')) {
    throw new SiteError('SITE_PUBLIC_PATH', `a public path starts with '/', got '${path}'`);
  }
  return `${BASE_PATH}${path}`;
}

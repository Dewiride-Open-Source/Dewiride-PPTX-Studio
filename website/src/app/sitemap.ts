import type { MetadataRoute } from 'next';

import { source } from '@/docs/source';
import { SITE_URL } from '@/site/navigation';
import { PACKAGES } from '@/site/packages';

export const dynamic = 'force-static';

export default function sitemap(): MetadataRoute.Sitemap {
  const at = (path: string) => ({ url: `${SITE_URL}${path}` });
  return [
    at('/'),
    at('/demos/'),
    ...PACKAGES.map((one) => at(`/demos/${one.name}/`)),
    ...source.getPages().map((page) => at(`${page.url}/`)),
  ];
}

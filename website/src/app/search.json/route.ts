import { createFromSource } from 'fumadocs-core/search/server';

import { source } from '@/docs/source';

// A static export has no server, so the whole index is one JSON file the
// search dialog downloads on first open.
export const dynamic = 'force-static';

export const { staticGET: GET } = createFromSource(source);

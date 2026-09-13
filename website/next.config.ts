import type { NextConfig } from 'next';

// '' locally, in the candidate gate and in the canary; the deploy sets the
// path GitHub Pages serves the repository under and checks it against Pages.
const basePath = process.env['BASE_PATH'] ?? '';
if (basePath !== '' && !/^\/[^/]+$/.test(basePath)) {
  throw new Error(`BASE_PATH must be '' or '/<segment>', got '${basePath}'`);
}

const config: NextConfig = {
  output: 'export',
  basePath,
  trailingSlash: true,
  images: { unoptimized: true },
  // `next/link` is prefixed by Next; a `fetch` of a public file is not, and
  // reads this through `publicUrl`.
  env: { NEXT_PUBLIC_BASE_PATH: basePath },
  // Next writes its own AGENTS.md and CLAUDE.md into this directory on dev
  // start. The repository already has one, at the root, and it is not this.
  agentRules: false,
};

export default config;

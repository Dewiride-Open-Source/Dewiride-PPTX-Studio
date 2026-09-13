import type { NextConfig } from 'next';

/**
 * `@pptx-studio/cli` reads font files off the disk, so it must be required at
 * run time rather than bundled into the server chunk.
 */
const config: NextConfig = {
  serverExternalPackages: ['@pptx-studio/cli'],
  // Next writes its own AGENTS.md and CLAUDE.md into this directory on dev
  // start. The repository already has one, at the root, and it is not this.
  agentRules: false,
};

export default config;

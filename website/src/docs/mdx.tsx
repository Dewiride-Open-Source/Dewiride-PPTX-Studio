import defaultMdxComponents from 'fumadocs-ui/mdx';
import type { MDXComponents } from 'mdx/types';

import { Roadmap } from '@/status/roadmap';

/** Fumadocs' defaults plus the site's own MDX components, by the names the pages use. */
export function getMDXComponents(components?: MDXComponents): MDXComponents {
  return { ...defaultMdxComponents, Roadmap, ...components };
}

import defaultMdxComponents from 'fumadocs-ui/mdx';
import type { MDXComponents } from 'mdx/types';

import { Demo } from '@/demos/embed';
import { ErrorCodes } from '@/reference/errors';
import { ApiTable } from '@/reference/table';
import { Roadmap } from '@/status/roadmap';
import { PackageHeader } from './package-header';

/** Fumadocs' defaults plus the site's own MDX components, by the names the pages use. */
export function getMDXComponents(components?: MDXComponents): MDXComponents {
  return {
    ...defaultMdxComponents,
    ApiTable,
    Demo,
    ErrorCodes,
    PackageHeader,
    Roadmap,
    ...components,
  };
}

'use client';

import dynamic from 'next/dynamic';
import type { ComponentType } from 'react';

import { Skeleton } from '@/design/state';
import type { PackageName } from '@/site/packages';
import type { DemoProps } from './registry';

const loading = () => <Skeleton className="w-full" />;

/**
 * One lazily loaded component per demo, made once at module scope.
 *
 * `ssr: false` because every demo reads a canvas, a Worker or the DOM, none of
 * which exist at build time; a chunk loads only when its demo is on screen.
 */
export const LOADED: Readonly<Record<PackageName, ComponentType<DemoProps>>> = {
  'render-dom': dynamic(() => import('./viewer/demo'), { ssr: false, loading }),
  'render-svg': dynamic(() => import('./editor/demo'), { ssr: false, loading }),
  model: dynamic(() => import('./inheritance/demo'), { ssr: false, loading }),
  opc: dynamic(() => import('./package/demo'), { ssr: false, loading }),
  xml: dynamic(() => import('./markup/demo'), { ssr: false, loading }),
  census: dynamic(() => import('./census/demo'), { ssr: false, loading }),
  validate: dynamic(() => import('./validate/demo'), { ssr: false, loading }),
  writer: dynamic(() => import('./export/demo'), { ssr: false, loading }),
  geometry: dynamic(() => import('./geometry/demo'), { ssr: false, loading }),
  paint: dynamic(() => import('./paint/demo'), { ssr: false, loading }),
  text: dynamic(() => import('./text/demo'), { ssr: false, loading }),
  cli: dynamic(() => import('./cli/demo'), { ssr: false, loading }),
};

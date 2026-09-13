'use client';

import { Demo } from '@/demos/embed';
import { Tabs } from '@/design/tabs';
import type { PackageName } from '@/site/packages';

const TABS: readonly { value: PackageName; label: string }[] = [
  { value: 'model', label: 'Inheritance' },
  { value: 'render-svg', label: 'Edit' },
  { value: 'census', label: 'Census' },
  { value: 'validate', label: 'Validate' },
  { value: 'writer', label: 'Export' },
  { value: 'opc', label: 'Package' },
  { value: 'xml', label: 'Markup' },
  { value: 'text', label: 'Text and fonts' },
];

/** Every capability on one deck: the live viewer on top, the rest of the packages beneath. */
export function Bench() {
  return (
    <div className="flex flex-col gap-6">
      <Demo name="render-dom" full />
      <Tabs label="Tools" tabs={TABS} initial="model">
        {(active) => <Demo name={active} full />}
      </Tabs>
    </div>
  );
}

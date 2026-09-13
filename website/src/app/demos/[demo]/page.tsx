import type { Metadata } from 'next';
import { notFound } from 'next/navigation';

import { Demo, DemoHeader } from '@/demos/embed';
import { DEMOS, demoNamed } from '@/demos/registry';

interface Props {
  readonly params: Promise<{ demo: string }>;
}

export default async function DemoRoute({ params }: Props) {
  const { demo } = await params;
  const entry = demoNamed(demo);
  if (entry === undefined) notFound();
  return (
    <div className="flex flex-col gap-4">
      <DemoHeader entry={entry} />
      <Demo name={entry.name} full />
    </div>
  );
}

export function generateStaticParams() {
  return DEMOS.map((demo) => ({ demo: demo.name }));
}

export async function generateMetadata({ params }: Props): Promise<Metadata> {
  const { demo } = await params;
  const entry = demoNamed(demo);
  if (entry === undefined) notFound();
  return { title: `${entry.title} · @pptx-studio/${entry.name}`, description: entry.blurb };
}

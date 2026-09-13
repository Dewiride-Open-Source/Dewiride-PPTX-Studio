import { HomeLayout } from 'fumadocs-ui/layouts/home';

import { Hero } from '@/landing/hero';
import { Honesty } from '@/landing/honesty';
import { Install } from '@/landing/install';
import { Measured } from '@/landing/measured';
import { PackageGrid } from '@/landing/packages';
import { Principles } from '@/landing/principles';
import { Showcase } from '@/landing/showcase';
import { Footer } from '@/shell/footer';
import { baseOptions } from '@/site/navigation';
import { StatusStrip } from '@/status/strip';

export default function Landing() {
  return (
    <HomeLayout {...baseOptions()}>
      <main id="main">
        <Hero />
        <Showcase />
        <Principles />
        <PackageGrid />
        <Install />
        <StatusStrip />
        <Measured />
        <Honesty />
      </main>
      <Footer />
    </HomeLayout>
  );
}

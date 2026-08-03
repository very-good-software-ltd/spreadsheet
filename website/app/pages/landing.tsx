import { Benchmarks } from "~/components/benchmarks";
import { Features } from "~/components/features";
import { Hero } from "~/components/hero";
import { SiteFooter } from "~/components/site-footer";
import { Usage } from "~/components/usage";
import { Demo } from "~/pages/viewer";

export function LandingPage() {
  return (
    <div className="min-h-screen bg-white text-gray-900 dark:bg-gray-950 dark:text-gray-100">
      <Hero />
      <Features />
      <section id="demo" className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
        <div className="mx-auto max-w-6xl">
          <h2 className="text-2xl font-bold">Try it</h2>
          <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
            Pick an .xlsx and it streams into the table below, entirely in your browser. Large files are fine. The table
            shows the first 100,000 rows while it counts every one.
          </p>
          <div className="mt-8">
            <Demo />
          </div>
        </div>
      </section>
      <Benchmarks />
      <Usage />
      <SiteFooter />
    </div>
  );
}

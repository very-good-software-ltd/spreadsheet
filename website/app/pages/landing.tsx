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
      <section className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
        <div className="mx-auto max-w-3xl">
          <h2 className="text-2xl font-bold">Why this exists</h2>
          <p className="mt-4 text-gray-600 dark:text-gray-300">
            <code className="font-mono">xlsx</code> (SheetJS) and <code className="font-mono">exceljs</code> have read
            and written spreadsheets in JavaScript for years, and plenty of code depends on them. If you need a format
            we don't handle, reach for them. This isn't about replacing them.
          </p>
          <p className="mt-4 text-gray-600 dark:text-gray-300">
            But they were written a long time ago, and it shows. Back then Node had no{" "}
            <code className="font-mono">ReadableStream</code> and no{" "}
            <code className="font-mono">DecompressionStream</code>, and ESM and TypeScript were nowhere near where they
            are now. So they carry their own stream code, a big dependency tree and a CommonJS build, and they can't
            drop any of it without breaking everyone already on them.
          </p>
          <p className="mt-4 text-gray-600 dark:text-gray-300">
            So we go the other way and start fresh, on the platform we actually have now. Streaming and decompression
            are built in, ESM is settled and TypeScript is everywhere. Everything this library does well comes from that
            fresh start.
          </p>
          <p className="mt-4 text-gray-600 dark:text-gray-300">
            It's early. We read <code className="font-mono">.xlsx</code> and <code className="font-mono">.ods</code> but
            only write <code className="font-mono">.xlsx</code>, nothing more exotic. Some templates we'd rather refuse
            than get wrong, so a pivot table reading from your data stops the save and tells you so. And we haven't seen the years of odd real-world files the older libraries have. If that's
            fine for what you're building, give it a try.
          </p>
        </div>
      </section>
      <Features />
      <section id="demo" className="border-t border-gray-100 px-6 py-16 dark:border-gray-900">
        <div className="mx-auto max-w-6xl">
          <h2 className="text-2xl font-bold">Try it</h2>
          <p className="mt-2 text-sm text-gray-600 dark:text-gray-400">
            Pick an .xlsx or .ods and it streams into the table below, entirely in your browser. Large files are fine.
            The table shows the first 100,000 rows while it counts every one.
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

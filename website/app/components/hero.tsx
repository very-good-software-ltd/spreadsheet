import { InstallCommand } from "~/components/install-command";
import { REPO_URL } from "~/site";

export function Hero() {
  return (
    <section className="px-6 pt-20 pb-16 text-center">
      <div className="mx-auto max-w-3xl">
        <h1 className="font-mono text-3xl font-bold tracking-tight sm:text-4xl">very-good-spreadsheet</h1>
        <p className="mt-5 text-lg text-gray-600 dark:text-gray-300">
          Read Excel <code className="font-mono">.xlsx</code> files in Node and the browser, without holding the whole
          sheet in memory.
        </p>
        <div className="mt-8 flex flex-col items-center gap-4">
          <InstallCommand />
          <div className="flex items-center gap-3">
            <a
              href={REPO_URL}
              target="_blank"
              rel="noreferrer"
              className="rounded-md bg-gray-900 px-4 py-2 text-sm font-medium text-white hover:bg-gray-700 dark:bg-gray-100 dark:text-gray-900 dark:hover:bg-gray-300"
            >
              GitHub
            </a>
            <a
              href="#demo"
              className="rounded-md border border-gray-300 px-4 py-2 text-sm font-medium hover:bg-gray-50 dark:border-gray-700 dark:hover:bg-gray-800"
            >
              Try the demo
            </a>
          </div>
          <div className="mt-2 flex flex-wrap items-center justify-center gap-x-5 gap-y-1 text-xs text-gray-500">
            <span>TypeScript</span>
            <span>ESM</span>
            <span>Node 24+</span>
          </div>
        </div>
      </div>
    </section>
  );
}

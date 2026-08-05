import { NPM_URL, REPO_URL } from "~/site";

export function SiteFooter() {
  return (
    <footer className="border-t border-gray-200 px-6 py-10 text-sm text-gray-500 dark:border-gray-800">
      <div className="mx-auto flex max-w-5xl flex-wrap items-center justify-between gap-4">
        <span>MIT licensed.</span>
        <div className="flex items-center gap-4">
          <a href={NPM_URL} target="_blank" rel="noreferrer" className="hover:text-gray-800 dark:hover:text-gray-200">
            npm
          </a>
          <a href={REPO_URL} target="_blank" rel="noreferrer" className="hover:text-gray-800 dark:hover:text-gray-200">
            GitHub
          </a>
        </div>
      </div>
    </footer>
  );
}

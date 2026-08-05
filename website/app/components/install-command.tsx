import { useState } from "react";
import { NPM_INSTALL } from "~/site";

const lastSpace = NPM_INSTALL.lastIndexOf(" ");
const command = NPM_INSTALL.slice(0, lastSpace);
const packageName = NPM_INSTALL.slice(lastSpace + 1);

export function InstallCommand() {
  const [copied, setCopied] = useState(false);

  function copy() {
    void navigator.clipboard.writeText(NPM_INSTALL).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    });
  }

  return (
    <div className="inline-flex max-w-full items-center gap-3 rounded-lg border border-gray-200 bg-gray-50 px-4 py-2 font-mono text-xs sm:text-sm dark:border-gray-700 dark:bg-gray-900">
      <span className="shrink-0 text-gray-400">$</span>
      <span>
        <span className="whitespace-nowrap">{command}</span> <span className="whitespace-nowrap">{packageName}</span>
      </span>
      <button
        type="button"
        onClick={copy}
        className="shrink-0 text-xs text-gray-400 hover:text-gray-700 dark:hover:text-gray-200"
      >
        {copied ? "copied" : "copy"}
      </button>
    </div>
  );
}

import { useState } from "react";
import { NPM_INSTALL } from "~/site";

export function InstallCommand() {
  const [copied, setCopied] = useState(false);

  function copy() {
    void navigator.clipboard.writeText(NPM_INSTALL).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    });
  }

  return (
    <div className="inline-flex items-center gap-3 rounded-lg border border-gray-200 bg-gray-50 px-4 py-2 font-mono text-sm dark:border-gray-700 dark:bg-gray-900">
      <span className="text-gray-400">$</span>
      <span>{NPM_INSTALL}</span>
      <button
        type="button"
        onClick={copy}
        className="text-xs text-gray-400 hover:text-gray-700 dark:hover:text-gray-200"
      >
        {copied ? "copied" : "copy"}
      </button>
    </div>
  );
}

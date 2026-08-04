import ShikiHighlighter, { createHighlighterCore, createJavaScriptRegexEngine } from "react-shiki/core";

const highlighter = await createHighlighterCore({
  themes: [import("@shikijs/themes/github-light"), import("@shikijs/themes/github-dark")],
  langs: [import("@shikijs/langs/typescript")],
  engine: createJavaScriptRegexEngine(),
});

export function SyntaxHighlight({ children }: { children: string }) {
  return (
    <ShikiHighlighter
      highlighter={highlighter}
      language="typescript"
      theme={{ light: "github-light", dark: "github-dark" }}
      showLanguage={false}
    >
      {children}
    </ShikiHighlighter>
  );
}

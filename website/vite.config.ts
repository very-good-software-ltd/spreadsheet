import { reactRouter } from "@react-router/dev/vite";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "vitest/config";

export default defineConfig({
  environments: {
    ssr: {
      build: {
        rolldownOptions: {
          input: "./server/app.ts",
        },
      },
    },
  },
  plugins: [
    tailwindcss(),
    // React Router doesn't play nicely with vitest, so we skip it when running tests.
    // https://github.com/remix-run/react-router/discussions/12655#discussioncomment-11720266
    !process.env.VITEST && reactRouter(),
  ],
  resolve: {
    tsconfigPaths: true,
  },
  test: {
    environment: "happy-dom",
    include: ["./tests/**/*.test.{js,jsx,ts,tsx}"],
    setupFiles: ["./tests/setup.ts"],
    reporters: process.env.CI ? "dot" : "default",
    watch: false,
  },
});

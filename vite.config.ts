import { defineConfig } from "vitest/config";

export default defineConfig({
  build: {
    lib: {
      entry: "src/index.ts",
      formats: ["es"],
      fileName: "index",
    },
    rolldownOptions: {
      external: ["fflate", "saxes", /^node:/],
    },
  },
  test: {
    include: ["test/**/*.test.ts"],
    watch: false,
    environment: "node",
  },
});

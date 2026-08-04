import dts from "unplugin-dts/vite";
import { defineConfig } from "vitest/config";

export default defineConfig(({ command }) => ({
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
  // No need to use this plugin when running tests
  plugins: command === "build" ? [dts({ include: ["src"] })] : [],
  test: {
    include: ["test/**/*.test.ts"],
    watch: false,
    environment: "node",
  },
}));

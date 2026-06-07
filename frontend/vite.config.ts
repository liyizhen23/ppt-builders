import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const certsDir = join(homedir(), ".office-addin-dev-certs");

export default defineConfig({
  plugins: [react()],
  server: {
    host: "localhost",
    port: 5173,
    strictPort: true,
    https: {
      cert: readFileSync(join(certsDir, "localhost.crt")),
      key: readFileSync(join(certsDir, "localhost.key"))
    },
    proxy: {
      "/api": "http://127.0.0.1:3000"
    }
  },
  build: {
    outDir: "dist",
    sourcemap: true
  }
});

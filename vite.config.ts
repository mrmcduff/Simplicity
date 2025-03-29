import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import mkcert from "vite-plugin-mkcert";
import { resolve } from "path";

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    mkcert(), // Automatically creates and uses a valid HTTPS certificate
  ],
  server: {
    port: 3000,
    https: true,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "src/taskpane/taskpane.html"),
        functions: resolve(__dirname, "src/functions/function-file.html"),
      },
    },
  },
  base: "/",
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
  publicDir: "assets",
});

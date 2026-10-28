import react from "@vitejs/plugin-react";
import fs from "fs";
import path from "path";
import { defineConfig, loadEnv } from "vite";

export default defineConfig(({ mode }) => {
  const dreamhost = mode === "dreamhost";
  const envDir = path.resolve(__dirname, dreamhost ? "config/dreamhost" : ".");
  const apiUrl = dreamhost ? loadEnv(mode, envDir, "VITE_SIGNALTUNER_API_URL").VITE_SIGNALTUNER_API_URL : undefined;
  if (dreamhost && apiUrl !== "https://signaltuner-back.azurewebsites.net") {
    throw new Error("DreamHost builds require the approved HTTPS Azure backend; remove conflicting shell/environment overrides.");
  }
  return {
    plugins: [react(), ...(dreamhost ? [{
      name: "dreamhost-static-html",
      transformIndexHtml: html => html.replace(/<!--app-(?:head|html)-->/g, ""),
    }] : [])],
    base: dreamhost ? "/" : "/tabs/home",
    envDir,
    // Only the explicitly approved public API URL may enter the DreamHost bundle.
    ...(dreamhost ? {
      envPrefix: [],
      define: { "import.meta.env.VITE_SIGNALTUNER_API_URL": JSON.stringify(apiUrl) },
      publicDir: false,
      build: { outDir: "dist/dreamhost", sourcemap: false, emptyOutDir: true },
    } : {}),
    esbuild: { tsconfigRaw: fs.readFileSync(path.join(__dirname, "tsconfig.app.json"), "utf8") },
  };
});

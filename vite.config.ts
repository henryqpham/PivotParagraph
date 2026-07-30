import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";
import { fileURLToPath } from "node:url";
import { appVersion } from "./appVersion";

/**
 * SINGLE-FILE build config (`npm run build:single`) — a SECOND build target that
 * sits alongside Next.js, which remains the dev/deploy path and is untouched by
 * this file. It exists because the app is 100% client-side (no API routes, no
 * server components, no data fetching), so the same `components/` + `lib/` code
 * can be bundled into ONE self-contained .html that runs from a file:// double-
 * click — shareable over a network drive with no hosting or install.
 *
 * `viteSingleFile` inlines every JS chunk and stylesheet into the HTML, so the
 * output has ZERO external requests (which is also what makes it work under the
 * file:// protocol, where module/CORS fetches are blocked).
 *
 * Tailwind needs no plugin here: Vite auto-loads the repo's existing
 * `postcss.config.mjs`, which already runs `@tailwindcss/postcss`. One CSS
 * pipeline, shared by both builds.
 */
/**
 * Strip `type="module"` (and the pointless `crossorigin`) off the inlined entry
 * script. Vite always tags its entry as a module, but module scripts are subject
 * to CORS rules that an opaque `file://` origin cannot satisfy — the surest way
 * for a double-clicked page to render blank. Safe here ONLY because the bundle is
 * emitted as an IIFE (see `rollupOptions.output.format`), so it contains no
 * import/export/import.meta and behaves identically as a classic script.
 * Runs after `viteSingleFile`, which is what puts the code inline in the first place.
 */
function classicScriptForFileProtocol() {
  return {
    name: "classic-script-for-file-protocol",
    enforce: "post" as const,
    generateBundle(_options: unknown, bundle: Record<string, unknown>) {
      for (const asset of Object.values(bundle)) {
        const a = asset as { fileName?: string; source?: unknown };
        if (!a.fileName?.endsWith(".html") || typeof a.source !== "string") {
          continue;
        }
        a.source = a.source
          .replace(/<script type="module" crossorigin>/g, "<script>")
          .replace(/<script type="module">/g, "<script>")
          // viteSingleFile turns the stylesheet <link> into a <style> but leaves
          // the now-meaningless link attributes behind.
          .replace(/<style rel="stylesheet" crossorigin>/g, "<style>");
      }
    },
  };
}

export default defineConfig({
  // Same build-time version constant the Next config exposes — the component
  // reads `process.env.NEXT_PUBLIC_APP_VERSION`, and this `define` replaces
  // that exact expression with the literal string in the bundle (Vite has no
  // runtime `process`). One badge, two builders, one source (appVersion.ts).
  define: {
    "process.env.NEXT_PUBLIC_APP_VERSION": JSON.stringify(appVersion()),
  },
  plugins: [react(), viteSingleFile(), classicScriptForFileProtocol()],
  resolve: {
    // Mirrors tsconfig's `@/*` -> `./*` path alias so the shared modules import
    // identically under both builders.
    alias: {
      "@": fileURLToPath(new URL(".", import.meta.url)),
    },
  },
  build: {
    outDir: "dist-single",
    emptyOutDir: true,
    // Force everything into one chunk; viteSingleFile then inlines it. Raising
    // the warning limit keeps the log clean — a ~1MB bundle is expected and fine
    // here, since the file loads from local disk with no network round-trip.
    assetsInlineLimit: 100_000_000,
    cssCodeSplit: false,
    chunkSizeWarningLimit: 4096,
    rollupOptions: {
      // IIFE, not ESM. This is the load-bearing choice for file:// support:
      // browsers apply module-script CORS rules to `<script type="module">`, and
      // an opaque file:// origin fails them — the page would silently render
      // blank on a double-click. A classic script has no such restriction.
      output: { format: "iife", inlineDynamicImports: true },
    },
  },
});

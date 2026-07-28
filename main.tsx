// Client entry for the SINGLE-FILE build only (`npm run build:single`).
//
// This is the Vite counterpart to Next's `app/layout.tsx` + `app/page.tsx`: it
// mounts the exact same `PasteInput` shell into `index.html`'s #root and pulls in
// the same `app/globals.css` design tokens, so the standalone .html and the
// Next.js app render identical UI from one source of truth.
//
// Both entries load the same system-font stack from globals.css (no webfont), so
// the standalone .html and the Next.js app are visually identical.

import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PasteInput } from "@/components/PasteInput";
import "@/app/globals.css";

function mount() {
  const container = document.getElementById("root");
  if (!container) throw new Error("Missing #root container in index.html");
  createRoot(container).render(
    <StrictMode>
      <main className="flex min-h-0 flex-1 flex-col">
        <PasteInput />
      </main>
    </StrictMode>,
  );
}

// Wait for the DOM when it isn't parsed yet. This is load-bearing for the
// single-file build: Vite hoists the entry script into <head>, and because we
// emit it as a CLASSIC script (module scripts are unusable from file://) it runs
// IMMEDIATELY — before <body>, so #root would not exist yet and the page would
// render blank. Module scripts get this deferral for free; classic ones don't.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", mount, { once: true });
} else {
  mount();
}

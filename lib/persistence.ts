// Session persistence: save the whole workspace (pasted tables + shared styling)
// to the browser's localStorage so a refresh, tab close, or crash doesn't discard
// the user's work. Everything is client-side and JSON-serializable; NOTHING is
// uploaded — this is local-only, matching the app's no-backend promise.
//
// The helpers are intentionally GENERIC (they store/return `unknown`) so this
// low-level module keeps no dependency on the component-level state shapes; the
// owner (PasteInput) defines the snapshot type and validates the loaded blob.
// The stored envelope is versioned: a bump silently invalidates old data instead
// of hydrating a stale/incompatible shape into React state.

const KEY = "pivotparagraph:session";
/**
 * The pre-rename storage key. Read (once, on load) so an existing workspace
 * survives the rename instead of silently vanishing; never written to again, and
 * cleared alongside the current key so "Clear all" leaves nothing behind.
 */
const LEGACY_KEY = "excel-word-sections:session";
const VERSION = 1;

type Envelope = { version: number; data: unknown };

/**
 * Persist a snapshot under the versioned key. No-ops when localStorage is
 * unavailable (SSR, private-mode quota, storage disabled) — persistence is a
 * best-effort convenience, never load-bearing.
 */
export function saveSession(data: unknown): void {
  if (typeof window === "undefined") return;
  try {
    const envelope: Envelope = { version: VERSION, data };
    window.localStorage.setItem(KEY, JSON.stringify(envelope));
  } catch {
    // Quota exceeded or storage blocked — drop silently.
  }
}

/**
 * Read the saved snapshot, or `null` when there is none, the JSON is corrupt, or
 * the stored version doesn't match this build. The caller validates the shape of
 * the returned `data` before trusting it.
 */
export function loadSession(): unknown {
  if (typeof window === "undefined") return null;
  try {
    // Fall back to the pre-rename key so a workspace saved under the old name
    // still loads; the next save writes it back under the current key.
    const raw =
      window.localStorage.getItem(KEY) ??
      window.localStorage.getItem(LEGACY_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Envelope | null;
    if (!parsed || parsed.version !== VERSION) return null;
    return parsed.data;
  } catch {
    return null;
  }
}

/** Remove the saved snapshot entirely. */
export function clearSession(): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.removeItem(KEY);
    window.localStorage.removeItem(LEGACY_KEY);
  } catch {
    // ignore
  }
}

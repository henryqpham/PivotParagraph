import { execSync } from "node:child_process";

/**
 * The app's display version, derived from git at BUILD time — imported by BOTH
 * build configs (`next.config.ts` and `vite.config.ts`), which run in Node, so
 * the version is computed once per build and inlined into the client bundle as
 * a plain string (no git or Node needed at runtime — the single-file build
 * stays self-contained).
 *
 * Scheme: `1.<commit count>` (+ the short hash for the tooltip) — the industry-
 * standard-looking "v1.58" the user asked for, with zero manual bumping: every
 * commit advances it. Outside a git checkout (a bare source download) it falls
 * back to "dev" rather than failing the build.
 */
export function appVersion(): string {
  try {
    const count = execSync("git rev-list --count HEAD", {
      stdio: ["ignore", "pipe", "ignore"],
    })
      .toString()
      .trim();
    const hash = execSync("git rev-parse --short HEAD", {
      stdio: ["ignore", "pipe", "ignore"],
    })
      .toString()
      .trim();
    return `1.${count} · ${hash}`;
  } catch {
    return "dev";
  }
}

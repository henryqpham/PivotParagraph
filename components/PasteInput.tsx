"use client";

import {
  useMemo,
  useRef,
  useState,
  type ClipboardEvent,
  type ReactNode,
} from "react";
import { parseClipboard } from "@/lib/parser";
import {
  buildWordHtml,
  htmlToPlainText,
  isHeadingStyleSet,
} from "@/lib/clipboard";
import type { HeadingStyle, LevelStyle } from "@/lib/clipboard";
import { DEFAULT_NUMBERING } from "@/lib/renderers";
import { tableToHtml, type TableState } from "./tableModel";
import { TableCard } from "./TableCard";
import { RenderedPreview } from "./RenderedPreview";
import { JsonPreview } from "./JsonPreview";

const HEADING_FONTS = [
  "Calibri Light",
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Cambria",
];

const MAX_TABLES = 100;

// One row of the per-level editor (size kept as a string so it can be cleared).
type LevelInput = { color: string; font: string; sizeInput: string; bold: boolean };
// Default look for an untouched level: plain Arial 11 black (matches a document
// body, not a blue Word Heading). All levels start identical.
const DEFAULT_LEVEL: LevelInput = {
  color: "#000000",
  font: "Arial",
  sizeInput: "11",
  bold: false,
};

/** A tiny uppercase pill that labels which scope a control panel affects. */
function ScopeBadge({ children }: { children: ReactNode }) {
  return (
    <span className="rounded border border-foreground/15 bg-foreground/[0.04] px-1.5 py-0.5 text-[10px] font-medium uppercase tracking-wide text-muted">
      {children}
    </span>
  );
}

export function PasteInput() {
  // Each paste appends one table; tables stack down the page in paste order.
  const [tables, setTables] = useState<TableState[]>([]);
  const [error, setError] = useState<string | null>(null);
  // Polite screen-reader announcement of the last paste result (a visible error
  // is announced via role="alert"; a success has no visible banner, so it is
  // mirrored here for non-sighted users).
  const [status, setStatus] = useState<string>("");
  // Which table's tab is open; only the active table is rendered (no scrolling
  // through the rest). New pastes become active; removing the active tab falls
  // to a neighbor.
  const [activeId, setActiveId] = useState<string | null>(null);
  const [copyAllState, setCopyAllState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  // Per-table "Copy for Word" state (the button lives in the preview pane header).
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  // Right-pane preview tab: the rendered outline or the raw parsed JSON.
  const [view, setView] = useState<"rendered" | "json">("rendered");
  // Whether the global "Level styles" panel (top band) is expanded.
  const [showStyles, setShowStyles] = useState(false);
  // Monotonic id source: never reused, so React keys stay stable when a middle
  // table is removed (a ref, so bumping it never triggers a render).
  const idRef = useRef(0);

  // Body font (default Arial) so unstyled body text doesn't fall back to Times
  // New Roman on a Word paste. Separate from the per-level row look.
  const [bodyFont, setBodyFont] = useState<string>("Arial");
  // The single heading-style source, shared across every table: per-level look
  // (Level 1 = the pivot title, Levels 2-9 = the nested rows by depth). SPARSE --
  // an index is written only when that level is edited; untouched levels use
  // DEFAULT_LEVEL (so they all start identical; "Reset levels" clears this back
  // to []). Sticky, even across "Clear all".
  const [levelStyles, setLevelStyles] = useState<LevelInput[]>([]);
  // Left-indent added per nesting level (inches). Held as a string so it can be
  // cleared/backspaced; clamped to [0, 2] for output.
  const [indentInput, setIndentInput] = useState<string>("0.2");
  // Optional Word style names: when set, the title / body paragraphs carry
  // `mso-style-name` so a "Use Destination Styles" paste adopts the destination
  // document's styles. Blank → the app's direct per-level look.
  const [headingStyleName, setHeadingStyleName] = useState<string>("Heading 1");
  const headingStyle = useMemo<HeadingStyle>(() => {
    const clampPt = (s: string, fallback: number) => {
      const n = parseInt(s, 10);
      return Number.isFinite(n) && n >= 1 && n <= 72 ? n : fallback;
    };
    // Full 9-entry per-level look: an edited level wins; otherwise DEFAULT_LEVEL.
    const levels: LevelStyle[] = Array.from({ length: 9 }, (_, i) => {
      const ls = levelStyles[i];
      return {
        color: ls?.color ?? DEFAULT_LEVEL.color,
        font: ls?.font ?? DEFAULT_LEVEL.font,
        size: clampPt(
          ls?.sizeInput ?? DEFAULT_LEVEL.sizeInput,
          parseInt(DEFAULT_LEVEL.sizeInput, 10),
        ),
        bold: ls?.bold ?? DEFAULT_LEVEL.bold,
      };
    });
    const parsedIndent = parseFloat(indentInput);
    const indentStep = Number.isFinite(parsedIndent)
      ? Math.min(2, Math.max(0, parsedIndent))
      : 0.2;
    return { levels, indentStep, headingStyleName };
  }, [levelStyles, indentInput, headingStyleName]);

  // Edit one level: snapshot DEFAULT_LEVEL into that index if it's not set yet,
  // then apply the patch. Keeps the array otherwise sparse.
  function setLevel(i: number, patch: Partial<LevelInput>) {
    setLevelStyles((prev) => {
      const next = [...prev];
      next[i] = { ...(next[i] ?? DEFAULT_LEVEL), ...patch };
      return next;
    });
  }

  function handlePaste(e: ClipboardEvent<HTMLDivElement>) {
    e.preventDefault();
    if (tables.length >= MAX_TABLES) {
      setError(
        `Reached the ${MAX_TABLES}-table limit. Remove a table to add another.`,
      );
      return;
    }
    try {
      const rows = parseClipboard(e.clipboardData);
      if (rows.length === 0) {
        setError(
          "Couldn't read a table from the clipboard. Copy a cell range from Excel or Google Sheets, then paste here.",
        );
        return;
      }
      const id = `t${++idRef.current}`;
      const next: TableState = {
        id,
        grid: rows,
        pivotLevels: [], // nothing placed yet; user builds the outline
        markers: [], // sparse -> default 1./a./i. cycle until the user picks
        fieldLabels: {}, // per-col label look; default (shown, plain) until edited
        sortDirs: {}, // per-col sort; absent col -> off (first-seen order)
        breakAfter: [], // per-level "blank line after"; default off
        numbering: DEFAULT_NUMBERING, // app-drawn multilevel numbers; off by default
        headingLevels: [], // per-level "make a Word heading"; default none
        sectionTitle: "",
      };
      setTables((prev) => [...prev, next]);
      setActiveId(id); // open the just-pasted table
      setError(null);
      setStatus(`Added table ${tables.length + 1}.`);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  function patchTable(id: string, patch: Partial<TableState>) {
    setTables((ts) => ts.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  }

  function removeTable(id: string) {
    // If the open tab is removed, fall to the next tab (or the previous one if
    // it was the last). Computed from the pre-removal order.
    setActiveId((curr) => {
      if (curr !== id) return curr;
      const idx = tables.findIndex((t) => t.id === id);
      const remaining = tables.filter((t) => t.id !== id);
      if (remaining.length === 0) return null;
      return remaining[Math.min(idx, remaining.length - 1)].id;
    });
    setTables((ts) => ts.filter((t) => t.id !== id));
  }

  function clearAll() {
    setTables([]);
    setActiveId(null);
    setError(null);
  }

  // Combined export: concatenate every table's fragment and wrap ONCE. Valid
  // because buildWordHtml's rewrites are global and it emits a single @page rule,
  // so the whole stack becomes one Word doc (tables in paste order). Built
  // synchronously before any await to keep the user gesture.
  async function copyAll() {
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyAllState("error");
      setTimeout(() => setCopyAllState("idle"), 2000);
      return;
    }
    const titleIsHeading = isHeadingStyleSet(headingStyleName);
    const combined = tables
      .map((t) => tableToHtml(t, titleIsHeading))
      .join("\n");
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([buildWordHtml(combined, headingStyle, bodyFont)], {
          type: "text/html",
        }),
        "text/plain": new Blob([htmlToPlainText(combined)], {
          type: "text/plain",
        }),
      });
      await navigator.clipboard.write([item]);
      setCopyAllState("copied");
    } catch {
      setCopyAllState("error");
    }
    setTimeout(() => setCopyAllState("idle"), 2000);
  }

  const atLimit = tables.length >= MAX_TABLES;
  // The open tab; fall back to the first table if the id ever goes stale.
  const activeTable = tables.find((t) => t.id === activeId) ?? tables[0] ?? null;

  // Whether the title is a real Word heading (a shared Heading style is set); it
  // sets the body heading-row level so the preview matches the export.
  const titleIsHeading = isHeadingStyleSet(headingStyleName);
  // The active table's rendered fragment for the pinned preview + per-table copy.
  const activeHtml = useMemo(
    () => (activeTable ? tableToHtml(activeTable, titleIsHeading) : ""),
    [activeTable, titleIsHeading],
  );

  // Copy JUST the active table to the clipboard as a Word document. Built
  // synchronously before any await to preserve the user-gesture requirement.
  async function copyForWord() {
    if (
      !activeTable ||
      activeHtml === "" ||
      typeof ClipboardItem === "undefined" ||
      !navigator.clipboard?.write
    ) {
      setCopyState("error");
      setTimeout(() => setCopyState("idle"), 2000);
      return;
    }
    try {
      const item = new ClipboardItem({
        "text/html": new Blob([buildWordHtml(activeHtml, headingStyle, bodyFont)], {
          type: "text/html",
        }),
        "text/plain": new Blob([htmlToPlainText(activeHtml)], {
          type: "text/plain",
        }),
      });
      await navigator.clipboard.write([item]);
      setCopyState("copied");
    } catch {
      setCopyState("error");
    }
    setTimeout(() => setCopyState("idle"), 2000);
  }

  // One level row per depth actually in use across all pivots: bucket (indent
  // level) count + 1 for a title (the title is level 1). Stacked fields share a
  // bucket, so they don't add depth. Clamped to Word's 9-level max.
  const maxDepth = Math.min(
    9,
    Math.max(
      1,
      ...tables.map(
        (t) => t.pivotLevels.length + (t.sectionTitle.trim() ? 1 : 0),
      ),
    ),
  );

  // The paste drop zone, reused for the empty state (big) and the workspace
  // (compact). Focus + Ctrl/Cmd+V appends a table.
  const pasteZone = (big: boolean) => (
    <div
      tabIndex={0}
      aria-label="Paste area. Focus here, then press Control plus V to paste a table copied from Excel or Google Sheets."
      aria-keyshortcuts="Control+V"
      onPaste={handlePaste}
      className={`flex cursor-text items-center justify-center rounded-xl border-2 border-dashed border-foreground/25 bg-foreground/[0.02] text-center text-foreground/60 outline-none transition-colors focus:border-foreground/50 focus:bg-foreground/[0.04] ${
        big ? "min-h-[10rem] p-8 text-base" : "min-h-[4.5rem] p-4 text-sm"
      }`}
    >
      {atLimit ? (
        <span>
          Reached the {MAX_TABLES}-table limit. Remove a table to add another.
        </span>
      ) : tables.length === 0 ? (
        <span>
          Click here and press{" "}
          <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
            Ctrl
          </kbd>{" "}
          +{" "}
          <kbd className="rounded border border-foreground/30 px-1.5 py-0.5 font-mono text-xs">
            V
          </kbd>{" "}
          to paste a table copied from Excel or Google Sheets.
        </span>
      ) : (
        <span>
          Paste another table to add a section (click here, then{" "}
          <kbd className="rounded border border-foreground/30 px-1 py-0.5 font-mono text-[11px]">
            Ctrl
          </kbd>
          +
          <kbd className="rounded border border-foreground/30 px-1 py-0.5 font-mono text-[11px]">
            V
          </kbd>
          ).
        </span>
      )}
    </div>
  );

  const errorBanner = error && (
    <p
      role="alert"
      className="rounded-lg border border-amber-500/40 bg-amber-500/10 px-4 py-3 text-sm text-amber-700 dark:text-amber-300"
    >
      {error}
    </p>
  );
  // Visually-hidden polite live region: announces a successful paste.
  const statusRegion = (
    <p className="sr-only" role="status" aria-live="polite">
      {status}
    </p>
  );

  const actionBtn =
    "rounded-md border border-foreground/20 px-3 py-1 text-xs font-medium transition-colors hover:bg-foreground/5 disabled:cursor-not-allowed disabled:opacity-50";

  // ---- Empty state: a single big paste zone, no chrome yet. -------------------
  if (tables.length === 0) {
    return (
      <div className="mx-auto flex w-full max-w-2xl flex-1 flex-col justify-center gap-6 px-6 py-12">
        <header className="flex flex-col gap-2">
          <h1 className="text-2xl font-bold tracking-tight">
            Excel &rarr; Word Sections
          </h1>
          <p className="text-foreground/60">
            Paste a table copied from Excel or Google Sheets to restructure it
            into a Word-ready nested outline, then copy it for Word.
          </p>
        </header>
        {pasteZone(true)}
        {errorBanner}
        {statusRegion}
      </div>
    );
  }

  // ---- Workspace: the 4-pane IDE shell. --------------------------------------
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      {/* TOP control band: identity + global actions, with a collapsible global
          styling panel (everything here applies to ALL tables). */}
      <header className="flex flex-col border-b border-foreground/15 bg-foreground/[0.02]">
        <div className="flex flex-wrap items-center gap-3 px-4 py-2">
          <h1 className="text-sm font-semibold tracking-tight">
            Excel &rarr; Word
          </h1>
          <span className="text-xs text-muted">
            {tables.length} of {MAX_TABLES} sections
          </span>
          <button
            type="button"
            onClick={() => setShowStyles((v) => !v)}
            aria-expanded={showStyles}
            className={actionBtn}
          >
            {showStyles ? "▾" : "▸"} Level styles &amp; document
          </button>
          <ScopeBadge>Applies to all tables</ScopeBadge>
          <div className="ml-auto flex items-center gap-2">
            <button
              type="button"
              onClick={copyAll}
              title="Copies every table as one Word doc (stacked in paste order)."
              className={actionBtn}
            >
              {copyAllState === "copied"
                ? "Copied all!"
                : copyAllState === "error"
                  ? "Copy failed"
                  : "Copy all"}
            </button>
            <button type="button" onClick={clearAll} className={actionBtn}>
              Clear all
            </button>
          </div>
        </div>

        {showStyles && (
          <div className="flex flex-col gap-3 border-t border-foreground/10 px-4 py-3 text-sm text-foreground/70">
            <p className="text-xs text-muted">
              Set a <strong>Heading style</strong> name (e.g.{" "}
              <strong>Heading 1</strong>) and paste with{" "}
              <strong>Use Destination Styles</strong> so the title adopts your
              document&rsquo;s heading. Leave it blank to use the app&rsquo;s
              direct level-1 look. The body always uses the Level styles below.
            </p>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
              <label className="flex items-center gap-1.5">
                Heading style
                <input
                  type="text"
                  value={headingStyleName}
                  onChange={(e) => setHeadingStyleName(e.target.value)}
                  placeholder="(direct)"
                  aria-label="Word style name for the title"
                  className="w-32 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                />
              </label>
              <label className="flex items-center gap-1.5">
                Body font
                <select
                  value={bodyFont}
                  onChange={(e) => setBodyFont(e.target.value)}
                  className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                >
                  {HEADING_FONTS.map((f) => (
                    <option key={f} value={f}>
                      {f}
                    </option>
                  ))}
                </select>
              </label>
              <label className="flex items-center gap-1.5">
                Indent/level (in)
                <input
                  type="text"
                  inputMode="decimal"
                  value={indentInput}
                  onChange={(e) =>
                    setIndentInput(e.target.value.replace(/[^0-9.]/g, ""))
                  }
                  aria-label="Indent per nesting level, in inches"
                  className="w-14 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                />
              </label>
              <button
                type="button"
                onClick={() => setLevelStyles([])}
                disabled={levelStyles.length === 0}
                title="Reset every level to the default look"
                className={actionBtn}
              >
                Reset levels
              </button>
            </div>
            <div className="flex flex-col gap-1.5">
              <span className="text-foreground/60">
                Level styles{" "}
                <span className="text-muted">
                  (Level 1 = title when no Heading style is set; 2-9 = nested rows)
                </span>
              </span>
              {Array.from({ length: maxDepth }, (_, i) => {
                const lv = headingStyle.levels[i];
                return (
                  <div
                    key={i}
                    className="flex flex-wrap items-center gap-x-3 gap-y-1.5"
                  >
                    <span className="w-14 text-foreground/60">Level {i + 1}</span>
                    <label className="flex items-center gap-1.5">
                      Color
                      <input
                        type="color"
                        value={lv.color}
                        onChange={(e) => setLevel(i, { color: e.target.value })}
                        aria-label={`Level ${i + 1} color`}
                        className="h-6 w-8 cursor-pointer rounded border border-foreground/20"
                      />
                    </label>
                    <label className="flex items-center gap-1.5">
                      Font
                      <select
                        value={lv.font}
                        onChange={(e) => setLevel(i, { font: e.target.value })}
                        aria-label={`Level ${i + 1} font`}
                        className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                      >
                        {HEADING_FONTS.map((f) => (
                          <option key={f} value={f}>
                            {f}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="flex items-center gap-1.5">
                      pt
                      <input
                        type="text"
                        inputMode="numeric"
                        value={levelStyles[i]?.sizeInput ?? DEFAULT_LEVEL.sizeInput}
                        onChange={(e) =>
                          setLevel(i, {
                            sizeInput: e.target.value.replace(/[^0-9]/g, ""),
                          })
                        }
                        aria-label={`Level ${i + 1} size in points`}
                        className="w-12 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                      />
                    </label>
                    <label className="flex items-center gap-1.5">
                      <input
                        type="checkbox"
                        checked={lv.bold}
                        onChange={(e) => setLevel(i, { bold: e.target.checked })}
                      />
                      Bold
                    </label>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </header>

      {errorBanner && <div className="px-4 pt-3">{errorBanner}</div>}
      {statusRegion}

      {/* WORKSPACE: left sections rail | center builder | right pinned preview.
          overflow-x-auto so a very narrow viewport can scroll to the right pane
          instead of the fixed-width panes being clipped by body overflow-hidden. */}
      <div className="flex min-h-0 flex-1 overflow-x-auto">
        {/* LEFT: sections navigator. */}
        <aside className="flex w-56 shrink-0 flex-col overflow-y-auto border-r border-foreground/15 bg-foreground/[0.015] p-2">
          <div className="flex items-center justify-between px-1 pb-1.5">
            <span className="text-xs font-medium uppercase tracking-wide text-muted">
              Sections
            </span>
          </div>
          <div className="flex flex-col gap-1">
            {tables.map((t, i) => {
              const isActive = activeTable?.id === t.id;
              const label = t.sectionTitle.trim() || `Table ${i + 1}`;
              return (
                <div
                  key={t.id}
                  className={`group flex items-center gap-1.5 rounded-md border-l-2 py-1.5 pr-1.5 pl-2 text-xs transition-colors ${
                    isActive
                      ? "border-l-foreground/70 bg-foreground/[0.07] font-medium text-foreground"
                      : "border-l-transparent text-foreground/60 hover:bg-foreground/5"
                  }`}
                >
                  <span className="flex h-4 w-4 shrink-0 items-center justify-center rounded bg-foreground/10 text-[10px] tabular-nums text-foreground/60">
                    {i + 1}
                  </span>
                  <button
                    type="button"
                    onClick={() => setActiveId(t.id)}
                    className="min-w-0 flex-1 truncate text-left"
                    title={label}
                  >
                    {label}
                  </button>
                  <button
                    type="button"
                    onClick={() => removeTable(t.id)}
                    aria-label={`Remove ${label}`}
                    title="Remove this table"
                    className="shrink-0 px-1 leading-none text-foreground/30 opacity-0 transition-opacity hover:text-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100"
                  >
                    &times;
                  </button>
                </div>
              );
            })}
          </div>
          <p className="mt-2 px-1 text-[11px] leading-snug text-muted">
            Paste another table (in the center) to add a section.
          </p>
        </aside>

        {/* CENTER: paste zone + the active table's builder. */}
        <section className="flex min-w-0 flex-1 flex-col gap-3 overflow-y-auto p-4">
          {pasteZone(false)}
          {activeTable && (
            <TableCard
              key={activeTable.id}
              table={activeTable}
              onChange={(patch) => patchTable(activeTable.id, patch)}
            />
          )}
        </section>

        {/* RIGHT: pinned live preview + per-table export. */}
        <section className="flex w-[38%] min-w-[20rem] shrink-0 flex-col overflow-hidden border-l border-foreground/15">
          <div className="flex items-center gap-2 border-b border-foreground/15 bg-foreground/[0.02] px-3 py-1.5">
            <div className="flex items-center gap-0.5 rounded-md border border-foreground/15 p-0.5 text-xs">
              <button
                type="button"
                aria-pressed={view === "rendered"}
                onClick={() => setView("rendered")}
                className={`rounded px-2 py-0.5 transition-colors ${
                  view === "rendered"
                    ? "bg-foreground/10 font-medium text-foreground"
                    : "text-foreground/60 hover:bg-foreground/5"
                }`}
              >
                Preview
              </button>
              <button
                type="button"
                aria-pressed={view === "json"}
                onClick={() => setView("json")}
                className={`rounded px-2 py-0.5 transition-colors ${
                  view === "json"
                    ? "bg-foreground/10 font-medium text-foreground"
                    : "text-foreground/60 hover:bg-foreground/5"
                }`}
              >
                JSON
              </button>
            </div>
            <ScopeBadge>This table</ScopeBadge>
            <button
              type="button"
              onClick={copyForWord}
              disabled={activeHtml === ""}
              title="Copies this table's HTML to paste into Word. For the title to match your document's Heading 1, paste with 'Use Destination Styles'."
              className={`ml-auto ${actionBtn}`}
            >
              {copyState === "copied"
                ? "Copied!"
                : copyState === "error"
                  ? "Copy failed"
                  : "Copy for Word"}
            </button>
          </div>
          <div className="min-h-0 flex-1 overflow-y-auto p-4">
            {view === "json" ? (
              activeTable ? (
                <JsonPreview grid={activeTable.grid} />
              ) : null
            ) : (
              <RenderedPreview
                html={activeHtml}
                headingStyle={headingStyle}
                bodyFont={bodyFont}
                emptyHint="Add fields in the center to build the outline. Stack fields at one indent level to show them together (like an Excel pivot's Row Labels)."
              />
            )}
          </div>
        </section>
      </div>
    </div>
  );
}

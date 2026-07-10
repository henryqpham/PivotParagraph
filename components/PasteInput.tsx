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
import { TableCard, type TitleInput } from "./TableCard";
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
// Default look for an untouched level: plain Arial 11 black (matches a document body).
const DEFAULT_LEVEL: LevelInput = {
  color: "#000000",
  font: "Arial",
  sizeInput: "11",
  bold: false,
};
// Default TITLE look (its own shared style now — Arial 11 black, matching the old
// level-1 default; only shows when the Heading dropdown is None).
const DEFAULT_TITLE: TitleInput = {
  font: "Arial",
  sizeInput: "11",
  color: "#000000",
  bold: false,
  italic: false,
  underline: false,
};

// ---- Fluent control recipes -----------------------------------------------
const BTN_PRIMARY =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-accent-fg bg-accent transition-colors hover:bg-accent-hover active:bg-accent-pressed disabled:cursor-not-allowed disabled:opacity-50";
const BTN_SUBTLE =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-text-secondary bg-transparent transition-colors hover:bg-surface-alt hover:text-foreground disabled:cursor-not-allowed disabled:opacity-50";
const BTN_STD =
  "inline-flex h-8 items-center gap-1.5 rounded border border-border-strong bg-surface px-3 text-sm font-semibold text-foreground transition-colors hover:bg-surface-alt disabled:cursor-not-allowed disabled:opacity-50";
const FIELD =
  "h-8 rounded border border-border-strong bg-surface px-2.5 text-sm text-foreground outline-none transition-colors focus:border-accent";
const BADGE_CLS =
  "rounded-sm bg-[color:color-mix(in_srgb,var(--muted)_14%,transparent)] px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-muted";

/** Fluent scope badge. */
function ScopeBadge({ children }: { children: ReactNode }) {
  return <span className={BADGE_CLS}>{children}</span>;
}

export function PasteInput() {
  const [tables, setTables] = useState<TableState[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("");
  const [activeId, setActiveId] = useState<string | null>(null);
  const [copyAllState, setCopyAllState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">("idle");
  const [view, setView] = useState<"rendered" | "json">("rendered");
  const [showStyles, setShowStyles] = useState(false);
  const idRef = useRef(0);

  // Document body font (default Arial).
  const [bodyFont, setBodyFont] = useState<string>("Arial");
  // Per-level BODY styling, shared across tables (the "level chart"). Sparse.
  const [levelStyles, setLevelStyles] = useState<LevelInput[]>([]);
  // Left-indent per nesting level (inches), clamped [0, 2].
  const [indentInput, setIndentInput] = useState<string>("0.2");
  // Global Word heading style the TITLE maps to ("" = None). Driven by the
  // Section Header dropdown in TableCard.
  const [headingStyleName, setHeadingStyleName] = useState<string>("Heading 1");
  // Shared TITLE look (its own controls in the Section Header group).
  const [titleInput, setTitleInput] = useState<TitleInput>(DEFAULT_TITLE);

  const headingStyle = useMemo<HeadingStyle>(() => {
    const clampPt = (s: string, fallback: number) => {
      const n = parseInt(s, 10);
      return Number.isFinite(n) && n >= 1 && n <= 72 ? n : fallback;
    };
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
    return {
      levels,
      indentStep,
      headingStyleName,
      titleStyle: {
        font: titleInput.font,
        size: clampPt(titleInput.sizeInput, 11),
        color: titleInput.color,
        bold: titleInput.bold,
        italic: titleInput.italic,
        underline: titleInput.underline,
      },
    };
  }, [levelStyles, indentInput, headingStyleName, titleInput]);

  function setLevel(i: number, patch: Partial<LevelInput>) {
    setLevelStyles((prev) => {
      const next = [...prev];
      next[i] = { ...(next[i] ?? DEFAULT_LEVEL), ...patch };
      return next;
    });
  }
  function setTitle(patch: Partial<TitleInput>) {
    setTitleInput((prev) => ({ ...prev, ...patch }));
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
        pivotLevels: [],
        markers: [],
        fieldLabels: {},
        sortDirs: {},
        breakAfter: [],
        numbering: DEFAULT_NUMBERING,
        headingLevels: [],
        sectionTitle: "",
      };
      setTables((prev) => [...prev, next]);
      setActiveId(id);
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

  async function copyAll() {
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyAllState("error");
      setTimeout(() => setCopyAllState("idle"), 2000);
      return;
    }
    const titleIsHeading = isHeadingStyleSet(headingStyleName);
    const combined = tables.map((t) => tableToHtml(t, titleIsHeading)).join("\n");
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
  const activeTable = tables.find((t) => t.id === activeId) ?? tables[0] ?? null;

  const titleIsHeading = isHeadingStyleSet(headingStyleName);
  const activeHtml = useMemo(
    () => (activeTable ? tableToHtml(activeTable, titleIsHeading) : ""),
    [activeTable, titleIsHeading],
  );

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

  const maxDepth = Math.min(
    9,
    Math.max(
      1,
      ...tables.map((t) => t.pivotLevels.length + (t.sectionTitle.trim() ? 1 : 0)),
    ),
  );

  const pasteZone = (big: boolean) => (
    <div
      tabIndex={0}
      aria-label="Paste area. Focus here, then press Control plus V to paste a table copied from Excel or Google Sheets."
      aria-keyshortcuts="Control+V"
      onPaste={handlePaste}
      className={`flex cursor-text items-center justify-center rounded-lg border-2 border-dashed border-border-strong bg-surface-alt text-center text-muted outline-none transition-colors focus:border-accent focus:bg-accent-subtle ${
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
          <kbd className="rounded border border-border-strong px-1.5 py-0.5 font-mono text-xs">
            Ctrl
          </kbd>{" "}
          +{" "}
          <kbd className="rounded border border-border-strong px-1.5 py-0.5 font-mono text-xs">
            V
          </kbd>{" "}
          to paste a table copied from Excel or Google Sheets.
        </span>
      ) : (
        <span>
          Paste another table to add a section (click here, then{" "}
          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
            Ctrl
          </kbd>
          +
          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
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
      className="rounded border border-[color:var(--danger-border)] bg-[color:var(--danger-bg)] px-4 py-3 text-sm text-danger"
    >
      {error}
    </p>
  );
  const statusRegion = (
    <p className="sr-only" role="status" aria-live="polite">
      {status}
    </p>
  );

  // ---- The Fluent 4-pane IDE (shown ALWAYS — even with no tables yet, so the
  //      full layout/outline is visible; the center shows a big paste zone plus
  //      empty Section Header / Levels placeholder cards until you paste). -------
  return (
    <div className="flex min-h-0 flex-1 flex-col bg-background">
      {/* TOP command band */}
      <header className="flex flex-col border-b border-border bg-surface shadow-[var(--shadow-2)]">
        <div className="flex flex-wrap items-center gap-2 px-2 py-1.5">
          <span className="ml-1 flex items-center gap-2 font-semibold">
            <span className="grid h-5 w-5 place-items-center rounded-[3px] bg-accent text-xs font-bold text-accent-fg">
              W
            </span>
            <h1 className="text-sm">Excel &rarr; Word</h1>
          </span>
          <span className="text-xs text-muted">
            {tables.length} of {MAX_TABLES} sections
          </span>
          <span className="mx-1 h-6 w-px bg-border" />
          <button
            type="button"
            onClick={() => setShowStyles((v) => !v)}
            aria-expanded={showStyles}
            className={BTN_SUBTLE}
          >
            {showStyles ? "▾" : "▸"} Level styles &amp; document
          </button>
          <ScopeBadge>All tables</ScopeBadge>
          <div className="ml-auto flex items-center gap-2">
            <button
              type="button"
              onClick={clearAll}
              disabled={tables.length === 0}
              className={BTN_SUBTLE}
            >
              Clear all
            </button>
            <button
              type="button"
              onClick={copyAll}
              disabled={tables.length === 0}
              title="Copies every table as one Word doc (stacked in paste order)."
              className={BTN_PRIMARY}
            >
              {copyAllState === "copied"
                ? "Copied all!"
                : copyAllState === "error"
                  ? "Copy failed"
                  : "Copy all"}
            </button>
          </div>
        </div>

        {showStyles && (
          <div className="flex flex-col gap-3 border-t border-border px-4 py-3 text-sm text-text-secondary">
            <p className="text-xs text-muted">
              These style the <strong>nested body rows</strong> and the document,
              and apply to every table. (The title is styled in its own{" "}
              <strong>Section Header</strong> group.)
            </p>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
              <label className="flex items-center gap-1.5">
                Body font
                <select
                  value={bodyFont}
                  onChange={(e) => setBodyFont(e.target.value)}
                  className={FIELD}
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
                  className={`${FIELD} w-14`}
                />
              </label>
              <button
                type="button"
                onClick={() => setLevelStyles([])}
                disabled={levelStyles.length === 0}
                title="Reset every level to the default look"
                className={BTN_STD}
              >
                Reset levels
              </button>
            </div>
            <div className="flex flex-col gap-1.5">
              <span className="text-xs font-semibold uppercase tracking-wide text-muted">
                Level styles{" "}
                <span className="font-normal normal-case">
                  (nested body rows, by depth)
                </span>
              </span>
              {Array.from({ length: maxDepth }, (_, i) => {
                const lv = headingStyle.levels[i];
                return (
                  <div
                    key={i}
                    className="flex flex-wrap items-center gap-x-3 gap-y-1.5"
                  >
                    <span className="w-14 text-text-secondary">Level {i + 1}</span>
                    <label className="flex items-center gap-1.5">
                      Color
                      <input
                        type="color"
                        value={lv.color}
                        onChange={(e) => setLevel(i, { color: e.target.value })}
                        aria-label={`Level ${i + 1} color`}
                        className="h-7 w-8 cursor-pointer rounded border border-border-strong"
                      />
                    </label>
                    <label className="flex items-center gap-1.5">
                      Font
                      <select
                        value={lv.font}
                        onChange={(e) => setLevel(i, { font: e.target.value })}
                        aria-label={`Level ${i + 1} font`}
                        className={FIELD}
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
                        className={`${FIELD} w-12`}
                      />
                    </label>
                    <label className="flex items-center gap-1.5">
                      <input
                        type="checkbox"
                        checked={lv.bold}
                        onChange={(e) => setLevel(i, { bold: e.target.checked })}
                        className="accent-[var(--accent)]"
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

      {/* WORKSPACE */}
      <div className="flex min-h-0 flex-1 overflow-x-auto">
        {/* LEFT: sections rail */}
        <aside className="flex w-56 shrink-0 flex-col overflow-y-auto border-r border-border bg-surface-alt p-2">
          <div className="border-b border-border px-2 pb-1.5 pt-1 text-xs font-semibold uppercase tracking-wide text-muted">
            Sections
          </div>
          <div className="mt-1 flex flex-col gap-0.5">
            {tables.map((t, i) => {
              const isActive = activeTable?.id === t.id;
              const label = t.sectionTitle.trim() || `Table ${i + 1}`;
              return (
                <div
                  key={t.id}
                  className={`group relative flex h-9 items-center gap-2 rounded pl-3 pr-1.5 text-sm transition-colors ${
                    isActive
                      ? "bg-accent-subtle font-semibold text-accent-text before:absolute before:left-0 before:top-1.5 before:bottom-1.5 before:w-[3px] before:rounded-full before:bg-accent before:content-['']"
                      : "text-text-secondary hover:bg-surface"
                  }`}
                >
                  <span className="grid h-[18px] w-[18px] shrink-0 place-items-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] tabular-nums text-muted">
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
                    className="shrink-0 px-1 leading-none text-muted opacity-0 transition-opacity hover:text-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100"
                  >
                    &times;
                  </button>
                </div>
              );
            })}
          </div>
          <p className="mt-2 px-2 text-[11px] leading-snug text-muted">
            {tables.length === 0
              ? "No sections yet — paste a table in the center to add one."
              : "Paste another table (in the center) to add a section."}
          </p>
        </aside>

        {/* CENTER: the two command groups (or empty placeholders before a paste) */}
        <section className="flex min-w-0 flex-1 flex-col gap-4 overflow-y-auto p-4">
          {pasteZone(tables.length === 0)}
          {activeTable ? (
            <TableCard
              key={activeTable.id}
              table={activeTable}
              onChange={(patch) => patchTable(activeTable.id, patch)}
              headingStyleName={headingStyleName}
              onHeadingStyleChange={setHeadingStyleName}
              title={titleInput}
              onTitleChange={setTitle}
            />
          ) : (
            <>
              <section className="rounded-lg border border-border bg-surface p-4 shadow-[var(--shadow-2)]">
                <h2 className="mb-3 border-b border-border pb-2 text-xs font-semibold uppercase tracking-wide text-muted">
                  Section Header
                </h2>
                <p className="text-sm text-muted">
                  The section title, its Word heading, and its look
                  (font/size/B/I/U/color) appear here once you paste a table.
                </p>
              </section>
              <section className="rounded-lg border border-border bg-surface p-4 shadow-[var(--shadow-2)]">
                <h2 className="mb-3 border-b border-border pb-2 text-xs font-semibold uppercase tracking-wide text-muted">
                  Levels
                </h2>
                <p className="text-sm text-muted">
                  Paste a table, then arrange its columns into the nested outline
                  here — add fields, indent, number, and mark Word headings.
                </p>
              </section>
            </>
          )}
        </section>

        {/* RIGHT: pinned preview */}
        <section className="flex w-[38%] min-w-80 shrink-0 flex-col overflow-hidden border-l border-border">
          <div className="flex items-center gap-2 border-b border-border bg-surface px-2.5 py-1.5">
            <div className="inline-flex gap-0.5 rounded bg-surface-alt p-0.5 text-xs">
              <button
                type="button"
                aria-pressed={view === "rendered"}
                onClick={() => setView("rendered")}
                className={`rounded px-2.5 py-1 font-semibold transition-colors ${
                  view === "rendered"
                    ? "bg-surface text-foreground shadow-[var(--shadow-2)]"
                    : "text-text-secondary hover:text-foreground"
                }`}
              >
                Preview
              </button>
              <button
                type="button"
                aria-pressed={view === "json"}
                onClick={() => setView("json")}
                className={`rounded px-2.5 py-1 font-semibold transition-colors ${
                  view === "json"
                    ? "bg-surface text-foreground shadow-[var(--shadow-2)]"
                    : "text-text-secondary hover:text-foreground"
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
              title="Copies this table's HTML to paste into Word. Paste with 'Use Destination Styles' so the title maps to your heading."
              className={`ml-auto ${BTN_PRIMARY}`}
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

"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { parseClipboard } from "@/lib/parser";
import {
  buildWordHtml,
  htmlToPlainText,
  headingLevel,
} from "@/lib/clipboard";
import type { HeadingStyle, LevelStyle } from "@/lib/clipboard";
import { loadSession, saveSession } from "@/lib/persistence";
import {
  tableToHtml,
  newTable,
  makeExampleTable,
  DEFAULT_LEVEL,
  type LevelInput,
  type TableState,
} from "./tableModel";
import { TableCard, FONTS, type TitleInput } from "./TableCard";
import { SectionsRail } from "./SectionsRail";
import { Popover } from "./Popover";
import { RenderedPreview } from "./RenderedPreview";
import { JsonPreview } from "./JsonPreview";

const MAX_TABLES = 100;

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

/** The full workspace state we persist to localStorage (all JSON-serializable). */
type SessionSnapshot = {
  tables: TableState[];
  levelStyles: LevelInput[];
  bodyFont: string;
  indentInput: string;
  headingStyleName: string;
  titleInput: TitleInput;
  activeId: string | null;
};

/** The post-copy confirmation shown near "Copy all" (and announced to a11y). */
type CopyNote =
  | { status: "ok"; count: number; hasHeading: boolean }
  | { status: "empty" }
  | { status: "error" };

/** A one-step undo snapshot for a destructive action (remove one / clear all). */
type UndoState = { label: string; tables: TableState[]; activeId: string | null };

// ---- Fluent control recipes -----------------------------------------------
const BTN_PRIMARY =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-accent-fg bg-accent transition-colors hover:bg-accent-hover active:bg-accent-pressed disabled:cursor-not-allowed disabled:opacity-50";
const BTN_SUBTLE =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-text-secondary bg-transparent transition-colors hover:bg-surface-alt hover:text-foreground disabled:cursor-not-allowed disabled:opacity-50";
// Input/select recipe for the Document popover (matches TableCard's FIELD).
const FIELD =
  "h-8 rounded border border-border-strong bg-surface px-2.5 text-sm text-foreground outline-none transition-colors hover:border-b-[color:var(--text-secondary)] focus:border-accent";

export function PasteInput() {
  const [tables, setTables] = useState<TableState[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("");
  const [activeId, setActiveId] = useState<string | null>(null);
  const [copyAllState, setCopyAllState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  const [view, setView] = useState<"rendered" | "json">("rendered");
  // Preview scope: the ACTIVE section (focused editing) or ALL sections stacked
  // exactly as Copy all emits them (the actual deliverable).
  const [previewScope, setPreviewScope] = useState<"active" | "all">("active");
  // Whether the top-band "Document" (global body settings) popover is open.
  const [showDoc, setShowDoc] = useState(false);
  // Post-copy confirmation + Word-paste guidance (persists until the next edit).
  const [copyNote, setCopyNote] = useState<CopyNote | null>(null);
  // Two-step guard on the destructive "Clear all".
  const [confirmClear, setConfirmClear] = useState(false);
  // One-step undo for the last remove/clear.
  const [undo, setUndo] = useState<UndoState | null>(null);
  // Gate persistence writes until after the initial localStorage load, so we never
  // clobber saved work with the empty initial state.
  const [hydrated, setHydrated] = useState(false);
  const idRef = useRef(0);
  const undoBtnRef = useRef<HTMLButtonElement>(null);
  const undoTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  // Latest snapshot + a hydration flag, kept in refs so the tab-hide flush can
  // save synchronously (from an event, without stale closure state) and never
  // overwrite saved work before the initial load has run.
  const snapshotRef = useRef<SessionSnapshot | null>(null);
  const hydratedRef = useRef(false);

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

  // ---- Persistence: rehydrate once on mount -------------------------------
  // Loading from localStorage is inherently a client-only, mount-time sync with an
  // external store: the first client render must match the empty server render, so
  // we restore here (not in a lazy initializer, which would cause a hydration
  // mismatch). eslint-plugin-react-hooks flags the synchronous setState, but every
  // SSR-safe alternative either mismatches hydration or trips the same rule, so the
  // disable is deliberate and scoped to this one legitimate effect.
  useEffect(() => {
    /* eslint-disable react-hooks/set-state-in-effect */
    const data = loadSession();
    if (
      data &&
      typeof data === "object" &&
      Array.isArray((data as SessionSnapshot).tables)
    ) {
      const s = data as SessionSnapshot;
      setTables(s.tables);
      setLevelStyles(Array.isArray(s.levelStyles) ? s.levelStyles : []);
      setBodyFont(s.bodyFont ?? "Arial");
      setIndentInput(s.indentInput ?? "0.2");
      setHeadingStyleName(s.headingStyleName ?? "Heading 1");
      setTitleInput(s.titleInput ?? DEFAULT_TITLE);
      const restoredActive =
        s.tables.find((t) => t.id === s.activeId)?.id ?? s.tables[0]?.id ?? null;
      setActiveId(restoredActive);
      // Re-seed the id counter past the largest restored id so new pastes never
      // collide with a restored `t<N>`.
      idRef.current = s.tables.reduce((max, t) => {
        const n = parseInt(String(t.id).replace(/\D/g, ""), 10);
        return Number.isFinite(n) ? Math.max(max, n) : max;
      }, 0);
    }
    hydratedRef.current = true;
    setHydrated(true);
    /* eslint-enable react-hooks/set-state-in-effect */
  }, []);

  // The full workspace snapshot; one memo feeds both the debounced save and the
  // synchronous tab-hide flush so they can never disagree.
  const snapshot = useMemo<SessionSnapshot>(
    () => ({
      tables,
      levelStyles,
      bodyFont,
      indentInput,
      headingStyleName,
      titleInput,
      activeId,
    }),
    [
      tables,
      levelStyles,
      bodyFont,
      indentInput,
      headingStyleName,
      titleInput,
      activeId,
    ],
  );
  useEffect(() => {
    snapshotRef.current = snapshot;
  }, [snapshot]);

  // ---- Persistence: debounced save on any workspace change ----------------
  useEffect(() => {
    if (!hydrated) return;
    const handle = setTimeout(() => saveSession(snapshot), 400);
    return () => clearTimeout(handle);
  }, [hydrated, snapshot]);

  // ---- Persistence: flush immediately when the tab is hidden or closed, so a
  //      change made inside the 400ms debounce window isn't lost. Gated on the
  //      hydration ref so an early hide never overwrites saved work with the
  //      empty initial state. ------------------------------------------------
  useEffect(() => {
    const flush = () => {
      if (hydratedRef.current && snapshotRef.current) {
        saveSession(snapshotRef.current);
      }
    };
    const onVisibility = () => {
      if (document.visibilityState === "hidden") flush();
    };
    window.addEventListener("pagehide", flush);
    document.addEventListener("visibilitychange", onVisibility);
    return () => {
      window.removeEventListener("pagehide", flush);
      document.removeEventListener("visibilitychange", onVisibility);
    };
  }, []);

  // Clear the copy confirmation + any pending clear-confirm when the user edits.
  // Called from the mutation handlers (an event-time call, NOT an effect) so the
  // note reads as "…until you change something" without a cascading render.
  function noteEdit() {
    setCopyNote(null);
    setConfirmClear(false);
  }

  // The pending "Clear all?" confirm auto-cancels if left untouched.
  useEffect(() => {
    if (!confirmClear) return;
    const handle = setTimeout(() => setConfirmClear(false), 4000);
    return () => clearTimeout(handle);
  }, [confirmClear]);

  // When an undo affordance appears, move focus onto it so a keyboard user who
  // just removed a section lands on Undo instead of being dropped to <body>.
  useEffect(() => {
    if (undo) undoBtnRef.current?.focus();
  }, [undo]);

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
    noteEdit();
    setLevelStyles((prev) => {
      const next = [...prev];
      next[i] = { ...(next[i] ?? DEFAULT_LEVEL), ...patch };
      return next;
    });
  }
  function setTitle(patch: Partial<TitleInput>) {
    noteEdit();
    setTitleInput((prev) => ({ ...prev, ...patch }));
  }
  // Wrap the doc/heading setters passed to children so any edit also clears the
  // transient copy note (the raw useState setters go straight through otherwise).
  function changeHeadingStyleName(v: string) {
    noteEdit();
    setHeadingStyleName(v);
  }
  function changeBodyFont(v: string) {
    noteEdit();
    setBodyFont(v);
  }
  function changeIndentInput(v: string) {
    noteEdit();
    setIndentInput(v);
  }
  function resetLevels() {
    noteEdit();
    setLevelStyles([]);
  }

  // ---- Paste ingestion (shared by the paste zone + the paste-anywhere
  //      document listener + Try-an-example) ---------------------------------
  function ingestGrid(grid: TableState["grid"]): boolean {
    if (tables.length >= MAX_TABLES) {
      setError(
        `Reached the ${MAX_TABLES}-section limit. Remove a section to add another.`,
      );
      return false;
    }
    noteEdit();
    const id = `t${++idRef.current}`;
    setTables((prev) => [...prev, newTable(id, grid)]);
    setActiveId(id);
    setError(null);
    setStatus(`Added section ${tables.length + 1}.`);
    return true;
  }

  function ingestClipboard(data: DataTransfer) {
    try {
      const rows = parseClipboard(data);
      if (rows.length === 0) {
        setError(
          "Couldn't read a table from the clipboard. Copy a cell range from Excel or Google Sheets, then paste here.",
        );
        return;
      }
      ingestGrid(rows);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Failed to parse the pasted data.",
      );
    }
  }

  function addExample() {
    if (tables.length >= MAX_TABLES) return;
    noteEdit();
    const id = `t${++idRef.current}`;
    setTables((prev) => [...prev, makeExampleTable(id)]);
    setActiveId(id);
    setError(null);
    setStatus("Loaded the example section.");
  }

  // Paste anywhere: a table copied in Excel drops in on Ctrl/Cmd+V from anywhere,
  // UNLESS focus is in a real text field (Title, Start, size inputs, selects) —
  // there the paste belongs to that field. Keep a ref to the latest handler so the
  // listener stays mounted once.
  const ingestRef = useRef(ingestClipboard);
  useEffect(() => {
    ingestRef.current = ingestClipboard;
  });
  useEffect(() => {
    function onPaste(e: ClipboardEvent) {
      const el = document.activeElement as HTMLElement | null;
      if (
        el &&
        (el.tagName === "INPUT" ||
          el.tagName === "TEXTAREA" ||
          el.tagName === "SELECT" ||
          el.isContentEditable)
      ) {
        return; // let the focused field handle its own paste
      }
      if (!e.clipboardData) return;
      e.preventDefault();
      ingestRef.current(e.clipboardData);
    }
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, []);

  function patchTable(id: string, patch: Partial<TableState>) {
    noteEdit();
    setTables((ts) => ts.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  }

  function reorderTable(id: string, dir: -1 | 1) {
    noteEdit();
    setTables((ts) => {
      const i = ts.findIndex((t) => t.id === id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= ts.length) return ts;
      const next = [...ts];
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });
  }

  function pushUndo(label: string) {
    setUndo({ label, tables, activeId });
    if (undoTimer.current) clearTimeout(undoTimer.current);
    undoTimer.current = setTimeout(() => setUndo(null), 7000);
  }

  function doUndo() {
    if (!undo) return;
    setCopyNote(null);
    setTables(undo.tables);
    setActiveId(undo.activeId);
    setUndo(null);
    if (undoTimer.current) clearTimeout(undoTimer.current);
  }

  function removeTable(id: string) {
    setCopyNote(null);
    pushUndo("Section removed");
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
    setCopyNote(null);
    pushUndo("All sections cleared");
    setTables([]);
    setActiveId(null);
    setError(null);
    setConfirmClear(false);
  }

  async function copyAll() {
    const titleLvl = headingLevel(headingStyleName);
    const sectionHtmls = tables.map((t) => tableToHtml(t, titleLvl));
    const contributing = sectionHtmls.filter((h) => h !== "").length;
    if (contributing === 0) {
      setCopyNote({ status: "empty" });
      return;
    }
    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      setCopyNote({ status: "error" });
      setCopyAllState("error");
      setTimeout(() => setCopyAllState("idle"), 2000);
      return;
    }
    const combined = sectionHtmls.join("\n");
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
      const hasHeading =
        headingStyleName !== "" ||
        tables.some((t) => (t.headingLevels ?? []).some(Boolean));
      setCopyNote({ status: "ok", count: contributing, hasHeading });
      setCopyAllState("copied");
    } catch {
      setCopyNote({ status: "error" });
      setCopyAllState("error");
    }
    setTimeout(() => setCopyAllState("idle"), 2000);
  }

  const atLimit = tables.length >= MAX_TABLES;
  const activeTable = tables.find((t) => t.id === activeId) ?? tables[0] ?? null;

  const titleLevel = headingLevel(headingStyleName);
  // Per-section fragments — the SINGLE source for both the combined preview and
  // (recomputed identically in copyAll) the export, so the two can't drift.
  const sectionHtmls = useMemo(
    () => tables.map((t) => tableToHtml(t, titleLevel)),
    [tables, titleLevel],
  );
  const activeHtml = useMemo(
    () => (activeTable ? tableToHtml(activeTable, titleLevel) : ""),
    [activeTable, titleLevel],
  );

  // ---- Paste zone (a focusable drop target; the actual paste is handled by the
  //      document-level listener, so a table drops in whether or not this has
  //      focus). ------------------------------------------------------------
  const pasteZone = (big: boolean) => (
    <div
      tabIndex={0}
      role="button"
      aria-label="Paste area. Copy a cell range from Excel or Google Sheets, then press Control or Command plus V anywhere on the page to add it as a section."
      aria-keyshortcuts="Control+V"
      className={`flex cursor-text items-center justify-center rounded-lg border-2 border-dashed border-border-strong bg-surface-alt text-center text-muted outline-none transition-colors focus:border-accent focus:bg-accent-subtle ${
        big ? "min-h-[9rem] p-8 text-base" : "min-h-[4.5rem] p-4 text-sm"
      }`}
    >
      {atLimit ? (
        <span>
          Reached the {MAX_TABLES}-section limit. Remove a section to add another.
        </span>
      ) : tables.length === 0 ? (
        <span>
          Press{" "}
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
          Paste another table to add a section (
          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
            Ctrl
          </kbd>
          +
          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
            V
          </kbd>
          , anywhere).
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

  // Visible + announced copy confirmation / failure. role=status announces the
  // success politely; role=alert interrupts for a failure.
  const copyBanner = copyNote && (
    <div
      role={copyNote.status === "error" ? "alert" : "status"}
      aria-live={copyNote.status === "error" ? "assertive" : "polite"}
      className={`flex items-start gap-3 rounded border px-4 py-2.5 text-sm ${
        copyNote.status === "error"
          ? "border-[color:var(--danger-border)] bg-[color:var(--danger-bg)] text-danger"
          : "border-[color:color-mix(in_srgb,var(--accent)_35%,transparent)] bg-accent-subtle text-foreground"
      }`}
    >
      <div className="min-w-0 flex-1">
        {copyNote.status === "error" ? (
          <span>
            Couldn&apos;t copy to the clipboard — your browser may block clipboard
            access. Try again, or use a Chromium-based browser.
          </span>
        ) : copyNote.status === "empty" ? (
          <span>
            Nothing to copy yet — add fields to a section, then Copy all.
          </span>
        ) : (
          <span>
            <strong>
              {copyNote.count} section{copyNote.count === 1 ? "" : "s"} copied.
            </strong>{" "}
            In Word, press{" "}
            <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
              Ctrl
            </kbd>
            +
            <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
              V
            </kbd>{" "}
            → choose <strong>Keep Source Formatting</strong> to keep this exact
            look
            {copyNote.hasHeading ? (
              <>
                , or <strong>Use Destination Styles</strong>{" "}
                to adopt the template&apos;s heading (those rows join the
                Navigation pane)
              </>
            ) : null}
            .
          </span>
        )}
      </div>
      <button
        type="button"
        onClick={() => setCopyNote(null)}
        aria-label="Dismiss"
        className="shrink-0 leading-none text-muted transition-colors hover:text-foreground"
      >
        &times;
      </button>
    </div>
  );

  const undoBanner = undo && (
    <div
      role="status"
      className="flex items-center gap-3 rounded border border-border-strong bg-surface px-4 py-2.5 text-sm text-text-secondary shadow-[var(--shadow-2)]"
    >
      <span className="flex-1">{undo.label}.</span>
      <button
        ref={undoBtnRef}
        type="button"
        onClick={doUndo}
        className="rounded px-2 py-1 text-sm font-semibold text-accent-text transition-colors hover:bg-accent-subtle"
      >
        Undo
      </button>
      <button
        type="button"
        onClick={() => setUndo(null)}
        aria-label="Dismiss"
        className="shrink-0 leading-none text-muted transition-colors hover:text-foreground"
      >
        &times;
      </button>
    </div>
  );

  const statusRegion = (
    <p className="sr-only" role="status" aria-live="polite">
      {status}
    </p>
  );

  const notifications =
    (errorBanner || copyBanner || undoBanner) ? (
      <div className="flex flex-col gap-2 px-4 pt-3">
        {errorBanner}
        {copyBanner}
        {undoBanner}
      </div>
    ) : null;

  // ---- The Fluent 4-pane IDE (shown ALWAYS — even with no tables yet, so the
  //      full layout/outline is visible; the center shows the onboarding until
  //      you paste). ---------------------------------------------------------
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
          <div className="ml-auto flex items-center gap-2">
            {/* ⚙ Document — GLOBAL body settings (all tables): body font, indent per
                nesting level, and reset-levels. Owned here (PasteInput holds the
                state), opened from a shared Popover. */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowDoc((v) => !v)}
                aria-expanded={showDoc}
                aria-label="Document body settings (all tables)"
                title="Body font, indent, and reset — applied to every section"
                className={BTN_SUBTLE}
              >
                &#9881; Document
              </button>
              <Popover open={showDoc} onClose={() => setShowDoc(false)}>
                <div className="flex w-60 flex-col gap-3">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted">
                    Document
                    <span className="ml-1.5 rounded-sm bg-accent-subtle px-1 text-[10px] font-semibold normal-case text-accent-text">
                      All tables
                    </span>
                  </div>
                  <label className="flex items-center justify-between gap-3 text-sm text-text-secondary">
                    Body font
                    <select
                      value={bodyFont}
                      onChange={(e) => changeBodyFont(e.target.value)}
                      aria-label="Document body font"
                      className={FIELD}
                    >
                      {FONTS.map((f) => (
                        <option key={f} value={f}>
                          {f}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label className="flex items-center justify-between gap-3 text-sm text-text-secondary">
                    Indent/level (in)
                    <input
                      type="text"
                      inputMode="decimal"
                      value={indentInput}
                      onChange={(e) =>
                        changeIndentInput(e.target.value.replace(/[^0-9.]/g, ""))
                      }
                      aria-label="Indent per nesting level, in inches"
                      className={`${FIELD} w-16`}
                    />
                  </label>
                  <button
                    type="button"
                    onClick={resetLevels}
                    disabled={levelStyles.length === 0}
                    title="Reset every level to the default look"
                    className="h-8 rounded border border-border-strong bg-surface px-2.5 text-xs font-medium text-text-secondary transition-colors hover:border-accent hover:text-accent-text disabled:pointer-events-none disabled:opacity-40"
                  >
                    Reset levels
                  </button>
                </div>
              </Popover>
            </div>
            {confirmClear ? (
              <span className="inline-flex items-center gap-1 rounded bg-surface-alt px-1.5">
                <span className="pl-1 text-xs text-muted">Clear all?</span>
                <button
                  type="button"
                  onClick={clearAll}
                  className="inline-flex h-8 items-center rounded px-2.5 text-sm font-semibold text-danger transition-colors hover:bg-[color:var(--danger-bg)]"
                >
                  Yes, clear
                </button>
                <button
                  type="button"
                  onClick={() => setConfirmClear(false)}
                  className={BTN_SUBTLE}
                >
                  Cancel
                </button>
              </span>
            ) : (
              <button
                type="button"
                onClick={() => setConfirmClear(true)}
                disabled={tables.length === 0}
                title="Remove every section (you can undo)"
                className={BTN_SUBTLE}
              >
                Clear all
              </button>
            )}
            <button
              type="button"
              onClick={copyAll}
              disabled={tables.length === 0}
              title="Copies every section as one Word doc (stacked in paste order)."
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
      </header>

      {notifications}
      {statusRegion}

      {/* WORKSPACE */}
      <div className="flex min-h-0 flex-1 overflow-x-auto">
        {/* LEFT: sections rail */}
        <SectionsRail
          tables={tables}
          activeId={activeTable?.id ?? null}
          onSelect={setActiveId}
          onRemove={removeTable}
          onReorder={reorderTable}
        />

        {/* CENTER: onboarding (empty) or the two command groups */}
        <section className="flex min-w-0 flex-1 flex-col gap-4 overflow-y-auto p-4">
          {tables.length === 0 ? (
            <div className="flex flex-col gap-5">
              <div className="rounded-lg border border-border bg-surface p-5 shadow-[var(--shadow-2)]">
                <h2 className="text-base font-semibold text-foreground">
                  Turn a wide spreadsheet into a Word-ready outline
                </h2>
                <p className="mt-1 text-sm text-muted">
                  Wide tables don&apos;t fit a Word page. Nest their columns into a
                  narrow outline that flows down the page instead.
                </p>
                <ol className="mt-3 flex flex-col gap-1.5 text-sm text-text-secondary">
                  {[
                    "Paste a wide Excel or Google Sheets table.",
                    "Arrange its columns into nested indent levels.",
                    "Copy all → paste into Word.",
                  ].map((step, i) => (
                    <li key={i} className="flex items-center gap-2.5">
                      <span className="grid h-5 w-5 shrink-0 place-items-center rounded-full bg-accent-subtle text-[11px] font-semibold text-accent-text">
                        {i + 1}
                      </span>
                      {step}
                    </li>
                  ))}
                </ol>
              </div>
              {pasteZone(true)}
              <div className="flex items-center justify-center gap-2 text-sm text-muted">
                <span>No spreadsheet handy?</span>
                <button
                  type="button"
                  onClick={addExample}
                  className="rounded border border-border-strong bg-surface px-3 py-1.5 text-sm font-semibold text-accent-text transition-colors hover:border-accent hover:bg-accent-subtle"
                >
                  Try an example
                </button>
              </div>
            </div>
          ) : (
            <>
              {pasteZone(false)}
              {activeTable && (
                <TableCard
                  key={activeTable.id}
                  table={activeTable}
                  onChange={(patch) => patchTable(activeTable.id, patch)}
                  headingStyleName={headingStyleName}
                  onHeadingStyleChange={changeHeadingStyleName}
                  title={titleInput}
                  onTitleChange={setTitle}
                  appearance={{
                    levelStyles,
                    onLevelChange: setLevel,
                  }}
                />
              )}
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
            {/* Scope toggle — only meaningful for the rendered preview with >1
                section: focus on the active one, or see the full stacked doc. */}
            {view === "rendered" && tables.length > 1 && (
              <div className="ml-auto inline-flex gap-0.5 rounded bg-surface-alt p-0.5 text-xs">
                <button
                  type="button"
                  aria-pressed={previewScope === "active"}
                  onClick={() => setPreviewScope("active")}
                  className={`rounded px-2.5 py-1 font-semibold transition-colors ${
                    previewScope === "active"
                      ? "bg-surface text-foreground shadow-[var(--shadow-2)]"
                      : "text-text-secondary hover:text-foreground"
                  }`}
                >
                  This section
                </button>
                <button
                  type="button"
                  aria-pressed={previewScope === "all"}
                  onClick={() => setPreviewScope("all")}
                  title="Preview every section stacked, exactly as Copy all produces it"
                  className={`rounded px-2.5 py-1 font-semibold transition-colors ${
                    previewScope === "all"
                      ? "bg-surface text-foreground shadow-[var(--shadow-2)]"
                      : "text-text-secondary hover:text-foreground"
                  }`}
                >
                  All sections
                </button>
              </div>
            )}
          </div>
          <div className="min-h-0 flex-1 overflow-y-auto p-4">
            {view === "json" ? (
              activeTable ? (
                <JsonPreview grid={activeTable.grid} />
              ) : null
            ) : previewScope === "all" && tables.length > 1 ? (
              <RenderedPreview
                sections={sectionHtmls}
                headingStyle={headingStyle}
                bodyFont={bodyFont}
                emptyHint="No section has any fields yet. Add fields in a section to build the outline."
              />
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

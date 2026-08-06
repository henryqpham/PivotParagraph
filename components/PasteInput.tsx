"use client";

import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
} from "react";
import { parseClipboard, parseFile } from "@/lib/parser";
import {
  buildWordHtml,
  htmlToPlainText,
  headingLevel,
  writeRichClipboard,
} from "@/lib/clipboard";
import type { HeadingStyle, LevelStyle } from "@/lib/clipboard";
import { loadSession, saveSession, clearSession } from "@/lib/persistence";
import {
  applyFieldSettingsByLevel,
  applyLabelsByLevel,
  buildStylePreset,
  parseStylePreset,
} from "@/lib/stylePreset";
import {
  tableToHtml,
  newTable,
  makeExampleTable,
  estimateSectionStats,
  dropEmptyColumns,
  DEFAULT_LEVEL,
  type LevelInput,
  type TableState,
} from "./tableModel";
import { TableCard, type TitleInput } from "./TableCard";
import { SectionsRail } from "./SectionsRail";
import { Popover } from "./Popover";
import { RenderedPreview } from "./RenderedPreview";
import { GridTable } from "./GridTable";
import { AddTableModal } from "./AddTableModal";
import { CommandPalette, type Command } from "./CommandPalette";

const MAX_TABLES = 100;

// Default TITLE look (its own shared style — Aptos 11 black, Microsoft 365's
// current default font, so the output reads like the source; only shows when
// the Heading dropdown is None).
const DEFAULT_TITLE: TitleInput = {
  font: "Aptos",
  sizeInput: "11",
  color: "#000000",
  bold: false,
  italic: false,
  underline: false,
};

/** The full workspace state we persist to localStorage (all JSON-serializable).
 *  Optional fields stay optional so pre-existing sessions/exports still load
 *  (defaults applied on read); `labelSep`/`bodyFont` are LEGACY leftovers of
 *  removed document-wide controls, ignored on read. */
type SessionSnapshot = {
  tables: TableState[];
  levelStyles: LevelInput[];
  /** LEGACY (ignored on read): the removed document-wide body font. */
  bodyFont?: string;
  /** LEGACY (ignored on read): the removed ⚙ Indent/level control —
   *  indentation is now DERIVED from the outline geometry (each level nests
   *  under its parent's text; see buildWordHtml / RenderedPreview). */
  indentInput?: string;
  headingStyleName: string;
  titleInput: TitleInput;
  activeId: string | null;
  labelSep?: string; // LEGACY (ignored)
  lineSpacing?: string;
  /** UI preference: opt-in drag-to-reorder on the Rows list (default off — the
   *  arrow buttons are the primary path). Persisted, but NOT a history-worthy
   *  edit (see historyChanged) and NOT part of a style preset. */
  dragRows?: boolean;
};

/** The post-copy confirmation shown near "Copy section" (and announced to a11y). */
type CopyNote =
  | { status: "ok"; hasHeading: boolean }
  | { status: "empty" }
  | { status: "error" };

// Full workspace undo/redo history. The oldest step silently drops once the cap is
// hit; a burst of rapid edits (typing, dragging a color) coalesces into ONE step
// via a short debounce so Ctrl+Z doesn't crawl character-by-character.
const HISTORY_CAP = 100;
const HISTORY_DEBOUNCE = 300;

// A history-worthy change = any workspace field EXCEPT pure section selection
// (`activeId`). Fields are compared by reference — every state update replaces the
// changed slice with a new object, so `!==` is an exact "did this actually change"
// test. Selecting a different tab shouldn't create an undo step, but the snapshot
// still carries activeId so undo restores the right focus.
function historyChanged(a: SessionSnapshot, b: SessionSnapshot): boolean {
  return (
    a.tables !== b.tables ||
    a.levelStyles !== b.levelStyles ||
    a.headingStyleName !== b.headingStyleName ||
    a.titleInput !== b.titleInput ||
    a.lineSpacing !== b.lineSpacing
    // `dragRows` is deliberately absent — a UI preference (like activeId), so
    // toggling it never creates an undo step, though it still persists.
  );
}


/** Allow-listed body line-spacing multipliers (the ⚙ Document control). */
const LINE_SPACINGS = [
  { value: "1", label: "Single" },
  { value: "1.15", label: "1.15 (default)" },
  { value: "1.5", label: "1.5" },
] as const;
const DEFAULT_LINE_SPACING = "1.15";

// ---- Fluent control recipes -----------------------------------------------
// Disabled = flat neutral (not faded blue): a 50%-opacity primary still reads as
// clickable; pointer-events-none also stops hover styles from firing on it.
const BTN_PRIMARY =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-accent-fg bg-accent transition-colors hover:bg-accent-hover active:bg-accent-pressed disabled:pointer-events-none disabled:bg-[color:color-mix(in_srgb,var(--muted)_14%,transparent)] disabled:text-muted";
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
  const [copyState, setCopyState] = useState<"idle" | "copied" | "error">(
    "idle",
  );
  const [view, setView] = useState<"rendered" | "table">("rendered");
  // Whether the top-band "Document" (global body settings) popover is open.
  const [showDoc, setShowDoc] = useState(false);
  // The "+ Add table" intake modal (paste / drop a file / example, previewed
  // before it commits) and the Ctrl+K command palette.
  const [showAddModal, setShowAddModal] = useState(false);
  const [showPalette, setShowPalette] = useState(false);
  // Mirrors `showAddModal` into a ref so the window paste listener can bail while
  // the modal owns paste (both listen on window; preventDefault alone wouldn't stop
  // the other from double-ingesting).
  const addModalOpenRef = useRef(false);
  useEffect(() => {
    addModalOpenRef.current = showAddModal;
  }, [showAddModal]);
  // Post-copy confirmation + Word-paste guidance (persists until the next edit).
  const [copyNote, setCopyNote] = useState<CopyNote | null>(null);
  // Two-step guard on the destructive "Clear all".
  const [confirmClear, setConfirmClear] = useState(false);
  // Full workspace undo/redo history (session-only, never persisted). Rather than
  // instrumenting every mutation, we OBSERVE the `snapshot` memo — which already
  // captures the whole workspace — and record a step whenever it meaningfully
  // changes. Held in a ref for race-free synchronous mutation across the debounce +
  // apply flow (`syncUndoAvail` mirrors availability into state for the buttons).
  // `past`/`future` are the classic two stacks;
  // `present` is the last-committed snapshot; `pending` is a pre-burst baseline
  // waiting for the debounce to commit.
  const historyRef = useRef<{
    past: SessionSnapshot[];
    present: SessionSnapshot | null;
    future: SessionSnapshot[];
  }>({ past: [], present: null, future: [] });
  const pendingRef = useRef<SessionSnapshot | null>(null);
  // Suppress recording while WE apply a snapshot (undo/redo/hydration), so those
  // don't get re-recorded as fresh edits.
  const applyingHistoryRef = useRef(false);
  const recordTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  // The history lives in a ref (race-free synchronous mutation), but the toolbar
  // buttons need reactive enabled-state — so we MIRROR "can undo / can redo" into
  // state and refresh it (via syncUndoAvail, defined below) at every mutation.
  // Reading the ref during render is disallowed here (react-hooks/refs); reading
  // this state is fine.
  const [undoAvail, setUndoAvail] = useState({ canUndo: false, canRedo: false });
  // Gate persistence writes until after the initial localStorage load, so we never
  // clobber saved work with the empty initial state.
  const [hydrated, setHydrated] = useState(false);
  const idRef = useRef(0);
  const copyTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const presetInputRef = useRef<HTMLInputElement>(null);
  // Latest snapshot + a hydration flag, kept in refs so the tab-hide flush can
  // save synchronously (from an event, without stale closure state) and never
  // overwrite saved work before the initial load has run.
  const snapshotRef = useRef<SessionSnapshot | null>(null);
  const hydratedRef = useRef(false);

  // The one body font every unpinned level inherits. The document-wide font
  // CONTROL was removed on feedback ("don't need whole document font control")
  // — the default is Aptos (Microsoft 365's standard) and a level's Look popover
  // can still pin any allow-listed font per level.
  const BODY_FONT = "Aptos";
  // Per-level BODY styling, shared across tables (the "level chart"). Sparse.
  const [levelStyles, setLevelStyles] = useState<LevelInput[]>([]);
  // Left-indent per nesting level (inches), clamped [0, 2].
  // Global Word heading style the TITLE maps to ("" = None). Driven by the
  // Section Title dropdown in TableCard.
  // "" = None: a new workspace applies NO Word-heading mapping until chosen.
  const [headingStyleName, setHeadingStyleName] = useState<string>("");
  // Shared TITLE look (its own controls in the Section Title group).
  const [titleInput, setTitleInput] = useState<TitleInput>(DEFAULT_TITLE);
  // Document-wide `Field name<sep>value` separator (allow-listed; default ": ").
  // Document-wide body line spacing as a string multiplier ("1"/"1.15"/"1.5").
  const [lineSpacing, setLineSpacing] = useState<string>(DEFAULT_LINE_SPACING);
  // Opt-in drag-to-reorder for the Rows list (⚙ Document). Arrows are default.
  const [dragRows, setDragRows] = useState<boolean>(false);

  // Apply a full workspace snapshot into state — shared by the mount-time
  // localStorage rehydration below AND by Import (a user-selected backup file), so
  // the two "load a snapshot" paths can never drift apart. Only ever calls stable
  // setState setters + writes a ref, so it has no reactive dependencies and is safe
  // as an empty-deps useCallback (satisfies exhaustive-deps without disabling it).
  const applySnapshot = useCallback((s: SessionSnapshot) => {
    setTables(s.tables);
    setLevelStyles(Array.isArray(s.levelStyles) ? s.levelStyles : []);
    setHeadingStyleName(s.headingStyleName ?? "");
    setTitleInput(s.titleInput ?? DEFAULT_TITLE);
    setLineSpacing(s.lineSpacing ?? DEFAULT_LINE_SPACING);
    setDragRows(s.dragRows === true);
    const restoredActive =
      s.tables.find((t) => t.id === s.activeId)?.id ?? s.tables[0]?.id ?? null;
    setActiveId(restoredActive);
    // Re-seed the id counter past the largest restored id so new pastes never
    // collide with a restored `t<N>`.
    idRef.current = s.tables.reduce((max, t) => {
      const n = parseInt(String(t.id).replace(/\D/g, ""), 10);
      return Number.isFinite(n) ? Math.max(max, n) : max;
    }, 0);
  }, []);

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
      // Restoring saved work is not an undoable "edit" — suppress recording so the
      // history doesn't start with an "undo back to empty" step.
      applyingHistoryRef.current = true;
      applySnapshot(data as SessionSnapshot);
    }
    hydratedRef.current = true;
    setHydrated(true);
    /* eslint-enable react-hooks/set-state-in-effect */
  }, [applySnapshot]);

  // The full workspace snapshot; one memo feeds both the debounced save and the
  // synchronous tab-hide flush so they can never disagree.
  const snapshot = useMemo<SessionSnapshot>(
    () => ({
      tables,
      levelStyles,
      headingStyleName,
      titleInput,
      activeId,
      lineSpacing,
      dragRows,
    }),
    [
      tables,
      levelStyles,
      headingStyleName,
      titleInput,
      activeId,
      lineSpacing,
      dragRows,
    ],
  );
  useEffect(() => {
    snapshotRef.current = snapshot;
  }, [snapshot]);

  // ---- Persistence: debounced save on any workspace change ----------------
  useEffect(() => {
    if (!hydrated) return;
    const handle = setTimeout(() => {
      // An empty workspace REMOVES the saved key entirely rather than persisting an
      // empty shell — so Clear all (or removing the last section) truly wipes the
      // cache instead of leaving a lingering entry behind.
      if (snapshot.tables.length === 0) clearSession();
      else saveSession(snapshot);
    }, 400);
    return () => clearTimeout(handle);
  }, [hydrated, snapshot]);

  // ---- Persistence: flush immediately when the tab is hidden or closed, so a
  //      change made inside the 400ms debounce window isn't lost. Gated on the
  //      hydration ref so an early hide never overwrites saved work with the
  //      empty initial state. ------------------------------------------------
  useEffect(() => {
    const flush = () => {
      if (!hydratedRef.current || !snapshotRef.current) return;
      const s = snapshotRef.current;
      // Same rule as the debounced save: an empty workspace removes the key.
      if (s.tables.length === 0) clearSession();
      else saveSession(s);
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

  // ---- Undo/redo history: observe `snapshot`, record one step per meaningful
  //      change, coalescing rapid bursts via a short debounce. ------------------
  // Recompute the mirrored can-undo/can-redo state from the ref. Called (never in
  // render) at every history mutation so the toolbar buttons stay live.
  const syncUndoAvail = useCallback(() => {
    const h = historyRef.current;
    setUndoAvail({
      canUndo: h.past.length > 0 || pendingRef.current !== null,
      canRedo: h.future.length > 0,
    });
  }, []);

  // Commit the pending pre-burst baseline into `past` (clearing any redo future).
  const commitPending = useCallback(() => {
    if (recordTimer.current) {
      clearTimeout(recordTimer.current);
      recordTimer.current = null;
    }
    const before = pendingRef.current;
    if (before === null) return;
    pendingRef.current = null;
    const h = historyRef.current;
    h.past = [...h.past.slice(-(HISTORY_CAP - 1)), before];
    h.future = [];
    h.present = snapshotRef.current;
    syncUndoAvail();
  }, [syncUndoAvail]);

  useEffect(() => {
    const h = historyRef.current;
    if (!hydrated) {
      h.present = snapshot;
      return;
    }
    if (applyingHistoryRef.current) {
      // This change came from applying an undo/redo (or hydration): adopt it as the
      // new baseline, but don't log it as a fresh edit.
      applyingHistoryRef.current = false;
      h.present = snapshot;
      return;
    }
    if (h.present === null) {
      h.present = snapshot;
      return;
    }
    if (!historyChanged(h.present, snapshot)) {
      // Selection-only (activeId) change: keep the baseline current so a later real
      // edit records the right pre-edit state, but don't create an undo step.
      if (pendingRef.current === null) h.present = snapshot;
      return;
    }
    // A real edit: capture the pre-burst baseline once, then (re)arm the debounce so
    // a rapid burst (typing, dragging a color) commits as a single step.
    if (pendingRef.current === null) {
      pendingRef.current = h.present;
      syncUndoAvail(); // enable Undo immediately, before the commit fires
    }
    if (recordTimer.current) clearTimeout(recordTimer.current);
    recordTimer.current = setTimeout(commitPending, HISTORY_DEBOUNCE);
  }, [hydrated, snapshot, commitPending, syncUndoAvail]);

  // Clear the copy confirmation + any pending clear-confirm when the user edits.
  // Called from the mutation handlers (an event-time call, NOT an effect) so the
  // note reads as "…until you change something" without a cascading render.
  function noteEdit() {
    setCopyNote(null);
    setConfirmClear(false);
    // A stale paste-error banner is no longer relevant once the user does anything
    // else, so clear it here too (mirrors how the copy note is cleared on edit).
    setError(null);
    // NOTE: undo/redo history is NOT touched here — it's driven by the snapshot
    // observer below, so ordinary edits ADD to history rather than clearing it.
  }

  // The pending "Clear all?" confirm auto-cancels if left untouched.
  useEffect(() => {
    if (!confirmClear) return;
    const handle = setTimeout(() => setConfirmClear(false), 4000);
    return () => clearTimeout(handle);
  }, [confirmClear]);

  const headingStyle = useMemo<HeadingStyle>(() => {
    const clampPt = (s: string, fallback: number) => {
      const n = parseInt(s, 10);
      return Number.isFinite(n) && n >= 1 && n <= 72 ? n : fallback;
    };
    const levels: LevelStyle[] = Array.from({ length: 9 }, (_, i) => {
      const ls = levelStyles[i];
      return {
        color: ls?.color ?? DEFAULT_LEVEL.color,
        // "" / unset = INHERIT the document-wide Body font — this is what makes
        // the ⚙ Document Body font control actually drive the body rows (both
        // the preview rules and the pasted inline styles read these levels); a
        // level's Look popover pins a specific font by setting a non-empty one.
        font: ls?.font || BODY_FONT,
        size: clampPt(
          ls?.sizeInput ?? DEFAULT_LEVEL.sizeInput,
          parseInt(DEFAULT_LEVEL.sizeInput, 10),
        ),
        bold: ls?.bold ?? false,
        italic: ls?.italic ?? false,
        underline: ls?.underline ?? false,
      };
    });
    const parsedSpacing = parseFloat(lineSpacing);
    return {
      levels,
      headingStyleName,
      titleStyle: {
        font: titleInput.font,
        size: clampPt(titleInput.sizeInput, 11),
        color: titleInput.color,
        bold: titleInput.bold,
        italic: titleInput.italic,
        underline: titleInput.underline,
      },
      lineSpacing: Number.isFinite(parsedSpacing)
        ? Math.min(3, Math.max(1, parsedSpacing))
        : 1.15,
    };
  }, [
    levelStyles,
    headingStyleName,
    titleInput,
    lineSpacing,
  ]);

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
  function changeLineSpacing(v: string) {
    noteEdit();
    if (LINE_SPACINGS.some((s) => s.value === v)) setLineSpacing(v);
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
    // Strip fully-empty spacer/padding columns on the way in so they never clutter
    // the Add-fields pool (automatic, no UI).
    const cleaned = dropEmptyColumns(grid);
    // A new section always lands BLANK (pivotLevels: []) — building the outline by
    // adding fields one at a time is the deliberate flow. (Two assists were tried
    // here and removed on feedback: an auto-arrange heuristic that pre-placed
    // every column, and a header-match "reuse that arrangement" banner. The
    // STYLE PRESET covers the recurring-report need now — replay the formatting,
    // rebuild the structure.)
    setTables((prev) => [...prev, newTable(id, cleaned)]);
    setActiveId(id);
    setError(null);
    setStatus(`Added section ${tables.length + 1}.`);
    return true;
  }

  function ingestClipboard(data: DataTransfer) {
    try {
      const rows = parseClipboard(data);
      // Require at least two cells so a stray Ctrl/Cmd+V of ordinary text (which
      // parses as a lone 1x1 cell) can't silently pollute the workspace with a junk
      // section — a real table always has more than one cell.
      const cells = rows.reduce((n, r) => n + r.length, 0);
      if (cells < 2) {
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
      // The Add-table modal owns paste while it is open (its own window listener +
      // preview-before-commit), so bail here to avoid a double ingest.
      if (addModalOpenRef.current) return;
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

  // Copy-anywhere: Ctrl/Cmd+Enter copies the active section — the mouse-free way
  // OUT, pairing with paste-anywhere as the mouse-free way IN. Same latest-handler
  // ref pattern so the listener mounts once. (copySection is a hoisted function
  // declaration, so referencing it here is safe.)
  const copyShortcutRef = useRef<() => void>(() => {});
  useEffect(() => {
    copyShortcutRef.current = () => void copySection();
  });
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (
        e.key === "Enter" &&
        (e.ctrlKey || e.metaKey) &&
        !e.shiftKey &&
        !e.altKey
      ) {
        e.preventDefault();
        copyShortcutRef.current();
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Ctrl/Cmd+K toggles the command palette — a global launcher over every action,
  // and the home that surfaces the app's otherwise-invisible shortcuts.
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (
        (e.ctrlKey || e.metaKey) &&
        !e.shiftKey &&
        !e.altKey &&
        e.key.toLowerCase() === "k"
      ) {
        e.preventDefault();
        setShowPalette((v) => !v);
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  function patchTable(id: string, patch: Partial<TableState>) {
    noteEdit();
    setTables((ts) => ts.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  }

  // Move the section at `from` to index `to` (drag-drop reorder from the rail).
  function moveTable(from: number, to: number) {
    noteEdit();
    setTables((ts) => {
      if (
        from < 0 ||
        from >= ts.length ||
        to < 0 ||
        to >= ts.length ||
        from === to
      )
        return ts;
      const next = [...ts];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      return next;
    });
  }

  // Undo: first FLUSH any in-flight edit (so a Ctrl+Z within the debounce window
  // still reverts the change you just made), then pop `past` → restore it, pushing
  // the state being replaced onto `future` so Redo can step forward. All history
  // mutation is synchronous on the ref; applySnapshot drives the actual state, and
  // applyingHistoryRef stops that from being re-recorded as a new edit.
  function doUndo() {
    commitPending();
    const h = historyRef.current;
    if (h.past.length === 0) return;
    const prev = h.past[h.past.length - 1];
    h.past = h.past.slice(0, -1);
    if (h.present) h.future = [...h.future.slice(-(HISTORY_CAP - 1)), h.present];
    h.present = prev;
    applyingHistoryRef.current = true;
    setCopyNote(null);
    applySnapshot(prev);
    syncUndoAvail();
    setStatus("Undid the last change.");
  }

  // The mirror image of doUndo.
  function doRedo() {
    commitPending();
    const h = historyRef.current;
    if (h.future.length === 0) return;
    const next = h.future[h.future.length - 1];
    h.future = h.future.slice(0, -1);
    if (h.present) h.past = [...h.past.slice(-(HISTORY_CAP - 1)), h.present];
    h.present = next;
    applyingHistoryRef.current = true;
    setCopyNote(null);
    applySnapshot(next);
    syncUndoAvail();
    setStatus("Redid the change.");
  }

  // Ctrl/Cmd+Z (undo) and Ctrl+Y / Ctrl/Cmd+Shift+Z (redo) anywhere except inside a
  // text field (matching the paste-anywhere guard below). Kept in refs — same
  // reason as ingestRef below — so the mount-only listener never closes over a
  // stale doUndo/doRedo from the render it was attached in.
  const doUndoRef = useRef(doUndo);
  const doRedoRef = useRef(doRedo);
  useEffect(() => {
    doUndoRef.current = doUndo;
    doRedoRef.current = doRedo;
  });
  useEffect(() => {
    function onKeyDown(e: KeyboardEvent) {
      if (!(e.ctrlKey || e.metaKey)) return;
      const key = e.key.toLowerCase();
      const isRedo = key === "y" || (key === "z" && e.shiftKey);
      const isUndo = key === "z" && !e.shiftKey;
      if (!isRedo && !isUndo) return;
      const el = document.activeElement as HTMLElement | null;
      if (
        el &&
        (el.tagName === "INPUT" ||
          el.tagName === "TEXTAREA" ||
          el.tagName === "SELECT" ||
          el.isContentEditable)
      ) {
        return; // let the focused field handle its own undo/redo
      }
      e.preventDefault();
      if (isRedo) doRedoRef.current();
      else doUndoRef.current();
    }
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  // ---- Style preset: export/import the FORMATTING only (the "parser config").
  //      Everything in it is global or LEVEL-keyed — never column-keyed — so a
  //      preset saved from this week's report replays on next week's completely
  //      different table. Import applies the globals plus every section's
  //      level-keyed settings (all sections, per the user's chosen scope); it
  //      never touches grids, structure (pivotLevels), titles, or per-column
  //      label/sort settings. ------------------------------------------------
  function exportStylePreset() {
    const payload = buildStylePreset(
      {
        levelStyles,
        titleInput,
        headingStyleName,
        lineSpacing,
      },
      activeTable,
    );
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `pivotparagraph-style-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    setShowDoc(false);
  }

  function handlePresetFile(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-selecting the same file to re-trigger onChange
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = parseStylePreset(JSON.parse(String(reader.result)));
        if (!data) throw new Error("not a style preset");
        setCopyNote(null);
        setError(null);
        setConfirmClear(false);
        // Checkpoint any in-flight edit so the whole preset application is ONE
        // clean undo step (the snapshot observer records the pre-import state).
        commitPending();
        // Globals — through the raw setters, with the same allow-list guards the
        // ⚙ controls use for the two enumerated fields.
        setLevelStyles(Array.isArray(data.levelStyles) ? data.levelStyles : []);
        setTitleInput(data.titleInput);
        setHeadingStyleName(data.headingStyleName);
        if (LINE_SPACINGS.some((s) => s.value === data.lineSpacing))
          setLineSpacing(data.lineSpacing as string);
        // Level-keyed settings onto EVERY section. structuredClone per table so
        // sections share no references; `markers: undefined` drops the legacy
        // fused field (markerSpecs is authoritative). The Rows-card label
        // emphasis re-applies by LEVEL POSITION to each section's placed fields
        // (applyLabelsByLevel) — the column-keyed part the preset can't carry raw.
        setTables((ts) =>
          ts.map((t) => ({
            ...t,
            ...structuredClone({
              markerSpecs: data.markerSpecs,
              numbering: data.numbering,
              headingLevels: data.headingLevels,
              headingRanks: data.headingRanks,
              breakAfter: data.breakAfter,
              gapAfter: data.gapAfter,
              pageBreakBefore: data.pageBreakBefore,
              fieldLabels: applyLabelsByLevel(t, data.labelsByLevel),
              ...applyFieldSettingsByLevel(
                t,
                data.sepsByLevel,
                data.looksByLevel,
                data.breaksByLevel,
              ),
            }),
            markers: undefined,
          })),
        );
        setShowDoc(false);
        setStatus(
          "Style preset applied to all sections. Undo available (Ctrl+Z).",
        );
      } catch {
        setError("Couldn't read that file as a style preset.");
      }
    };
    reader.onerror = () => setError("Couldn't read that file.");
    reader.readAsText(file);
  }

  // Deep-clone a section under a fresh id, inserted right after the original, and
  // make the copy active. Serves the "several near-identical tables" workflow
  // (same headers, different month/region). structuredClone handles the nested
  // pivotLevels/fieldLabels/numbering so the copy shares no references.
  function duplicateTable(id: string) {
    if (tables.length >= MAX_TABLES) {
      setError(
        `Reached the ${MAX_TABLES}-section limit. Remove a section to add another.`,
      );
      return;
    }
    const src = tables.find((t) => t.id === id);
    if (!src) return;
    setCopyNote(null);
    commitPending(); // clean undo step
    const newId = `t${++idRef.current}`;
    const clone: TableState = { ...structuredClone(src), id: newId };
    setTables((ts) => {
      const idx = ts.findIndex((t) => t.id === id);
      const next = [...ts];
      next.splice(idx + 1, 0, clone);
      return next;
    });
    setActiveId(newId);
    setStatus("Section duplicated.");
  }

  function removeTable(id: string) {
    setCopyNote(null);
    // Checkpoint any in-flight edit so this removal is its own clean undo step; the
    // observer records the pre-removal state automatically.
    commitPending();
    setActiveId((curr) => {
      if (curr !== id) return curr;
      const idx = tables.findIndex((t) => t.id === id);
      const remaining = tables.filter((t) => t.id !== id);
      if (remaining.length === 0) return null;
      return remaining[Math.min(idx, remaining.length - 1)].id;
    });
    setTables((ts) => ts.filter((t) => t.id !== id));
    setStatus("Section removed. Undo available (Ctrl+Z).");
  }

  function clearAll() {
    setCopyNote(null);
    // Checkpoint any in-flight edit so the clear is its own undo step.
    commitPending();
    setTables([]);
    setActiveId(null);
    setError(null);
    setConfirmClear(false);
    // Remove the localStorage entry right away (the debounced effect would also do
    // this once state settles, but wipe immediately so the cache is gone the moment
    // you clear, not 400ms later).
    clearSession();
    setStatus("All sections cleared. Undo available (Ctrl+Z).");
  }

  // Flash the transient copy-button state, replacing any prior 2s reset timer so a
  // rapid re-copy doesn't get its "Copied!" confirmation cut short by the earlier
  // click's timer firing.
  function flashCopyState(s: "copied" | "error") {
    setCopyState(s);
    if (copyTimer.current) clearTimeout(copyTimer.current);
    copyTimer.current = setTimeout(() => setCopyState("idle"), 2000);
  }

  // Copy ONLY the active section — each section is its own standalone Word doc.
  async function copySection() {
    if (!activeTable) return;
    const titleLvl = headingLevel(headingStyleName);
    const html = tableToHtml(activeTable, titleLvl);
    if (html === "") {
      setCopyNote({ status: "empty" });
      return;
    }
    // `writeRichClipboard` tries the async Clipboard API and falls back to the
    // legacy execCommand route, so Copy still works where the modern API is
    // refused — notably the single-file build opened straight from disk
    // (origin `null`, not a secure context).
    const ok = await writeRichClipboard(
      buildWordHtml(html, headingStyle, BODY_FONT),
      htmlToPlainText(html),
    );
    if (!ok) {
      setCopyNote({ status: "error" });
      flashCopyState("error");
      return;
    }
    // Only claim a Word-heading mapping when one is ACTUALLY emitted: the title
    // maps to a heading only if a Heading style is chosen AND the title has text
    // (a blank title renders no `ws-title`), and a body level maps only if it's
    // both marked AND actually exists in the structure. Otherwise the "Use
    // Destination Styles" guidance would be misleading when nothing actually
    // maps (a chosen style with a blank title emits no heading, for instance).
    const titleIsHeading =
      headingStyleName !== "" && activeTable.sectionTitle.trim() !== "";
    const bodyHasHeading = (activeTable.headingLevels ?? []).some(
      (h, i) => h && i < activeTable.pivotLevels.length,
    );
    setCopyNote({ status: "ok", hasHeading: titleIsHeading || bodyHasHeading });
    flashCopyState("copied");
  }

  const atLimit = tables.length >= MAX_TABLES;
  const activeTable = tables.find((t) => t.id === activeId) ?? tables[0] ?? null;

  const titleLevel = headingLevel(headingStyleName);
  // The active section's rendered fragment — the single source for both the right-
  // pane preview and (recomputed identically in copySection) the copy, so the two
  // can't drift.
  const activeHtml = useMemo(
    () => (activeTable ? tableToHtml(activeTable, titleLevel) : ""),
    [activeTable, titleLevel],
  );
  // Rows/levels/rough-page-count readout for the active section — a quick at-a-
  // glance answer to "does this still fit a page," shown next to the preview tabs.
  const activeStats = useMemo(
    () => (activeTable ? estimateSectionStats(activeTable, activeHtml) : null),
    [activeTable, activeHtml],
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
            Nothing to copy yet — add fields to this section first.
          </span>
        ) : (
          <span>
            <strong>Section copied.</strong> In Word, press{" "}
            <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
              Ctrl
            </kbd>
            +
            <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
              V
            </kbd>{" "}
            {copyNote.hasHeading ? (
              <>
                → <strong>Use Destination Styles</strong>{" "}= your
                document&apos;s real Heading styles (Navigation pane +
                numbering); <strong>Keep Source Formatting</strong>{" "}= this
                exact preview look. Avoid <em>Merge Formatting</em>{" "}(strips
                the heading mapping).
              </>
            ) : (
              <>
                → choose <strong>Keep Source Formatting</strong>{" "}to keep
                this exact look.
              </>
            )}
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

  // Undo/Redo now live as persistent toolbar buttons in the top band (see below)
  // rather than a transient banner — their disabled state + tooltip communicate
  // availability, and setStatus(...) calls (removeTable/clearAll/doUndo/doRedo)
  // keep screen readers informed via the sr-only status region below.
  const statusRegion = (
    <p className="sr-only" role="status" aria-live="polite">
      {status}
    </p>
  );

  const notifications =
    errorBanner || copyBanner ? (
      <div className="flex flex-col gap-2 px-4 pt-3">
        {errorBanner}
        {copyBanner}
      </div>
    ) : null;

  // Read the mirrored availability state (kept in sync by syncUndoAvail at every
  // history mutation) — NOT the ref, which can't be read during render.
  const { canUndo, canRedo } = undoAvail;

  // The Ctrl+K command palette's vocabulary — every top-level action plus a
  // "Go to section" entry per table. Routes through the existing handlers, so all
  // of it inherits undo/redo and the confirmations. Disabled entries stay visible
  // (the launcher's menu is stable) but unselectable.
  const commands: Command[] = [
    {
      id: "add",
      label: "Add a table…",
      section: "Add",
      keywords: "paste import file upload new section",
      run: () => setShowAddModal(true),
    },
    {
      id: "example",
      label: "Load an example section",
      section: "Add",
      keywords: "demo sample try",
      run: addExample,
    },
    {
      id: "copy",
      label: "Copy section to clipboard",
      section: "This section",
      shortcut: "Ctrl+Enter",
      keywords: "export word clipboard",
      disabled: !activeTable,
      run: () => void copySection(),
    },
    {
      id: "dup",
      label: "Duplicate this section",
      section: "This section",
      keywords: "clone",
      disabled: !activeTable,
      run: () => activeTable && duplicateTable(activeTable.id),
    },
    {
      id: "remove",
      label: "Remove this section",
      section: "This section",
      keywords: "delete",
      disabled: !activeTable,
      run: () => activeTable && removeTable(activeTable.id),
    },
    {
      id: "undo",
      label: "Undo",
      section: "Edit",
      shortcut: "Ctrl+Z",
      disabled: !canUndo,
      run: doUndo,
    },
    {
      id: "redo",
      label: "Redo",
      section: "Edit",
      shortcut: "Ctrl+Y",
      disabled: !canRedo,
      run: doRedo,
    },
    {
      id: "doc",
      label: "Document settings",
      section: "Edit",
      keywords: "font spacing indent body separator backup export import",
      run: () => setShowDoc(true),
    },
    {
      id: "preset-export",
      label: "Export style preset",
      section: "Edit",
      hint: "save the formatting as .json",
      keywords: "config save style levels markers headings download",
      run: exportStylePreset,
    },
    {
      id: "preset-import",
      label: "Import style preset…",
      section: "Edit",
      hint: "apply saved formatting to all sections",
      keywords: "config load style levels markers headings apply replay",
      run: () => presetInputRef.current?.click(),
    },
    {
      id: "clear",
      label: "Clear all sections",
      section: "Edit",
      keywords: "delete reset",
      disabled: tables.length === 0,
      run: () => setConfirmClear(true),
    },
    ...tables.map((t, i) => ({
      id: `go-${t.id}`,
      label: `Go to ${t.sectionTitle.trim() || `Section ${i + 1}`}`,
      section: "Go to section",
      keywords: "switch select navigate",
      run: () => setActiveId(t.id),
    })),
  ];

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
              P
            </span>
            <h1 className="text-sm">PivotParagraph</h1>
          </span>
          {/* Plain count, not pagination; surface the 100-section cap only when
              it's actually close enough to matter. */}
          {tables.length > 0 && (
            <span className="text-xs text-muted">
              {tables.length >= MAX_TABLES - 10
                ? `${tables.length} of ${MAX_TABLES} sections`
                : `${tables.length} section${tables.length === 1 ? "" : "s"}`}
            </span>
          )}
          {/* Undo/Redo — persistent toolbar buttons covering EVERY workspace change
              (reorder, fonts, bold, indent, markers, title, add/remove, clear,
              import). Disabled + greyed out when there's nothing to undo/redo. Kept
              on the LEFT, away from the destructive/primary actions on the right. */}
          <div className="ml-2 flex items-center gap-0.5 border-l border-border pl-2">
            <button
              type="button"
              onClick={doUndo}
              disabled={!canUndo}
              aria-label="Undo"
              title={canUndo ? "Undo (Ctrl+Z)" : "Nothing to undo"}
              className="grid h-8 w-8 place-items-center rounded text-base text-text-secondary transition-colors hover:bg-surface-alt hover:text-foreground disabled:cursor-not-allowed disabled:opacity-30 disabled:hover:bg-transparent"
            >
              &#8630;
            </button>
            <button
              type="button"
              onClick={doRedo}
              disabled={!canRedo}
              aria-label="Redo"
              title={canRedo ? "Redo (Ctrl+Y)" : "Nothing to redo"}
              className="grid h-8 w-8 place-items-center rounded text-base text-text-secondary transition-colors hover:bg-surface-alt hover:text-foreground disabled:cursor-not-allowed disabled:opacity-30 disabled:hover:bg-transparent"
            >
              &#8631;
            </button>
          </div>
          <div className="ml-auto flex items-center gap-2">
            {/* (The Add-table button lives in the Sections rail — sections are
                made where they live, the layers-panel convention. The command
                palette is keyboard-only: Ctrl/Cmd+K.) */}
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
                  {/* (The document-wide Body font + Label separator controls
                      were removed on feedback — the font is Aptos with per-level
                      pins in each Look popover, and label separators are now
                      per-level free text in the Body Text Definition matrix.) */}
                  <label className="flex items-center justify-between gap-3 text-sm text-text-secondary">
                    Line spacing
                    <select
                      value={lineSpacing}
                      onChange={(e) => changeLineSpacing(e.target.value)}
                      aria-label="Body line spacing"
                      className={FIELD}
                    >
                      {LINE_SPACINGS.map((s) => (
                        <option key={s.value} value={s.value}>
                          {s.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  {/* No Indent-per-level control: indentation is DERIVED
                      from the outline geometry — each level's line nests under
                      its parent's text, sized by the measured number/marker
                      widths (0.25in steps when a level has no lead). */}
                  <button
                    type="button"
                    onClick={resetLevels}
                    disabled={levelStyles.length === 0}
                    title="Reset every level to the default look"
                    className="h-8 rounded border border-border-strong bg-surface px-2.5 text-xs font-medium text-text-secondary transition-colors hover:border-accent hover:text-accent-text disabled:pointer-events-none disabled:opacity-40"
                  >
                    Reset levels
                  </button>
                  {/* Style preset — the FORMATTING alone (level looks, markers,
                      numbering, headings, blank lines, fonts), level-keyed so it
                      replays on any document. The recurring-report answer for
                      "same house style, different table every week". */}
                  <div className="flex flex-col gap-1.5 border-t border-border pt-3">
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-muted">
                      Style preset
                    </span>
                    <div className="flex gap-1.5">
                      <button
                        type="button"
                        onClick={exportStylePreset}
                        title="Save the current formatting (levels, markers, headings, fonts) as a .json preset"
                        className="h-8 flex-1 rounded border border-border-strong bg-surface px-2.5 text-xs font-medium text-text-secondary transition-colors hover:border-accent hover:text-accent-text"
                      >
                        Export
                      </button>
                      <button
                        type="button"
                        onClick={() => presetInputRef.current?.click()}
                        title="Apply a saved style preset to every section (undoable)"
                        className="h-8 flex-1 rounded border border-border-strong bg-surface px-2.5 text-xs font-medium text-text-secondary transition-colors hover:border-accent hover:text-accent-text"
                      >
                        Import&hellip;
                      </button>
                    </div>
                  </div>
                </div>
              </Popover>
              <input
                ref={presetInputRef}
                type="file"
                accept="application/json"
                onChange={handlePresetFile}
                className="hidden"
                aria-hidden="true"
                tabIndex={-1}
              />
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
              onClick={copySection}
              disabled={!activeTable}
              title="Copies the active section as its own Word doc, ready to paste (Ctrl+Enter)"
              className={BTN_PRIMARY}
            >
              {copyState === "copied"
                ? "Copied!"
                : copyState === "error"
                  ? "Copy failed"
                  : "Copy section"}
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
          onReorder={moveTable}
          onDuplicate={duplicateTable}
          onAdd={() => setShowAddModal(true)}
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
                    "Copy section → paste into Word.",
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
              {/* No in-body add affordance — the rail's ＋ Add section is THE
                  entry point (its hint also teaches paste-anywhere). */}
            </div>
          ) : (
            <>
              {activeTable && (
                <TableCard
                  key={activeTable.id}
                  table={activeTable}
                  onChange={(patch) => patchTable(activeTable.id, patch)}
                  headingStyleName={headingStyleName}
                  onHeadingStyleChange={changeHeadingStyleName}
                  title={titleInput}
                  onTitleChange={setTitle}
                  dragRows={dragRows}
                  onDragRowsChange={setDragRows}
                  appearance={{
                    levelStyles,
                    onLevelChange: setLevel,
                    bodyFont: BODY_FONT,
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
              {(
                [
                  { key: "rendered", label: "Preview" },
                  { key: "table", label: "Table" },
                ] as const
              ).map((t) => (
                <button
                  key={t.key}
                  type="button"
                  aria-pressed={view === t.key}
                  onClick={() => setView(t.key)}
                  title={
                    t.key === "table"
                      ? "The pasted table, as-is (before the outline)"
                      : undefined
                  }
                  className={`rounded px-2.5 py-1 font-semibold transition-colors ${
                    view === t.key
                      ? "bg-surface text-foreground shadow-[var(--shadow-2)]"
                      : "text-text-secondary hover:text-foreground"
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>
            {activeStats && (
              <span
                className="ml-auto whitespace-nowrap text-xs tabular-nums text-muted"
                title="Rough estimate — long values wrap to more lines than this counts."
              >
                {activeStats.rows} row{activeStats.rows === 1 ? "" : "s"} ·{" "}
                {activeStats.levels} level{activeStats.levels === 1 ? "" : "s"}
                {activeStats.pages > 0 && (
                  <>
                    {" "}
                    · ~{activeStats.pages} page
                    {activeStats.pages === 1 ? "" : "s"}
                  </>
                )}
              </span>
            )}
          </div>
          <div className="min-h-0 flex-1 overflow-y-auto p-4">
            {view === "table" ? (
              activeTable ? (
                <GridTable
                  grid={activeTable.grid}
                  headerRow={activeTable.headerRow ?? 0}
                />
              ) : (
                <p className="text-sm text-foreground/60">
                  Paste a table to see it here, exactly as it came from Excel.
                </p>
              )
            ) : (
              <RenderedPreview
                html={activeHtml}
                headingStyle={headingStyle}
                bodyFont={BODY_FONT}
                emptyHint="Add fields in the center to build the outline. Stack fields at one indent level to show them together (like an Excel pivot's Row Labels)."
              />
            )}
          </div>
        </section>
      </div>

      {/* ＋ Add-table intake modal (paste / drop a file / example, previewed
          before commit) and the Ctrl+K command palette — both portal to <body>. */}
      <AddTableModal
        open={showAddModal}
        onClose={() => setShowAddModal(false)}
        onCommit={(grid) => {
          ingestGrid(grid);
        }}
        parseClipboardData={parseClipboard}
        parseFile={parseFile}
        onUseExample={() => {
          addExample();
          setShowAddModal(false);
        }}
        atLimit={atLimit}
        limitMessage={`Reached the ${MAX_TABLES}-section limit. Remove a section to add another.`}
      />
      <CommandPalette
        open={showPalette}
        onClose={() => setShowPalette(false)}
        commands={commands}
        placeholder="Search commands…"
      />

      {/* Version badge — build-time git-derived (appVersion.ts): "1.<commit
          count>", hash in the tooltip. Fixed bottom-right, out of every pane's
          flow; identical in the Next and single-file builds. */}
      <span
        title={`PivotParagraph version ${process.env.NEXT_PUBLIC_APP_VERSION ?? "dev"} — build-stamped from git`}
        className="fixed bottom-1.5 right-3 z-30 select-none text-[10px] tabular-nums text-muted/80"
      >
        v{(process.env.NEXT_PUBLIC_APP_VERSION ?? "dev").split(" · ")[0]}
      </span>
    </div>
  );
}

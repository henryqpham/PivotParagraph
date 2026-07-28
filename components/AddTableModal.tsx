"use client";

import {
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type ChangeEvent,
  type DragEvent,
  type ReactPortal,
} from "react";
import { createPortal } from "react-dom";
import type { Grid } from "@/lib/types";

/**
 * "Add a table" — a Dropbox-style intake modal that lets the user PREVIEW what
 * the parser read before it becomes a section. The app's original affordance was
 * a passive dashed paste-zone (see `PasteInput.pasteZone`) wired to a global paste
 * listener: a table appeared with no confirmation, so a mis-copy (one stray cell,
 * a non-table selection) silently became a section. This modal makes intake
 * DELIBERATE and legible — you paste / drop / browse, SEE a "read N rows × M
 * columns" summary plus a mini-table, and only THEN commit — while keeping the
 * paste-anywhere convenience (the window paste listener) intact while it is open.
 *
 * It is a pure UI shell over the parent's pipeline: the caller injects the two
 * parse functions (`parseClipboardData` = the app's `parseClipboard`, `parseFile`
 * = the async file parser) and the commit/example callbacks, so this component
 * owns no parsing or ingest logic of its own. `onCommit(grid)` hands the RAW grid
 * back to the parent, which runs the real ingest (dropEmptyColumns, the section
 * cap, the arrangement-reuse hint); the modal just closes afterward.
 *
 * Two phases, driven by whether a grid has been read:
 *  - PHASE 1 (INTAKE): nothing read yet — drop well + Browse + live paste
 *    listener + "Try an example".
 *  - PHASE 2 (PREVIEW): a grid is in hand — summary + mini-table + Add / Clear.
 *
 * Structurally it mirrors `Popover`: portaled to `document.body`, dismissed by an
 * Escape keypress or a click on the dimmed full-screen backdrop. It is a CENTERED
 * modal (not anchored to a trigger), so it uses flex-centering instead of the
 * Popover's imperative measure-and-place.
 */

// The app's Fluent control recipes, duplicated here because PasteInput does not
// export them (they are module-private). Keep these in sync with PasteInput's
// BTN_PRIMARY / BTN_SUBTLE if that palette ever changes.
const BTN_PRIMARY =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-accent-fg bg-accent transition-colors hover:bg-accent-hover active:bg-accent-pressed disabled:pointer-events-none disabled:bg-[color:color-mix(in_srgb,var(--muted)_14%,transparent)] disabled:text-muted";
const BTN_SUBTLE =
  "inline-flex h-8 items-center gap-1.5 rounded px-3 text-sm font-semibold text-text-secondary bg-transparent transition-colors hover:bg-surface-alt hover:text-foreground disabled:cursor-not-allowed disabled:opacity-50";

// The single inline error shown for any unreadable input (a too-small paste, a
// non-table drop, or a parse failure) — one message so paste / drop / browse all
// fail identically.
const READ_ERROR =
  "Couldn't read a table from that — copy a cell range from Excel/Sheets, or drop an .xlsx/.csv.";

// A parse must yield at least this many cells to count as a table; a lone cell (or
// an empty grid) is almost always a mis-copy, not a range worth restructuring.
const MIN_CELLS = 2;


/** Render a cell as display text (mirrors GridTable: null/undefined -> ""). */
const cellText = (c: unknown): string => (c == null ? "" : String(c));

export function AddTableModal({
  open,
  onClose,
  onCommit,
  parseClipboardData,
  parseFile,
  onUseExample,
  atLimit,
  limitMessage,
}: {
  open: boolean;
  onClose: () => void;
  /** Commit the parsed grid as a new section. The parent runs its ingest pipeline
   *  (dropEmptyColumns, MAX_TABLES guard, arrange hint). The modal calls this then onClose. */
  onCommit: (grid: Grid) => void;
  /** Parse a clipboard DataTransfer to a Grid (= the app's parseClipboard). */
  parseClipboardData: (data: DataTransfer) => Grid;
  /** Parse a dropped/picked spreadsheet FILE to a Grid (async; = the app's parseFile). */
  parseFile: (file: File) => Promise<Grid>;
  /** Load the built-in example section (parent handles it) then the modal closes. */
  onUseExample: () => void;
  /** True when the section cap is reached — show the limit message, block commit. */
  atLimit?: boolean;
  limitMessage?: string;
}): ReactPortal | null {
  // The read grid IS the phase selector: null = intake, non-null = preview. No
  // separate "phase" enum, so the two can never disagree.
  const [readGrid, setReadGrid] = useState<Grid | null>(null);
  const [error, setError] = useState<string | null>(null);
  // Async file parse in flight — drives the well's "Reading…" state and (with the
  // grid absent) suppresses the paste listener so a mid-read paste can't race it.
  const [reading, setReading] = useState(false);
  // Drag hover highlight on the drop well (a real dragenter/over/leave latch).
  const [dragActive, setDragActive] = useState(false);

  const titleId = useId();
  const wellRef = useRef<HTMLDivElement>(null);
  const closeRef = useRef<HTMLButtonElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Validate a freshly parsed grid, then either advance to preview or surface the
  // inline error and stay in intake. Stable so effects/handlers below don't
  // needlessly re-subscribe. Total-cell count (not row/col) is the guard so both a
  // single-cell paste and an empty grid fall out the same way.
  const acceptGrid = useCallback((grid: Grid) => {
    const cells = grid.reduce((n, row) => n + row.length, 0);
    if (cells < MIN_CELLS) {
      setError(READ_ERROR);
      setReadGrid(null);
      return;
    }
    setError(null);
    setReadGrid(grid);
  }, []);

  // Read a dropped/picked FILE. The brief `reading` flag lets the well show a
  // spinner; any thrown parse error collapses to the same READ_ERROR as a bad
  // paste so failure is uniform.
  const acceptFile = useCallback(
    async (file: File) => {
      setReading(true);
      setError(null);
      try {
        const grid = await parseFile(file);
        acceptGrid(grid);
      } catch {
        setError(READ_ERROR);
        setReadGrid(null);
      } finally {
        setReading(false);
      }
    },
    [parseFile, acceptGrid],
  );

  // Fresh start on every open→true transition: clear any leftover read/error state
  // so reopening never shows a stale preview, then focus the well (or the close
  // button when the well is hidden at the cap) so Escape/keyboard works at once.
  // rAF defers the focus until after the portal has painted its children.
  useEffect(() => {
    if (!open) return;
    // Deliberate reset-on-open (sync to an external trigger, which this rule is not
    // meant to forbid): wipe any leftover read/error state so a reopen is fresh.
    /* eslint-disable react-hooks/set-state-in-effect */
    setReadGrid(null);
    setError(null);
    setReading(false);
    setDragActive(false);
    /* eslint-enable react-hooks/set-state-in-effect */
    const raf = requestAnimationFrame(() => {
      (wellRef.current ?? closeRef.current)?.focus();
    });
    return () => cancelAnimationFrame(raf);
  }, [open]);

  // Escape closes (matches Popover). Kept separate from the intake-only paste
  // listener so it stays active in BOTH phases.
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  // "Listening for paste": while intake is active (open, nothing read yet, not
  // reading a file, not at the cap) a WINDOW paste handler feeds the clipboard
  // straight into the parser. The parent suppresses its own global paste while
  // this modal is open, so there is no double-handling. preventDefault stops the
  // browser from also dropping the pasted text into any focused field.
  const canPaste = open && readGrid === null && !reading && !atLimit;
  useEffect(() => {
    if (!canPaste) return;
    const onPaste = (e: ClipboardEvent) => {
      if (!e.clipboardData) return;
      e.preventDefault();
      acceptGrid(parseClipboardData(e.clipboardData));
    };
    window.addEventListener("paste", onPaste);
    return () => window.removeEventListener("paste", onPaste);
  }, [canPaste, parseClipboardData, acceptGrid]);

  // ---- Handlers -----------------------------------------------------------

  const onDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault(); // required for the drop to fire
    if (!dragActive) setDragActive(true);
  };
  const onDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragActive(false);
  };
  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragActive(false);
    const file = e.dataTransfer.files[0];
    if (file) void acceptFile(file);
  };
  const onFilePicked = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    // Reset the input so re-picking the SAME file still fires a change event.
    e.target.value = "";
    if (file) void acceptFile(file);
  };

  const clearRead = () => {
    setReadGrid(null);
    setError(null);
    // Return focus to the well so paste/keyboard is immediately live again.
    requestAnimationFrame(() => wellRef.current?.focus());
  };

  const commit = (grid: Grid) => {
    onCommit(grid);
    onClose();
  };

  // Nothing in the DOM when closed, and never during SSR (no document to portal to).
  if (!open || typeof document === "undefined") return null;

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      {/* Dimmed backdrop — a real button so a click (or the assistive "close"
          affordance) dismisses; tabIndex -1 keeps it out of the tab order. */}
      <button
        type="button"
        aria-label="Close"
        tabIndex={-1}
        onClick={onClose}
        className="absolute inset-0 cursor-default bg-black/40"
      />
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        /* Wide enough that a real pasted table stays legible in the preview —
           the whole point of confirming before commit. The intake phase simply
           centers its drop well inside the extra width. */
        className="relative flex max-h-[90vh] w-full max-w-[56rem] flex-col overflow-hidden rounded-xl border border-border-strong bg-surface text-foreground shadow-[var(--shadow-4)]"
      >
        {/* Header ---------------------------------------------------------- */}
        <div className="flex items-center justify-between border-b border-border px-5 py-3.5">
          <h2 id={titleId} className="text-base font-semibold text-foreground">
            Add a table
          </h2>
          <button
            ref={closeRef}
            type="button"
            aria-label="Close"
            onClick={onClose}
            className="inline-flex h-7 w-7 items-center justify-center rounded text-muted transition-colors hover:bg-surface-alt hover:text-foreground"
          >
            <span aria-hidden className="text-base leading-none">
              &times;
            </span>
          </button>
        </div>

        {/* Body ------------------------------------------------------------ */}
        <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto px-5 py-5">
          {readGrid === null ? (
            // ---- PHASE 1: INTAKE -------------------------------------------
            <>
              {atLimit ? (
                // At the cap: no live intake. Show the message in a disabled well.
                <div className="flex min-h-[9rem] items-center justify-center rounded-lg border-2 border-dashed border-border bg-surface-alt px-6 text-center text-sm text-muted">
                  {limitMessage ??
                    "Section limit reached. Remove a section to add another."}
                </div>
              ) : (
                <>
                  <div
                    ref={wellRef}
                    tabIndex={0}
                    role="button"
                    aria-label="Drop a spreadsheet file here, or paste a copied cell range with Control or Command plus V."
                    aria-keyshortcuts="Control+V"
                    onDragOver={onDragOver}
                    onDragEnter={onDragOver}
                    onDragLeave={onDragLeave}
                    onDrop={onDrop}
                    className={`flex min-h-[9rem] flex-col items-center justify-center gap-2 rounded-lg border-2 border-dashed px-6 py-6 text-center outline-none transition-colors ${
                      dragActive
                        ? "border-accent bg-accent-subtle"
                        : "border-border-strong bg-surface-alt focus:border-accent focus:bg-accent-subtle"
                    }`}
                  >
                    {reading ? (
                      <span className="flex items-center gap-2 text-sm text-muted">
                        <span
                          aria-hidden
                          className="h-4 w-4 animate-spin rounded-full border-2 border-border-strong border-t-accent"
                        />
                        Reading…
                      </span>
                    ) : (
                      <>
                        <svg
                          aria-hidden
                          width="28"
                          height="28"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          className="text-accent-text"
                        >
                          <path d="M12 15V3" />
                          <path d="M7 8l5-5 5 5" />
                          <path d="M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4" />
                        </svg>
                        <p className="text-sm font-medium text-foreground">
                          Paste, drop a file, or browse
                        </p>
                        <p className="text-xs text-muted">
                          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
                            Ctrl
                          </kbd>
                          +
                          <kbd className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]">
                            V
                          </kbd>{" "}
                          from Excel · or drop an .xlsx / .csv here
                        </p>
                        <button
                          type="button"
                          onClick={() => fileInputRef.current?.click()}
                          className={`mt-1 ${BTN_SUBTLE} border border-border-strong`}
                        >
                          Browse files
                        </button>
                      </>
                    )}
                  </div>

                  {/* Live intake indicator — the pulsing dot signals the window
                      paste listener is armed. */}
                  <p className="flex items-center gap-1.5 text-xs text-muted">
                    <span aria-hidden className="animate-pulse text-accent">
                      &#9679;
                    </span>
                    Listening for paste…
                  </p>

                  {error && (
                    <p
                      role="alert"
                      className="rounded border border-[color:var(--danger-border)] bg-[color:var(--danger-bg)] px-3 py-2 text-sm text-danger"
                    >
                      {error}
                    </p>
                  )}

                  <p className="text-xs text-muted">
                    No spreadsheet?{" "}
                    <button
                      type="button"
                      onClick={onUseExample}
                      className="font-semibold text-accent-text hover:underline"
                    >
                      Try an example
                    </button>
                  </p>
                </>
              )}
            </>
          ) : (
            // ---- PHASE 2: PREVIEW ------------------------------------------
            <Preview grid={readGrid} />
          )}
        </div>

        {/* Footer (preview actions only) ----------------------------------- */}
        {readGrid !== null && (
          <div className="flex items-center justify-end gap-2 border-t border-border px-5 py-3.5">
            <button type="button" onClick={clearRead} className={BTN_SUBTLE}>
              Clear
            </button>
            <button
              type="button"
              onClick={() => commit(readGrid)}
              disabled={atLimit}
              title={atLimit ? limitMessage : undefined}
              className={BTN_PRIMARY}
            >
              Add section
            </button>
          </div>
        )}

        {/* Hidden picker driven by the "Browse files" button. Accepts the
            spreadsheet/flat formats the parser understands. */}
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv,.tsv,.txt"
          onChange={onFilePicked}
          className="hidden"
        />
      </div>
    </div>,
    document.body,
  );
}

/**
 * The phase-2 read-back: a "read N rows × M columns" summary plus a compact,
 * scrollable mini-table of the first few rows. Deliberately lightweight (no
 * GridTable import) — it only needs to prove the parse landed on real columns, so
 * it mirrors GridTable's plain look (distinct bold header on `bg-surface-alt`,
 * hairline borders, `text-xs`) without the sticky gutter/header machinery.
 */
function Preview({ grid }: { grid: Grid }) {
  // N excludes the header row (row 0), clamped so a header-only grid reads "0".
  const dataRows = Math.max(0, grid.length - 1);
  const cols = grid.reduce((max, row) => Math.max(max, row.length), 0);
  // Show EVERY row: this is the confirm-what-I-pasted step, so a truncated
  // sample defeats the point — you can't spot a missing tail or a shifted column
  // in the first five rows. The container scrolls (both axes) and the header row
  // sticks, so a 500-row paste stays reviewable without swamping the modal.
  const shown = grid;

  return (
    <>
      <p className="text-sm font-medium text-success">
        <span aria-hidden>&#10003;</span> Read {dataRows} row
        {dataRows === 1 ? "" : "s"} &times; {cols} column{cols === 1 ? "" : "s"}
      </p>

      <div className="max-h-[55vh] overflow-auto rounded-lg border border-border">
        <table className="w-full border-collapse text-xs">
          <tbody>
            {shown.map((row, ri) => (
              <tr
                key={ri}
                className={
                  ri === 0
                    ? "sticky top-0 z-10 bg-surface-alt"
                    : ri % 2
                      ? "bg-surface-alt"
                      : "bg-surface"
                }
              >
                {Array.from({ length: cols }, (_, c) => {
                  const text = cellText(row[c]);
                  return ri === 0 ? (
                    <th
                      key={c}
                      scope="col"
                      className="whitespace-nowrap border-b border-border-strong px-2.5 py-1.5 text-left font-semibold text-foreground"
                    >
                      {text || (
                        <span className="font-normal text-muted">
                          Column {c + 1}
                        </span>
                      )}
                    </th>
                  ) : (
                    <td
                      key={c}
                      className="max-w-[16rem] truncate border-b border-border px-2.5 py-1.5 align-top text-foreground"
                    >
                      {text}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

    </>
  );
}

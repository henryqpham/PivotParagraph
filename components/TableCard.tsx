"use client";

import {
  Fragment,
  memo,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from "react";
import { headingLevel } from "@/lib/clipboard";
import {
  defaultMarkerSpec,
  isCounterType,
  type MarkerType,
  type MarkerDelim,
} from "@/lib/renderers";
import { DEFAULT_FIELD_LABEL, type FieldLabel } from "@/lib/types";
import { Popover } from "./Popover";
import {
  addField,
  canIndent,
  canOutdent,
  indentField,
  moveField,
  outdentField,
  removeField,
  unusedColumns,
  bodyGrid,
  resolveMarkerSpecs,
  DEFAULT_LEVEL,
  MAX_LEVELS,
  type LevelInput,
  type TableState,
} from "./tableModel";

/** The shared title look, held as UI inputs (size a string so it can be cleared). */
export type TitleInput = {
  font: string;
  sizeInput: string;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
};

/** Allow-listed body/title fonts. Exported so PasteInput's Document popover can
 *  reuse the same list without redefining it. */
export const FONTS = [
  "Calibri Light",
  "Calibri",
  "Arial",
  "Times New Roman",
  "Georgia",
  "Cambria",
];

/** Word heading styles the title can map to ("" = None = the app's direct look). */
const HEADING_OPTIONS = [
  { value: "", label: "None (direct look)" },
  { value: "Heading 1", label: "Heading 1" },
  { value: "Heading 2", label: "Heading 2" },
  { value: "Heading 3", label: "Heading 3" },
  { value: "Heading 4", label: "Heading 4" },
];
/** Dropdown sentinel for "map the title to a custom destination style name". */
const CUSTOM_HEADING = "__custom__";

/** Marker COUNTER types, labeled by the glyph they produce (narrow for the
 *  matrix). Split from the delimiter so any type × any delimiter is reachable. */
const MARKER_TYPE_OPTIONS: { value: MarkerType; label: string }[] = [
  { value: "decimal", label: "1" },
  { value: "lowerAlpha", label: "a" },
  { value: "upperAlpha", label: "A" },
  { value: "lowerRoman", label: "i" },
  { value: "upperRoman", label: "I" },
  { value: "bullet", label: "•" },
  { value: "dash", label: "–" },
  { value: "none", label: "None" },
];
/** Marker DELIMITER glued after a counter (disabled for symbol/none types).
 *  "None" is spelled out rather than drawn as an em-dash: a "—" glyph reads like a
 *  dash you'd GET after the number, when it actually means no delimiter at all. */
const MARKER_DELIM_OPTIONS: { value: MarkerDelim; label: string }[] = [
  { value: "dot", label: "." },
  { value: "paren", label: ")" },
  { value: "none", label: "None" },
];

// ---- Fluent control recipes (shared) --------------------------------------
const CARD =
  "rounded-lg border border-border bg-surface p-4 shadow-[var(--shadow-2)]";
const GROUP_HEADER =
  "mb-3 border-b border-border pb-2 text-xs font-semibold uppercase tracking-wide text-muted";
const SUB =
  "mt-4 mb-1.5 text-xs font-semibold uppercase tracking-wide text-muted";
const FIELD =
  "h-8 rounded border border-border-strong bg-surface px-2.5 text-sm text-foreground outline-none transition-colors hover:border-b-[color:var(--text-secondary)] focus:border-accent";
const BADGE =
  "rounded-sm bg-[color:color-mix(in_srgb,var(--muted)_14%,transparent)] px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-wide text-muted";
// Toggle button (B/I/U/sort/Aa): brand-tinted "pressed-in" when active. Module-
// level so both TableCardInner and the shared LookControl use the one recipe.
const tgl = (active: boolean) =>
  `flex h-7 min-w-7 items-center justify-center rounded border px-1.5 text-xs transition-colors disabled:opacity-40 ${
    active
      ? "border-accent bg-accent-subtle text-accent-text"
      : "border-transparent text-text-secondary hover:bg-surface-alt hover:text-foreground"
  }`;

/**
 * The shared per-target text "Look" control: a color swatch + an "Aa▾" trigger
 * that opens a popover with Font / Size / B / I / U — used IDENTICALLY for the
 * Section Title and every per-level row, so "the same thing looks the same"
 * (Word's one-Modify-Style-dialog model). `value`/`onChange` fit both the
 * title's `TitleInput` and a level's `LevelInput`. `bodyFontOption` (level rows
 * only) prepends a "Body font (…)" inherit choice; the title always picks a real
 * font. Open state is controlled by the parent so only one popover shows at once.
 */
type LookValue = {
  font: string;
  sizeInput: string;
  color: string;
  bold: boolean;
  italic?: boolean;
  underline?: boolean;
};
function LookControl({
  value,
  onChange,
  open,
  onOpenChange,
  label,
  bodyFontOption,
}: {
  value: LookValue;
  onChange: (patch: Partial<LookValue>) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  label: string;
  bodyFontOption?: string;
}) {
  return (
    <div className="relative flex items-center gap-1">
      <input
        type="color"
        value={value.color}
        onChange={(e) => onChange({ color: e.target.value })}
        aria-label={`${label} text color (all tables)`}
        title="Text color · all tables"
        className="h-6 w-6 cursor-pointer rounded border border-border-strong"
      />
      <button
        type="button"
        onClick={() => onOpenChange(!open)}
        aria-expanded={open}
        aria-label={`${label} font, size, and emphasis (all tables)`}
        title="Font · size · bold · italic · underline · all tables"
        className="rounded border border-border-strong bg-surface px-1 text-[11px] font-semibold text-text-secondary transition-colors hover:border-accent hover:text-accent-text"
      >
        Aa▾
      </button>
      <Popover open={open} onClose={() => onOpenChange(false)}>
        <div className="flex w-52 flex-col gap-2 text-sm text-text-secondary">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-muted">
            {label} look
            <span className="ml-1 font-normal normal-case text-accent-text">
              · all tables
            </span>
          </div>
          <label className="flex items-center justify-between gap-3">
            Font
            <select
              value={value.font}
              onChange={(e) => onChange({ font: e.target.value })}
              aria-label={`${label} font${bodyFontOption ? " (empty = the document Body font)" : ""}`}
              className={FIELD}
            >
              {bodyFontOption !== undefined && (
                <option value="">Body font ({bodyFontOption})</option>
              )}
              {FONTS.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
          </label>
          <label className="flex items-center justify-between gap-3">
            Size (pt)
            <input
              type="text"
              inputMode="numeric"
              value={value.sizeInput}
              onChange={(e) =>
                onChange({ sizeInput: e.target.value.replace(/[^0-9]/g, "") })
              }
              aria-label={`${label} size in points`}
              className={`${FIELD} w-16`}
            />
          </label>
          {/* B / I / U as toggle buttons — the ONE emphasis convention across
              the whole tool (title, levels, and the field-label toggles all use
              this button style). */}
          <div className="flex items-center gap-1">
            <button
              type="button"
              aria-pressed={value.bold}
              title={`Bold ${label}`}
              onClick={() => onChange({ bold: !value.bold })}
              className={`${tgl(!!value.bold)} font-bold`}
            >
              B
            </button>
            <button
              type="button"
              aria-pressed={!!value.italic}
              title={`Italicize ${label}`}
              onClick={() => onChange({ italic: !value.italic })}
              className={`${tgl(!!value.italic)} italic`}
            >
              I
            </button>
            <button
              type="button"
              aria-pressed={!!value.underline}
              title={`Underline ${label}`}
              onClick={() => onChange({ underline: !value.underline })}
              className={`${tgl(!!value.underline)} underline`}
            >
              U
            </button>
          </div>
        </div>
      </Popover>
    </div>
  );
}

/** Per-row Windows-Explorer connector state, derived from the flat depth list. */
type RowGuides = {
  /** For each ancestor depth a in 0..b-1: does that level have another sibling
   *  further down (→ a continuing │) or not (→ a blank spacer)? Length === depth. */
  ancestors: boolean[];
  /** Is this row the last field of its level in its parent group (└ vs ├)? */
  last: boolean;
};

/**
 * Windows-Explorer connector guides for the flat, depth-indexed Structure list. For
 * each row and each ancestor depth we look ahead for the next sibling at that depth
 * (stopping if we leave the ancestor's scope), and likewise decide whether the row
 * is the last at its own depth. Correct for stacked fields (several fields at one
 * depth). O(n·depth) — fine for a field list.
 */
function rowGuides(placed: { b: number }[]): RowGuides[] {
  const depths = placed.map((p) => p.b);
  const n = depths.length;
  return depths.map((d, i) => {
    const ancestors: boolean[] = [];
    for (let a = 0; a < d; a++) {
      let continues = false;
      for (let j = i + 1; j < n; j++) {
        if (depths[j] < a) break; // left the ancestor's parent group
        if (depths[j] === a) {
          continues = true; // a later sibling exists at depth a
          break;
        }
      }
      ancestors.push(continues);
    }
    let last = true;
    for (let j = i + 1; j < n; j++) {
      if (depths[j] < d) break;
      if (depths[j] === d) {
        last = false; // a later sibling at this row's own depth
        break;
      }
    }
    return { ancestors, last };
  });
}

/** One 20px guide column: a continuing │, an elbow (├/└), or a blank spacer.
 *  Absolute 1px borders on --border-strong render crisply in light + dark. */
function GuideCell({ kind }: { kind: "line" | "tee" | "corner" | "blank" }) {
  const v = "absolute left-1/2 border-l border-border-strong";
  return (
    <span className="relative w-5 shrink-0">
      {(kind === "line" || kind === "tee") && (
        <span className={`${v} top-0 bottom-0`} />
      )}
      {kind === "corner" && <span className={`${v} top-0 h-1/2`} />}
      {(kind === "tee" || kind === "corner") && (
        <span className="absolute left-1/2 right-0 top-1/2 border-t border-border-strong" />
      )}
    </span>
  );
}

/**
 * The GLOBAL per-depth body look (owned by PasteInput), edited inline in the
 * matrix's "Look" column so a level's look sits with the level that defines it.
 * Document-wide settings (body font / indent / reset) now live in PasteInput's
 * top-band "Document" popover, so this contract is just the two the Look column
 * needs (both shared across every section).
 */
type AppearanceControls = {
  levelStyles: LevelInput[];
  onLevelChange: (i: number, patch: Partial<LevelInput>) => void;
  /** The document-wide Body font — what an unpinned level's font ("" = inherit)
   *  resolves to, shown in the Look popover's "Body font (…)" option. */
  bodyFont: string;
};

type Props = {
  table: TableState;
  onChange: (patch: Partial<TableState>) => void;
  /** Global Word-heading style the TITLE maps to ("" = None). Shared. */
  headingStyleName: string;
  onHeadingStyleChange: (value: string) => void;
  /** Shared title look (font/size/color + bold/italic/underline). */
  title: TitleInput;
  onTitleChange: (patch: Partial<TitleInput>) => void;
  /** Global per-depth body look (all tables), edited in the matrix "Look" column. */
  appearance: AppearanceControls;
};

function TableCardInner({
  table,
  onChange,
  headingStyleName,
  onHeadingStyleChange,
  title,
  onTitleChange,
  appearance,
}: Props) {
  // Start-number field held as a string so it can be cleared/retyped (the card is
  // keyed by table id, so this resets per table).
  const [startInput, setStartInput] = useState(String(table.numbering.start));
  // Which "Look" popover is open, keyed "title" or `lvl-<i>`, or null. One key
  // space so only ONE Look popover (title OR a level) shows at a time.
  const [openPop, setOpenPop] = useState<string | null>(null);
  // A plain-text filter over the Add-fields pool, only shown once it's long enough
  // to need one (a wide spreadsheet can have dozens of columns).
  const [fieldFilter, setFieldFilter] = useState("");
  // Whether the title's Word-heading control is in "Custom style…" mode (a text
  // input for any destination style name). Initialized from the CURRENT name —
  // a name outside the dropdown's options can only be a custom one. The card
  // remounts per section (keyed by table id), so this re-derives on switch.
  const [customHeading, setCustomHeading] = useState(
    () =>
      headingStyleName !== "" &&
      !HEADING_OPTIONS.some((o) => o.value === headingStyleName),
  );

  const { pivotLevels } = table;
  // A body bucket's GLOBAL level-style slot = its indent depth, +1 when a Section
  // title sits above it (mirrors the renderer's walk(nodes, 2, 1)); clamp to the
  // 9-entry chart. Keeps each matrix "Look" cell driving the same data-level the
  // Word export uses, so preview and paste stay in sync.
  const titleOffset = table.sectionTitle.trim() ? 1 : 0;
  const levelIdxForBucket = (b: number) => Math.min(8, b + titleOffset);
  // The Word heading each level maps to, mirroring the renderer: AUTO = the
  // CONTIGUOUS rank — the title's heading (only when a title is actually emitted,
  // like the renderer's `bodyHeadingBase = title ? titleLevel : 0`) + how many
  // heading-checked levels sit at or above this depth — so checking levels 1 and 4
  // reads H2 → H3, never a skipped H2 → H5 outline. An explicit `headingRanks`
  // entry (1–9; 0 = auto) overrides, for templates whose numbering keys off a
  // specific rank. `skips` flags an override that jumps past a rank (Word's
  // accessibility checker dings gapped outlines) so the chip can warn.
  const headingBase = table.sectionTitle.trim()
    ? headingLevel(headingStyleName)
    : 0;
  const headingInfo = useMemo(() => {
    const out: { k: number; auto: number; explicit: number; skips: boolean }[] =
      [];
    let prevK = headingBase;
    for (let i = 0; i < pivotLevels.length; i++) {
      const isHeading = table.headingLevels[i] === true;
      const raw = table.headingRanks?.[i];
      const explicit =
        typeof raw === "number" && raw >= 1 && raw <= 9 ? Math.floor(raw) : 0;
      // Auto = one rank under the previous heading (title, then each checked
      // level in order, PINS INCLUDED) — so an auto level can never skip a rank
      // and `skips` (the amber warning) can only come from an explicit pin.
      const auto = Math.min(Math.max(prevK + 1, 1), 9);
      const k = explicit || auto;
      const skips = isHeading && k - prevK > 1;
      if (isHeading) prevK = k;
      out.push({ k, auto, explicit, skips });
    }
    return out;
  }, [pivotLevels, table.headingLevels, table.headingRanks, headingBase]);

  // Per-level marker specs (split type + delimiter), migrated from any legacy
  // fused `markers`. Patch one level's type or delimiter and write the whole
  // (normalized) array back so a legacy source is upgraded on first edit.
  const markerSpecs = resolveMarkerSpecs(table);
  const setMarkerSpec = (i: number, patch: Partial<{ type: MarkerType; delim: MarkerDelim }>) => {
    const next = [...markerSpecs];
    next[i] = { ...(next[i] ?? defaultMarkerSpec(i + 1)), ...patch };
    // Drop any legacy `markers` we just superseded so the table doesn't persist
    // both shapes (markerSpecs is authoritative from here on).
    onChange({ markerSpecs: next, markers: undefined });
  };

  // Field names come from the EFFECTIVE header row (bodyGrid honors a header offset
  // set in the Table view), so skipping banner rows renames the fields everywhere.
  const headers = useMemo(() => {
    const bg = bodyGrid(table);
    return bg[0] ? bg[0].map((c) => (c == null ? "" : String(c))) : [];
  }, [table]);

  const unused = useMemo(
    () => unusedColumns(headers.length, pivotLevels),
    [headers.length, pivotLevels],
  );
  const showFieldFilter = unused.length > 8;
  const fieldFilterLower = fieldFilter.trim().toLowerCase();
  const visibleUnused = fieldFilterLower
    ? unused.filter((col) =>
        (headers[col] || `Column ${col + 1}`)
          .toLowerCase()
          .includes(fieldFilterLower),
      )
    : unused;

  const placed = useMemo(() => {
    // `k` = the field's index WITHIN its bucket. k === 0 is the level OWNER — the
    // one row that carries the per-level controls (marker / heading / look); a
    // stacked sibling (k > 0) shows only field controls + a "shares level" note.
    const out: { col: number; b: number; fi: number; k: number }[] = [];
    let fi = 0;
    pivotLevels.forEach((bucket, b) =>
      bucket.forEach((col, k) => {
        out.push({ col, b, fi, k });
        fi++;
      }),
    );
    return out;
  }, [pivotLevels]);

  const guides = useMemo(() => rowGuides(placed), [placed]);

  // First data row — one representative value per column, for the live
  // micro-preview that renders each field styled in its level's real look.
  const sampleRow = useMemo(() => bodyGrid(table)[1] ?? [], [table]);

  // The resolved on-page look for a bucket's rows (mirrors the renderer/preview):
  // the shared per-depth level style, its font falling back to the Body font.
  const levelLook = (b: number): CSSProperties => {
    const lv = appearance.levelStyles[levelIdxForBucket(b)] ?? DEFAULT_LEVEL;
    return {
      color: lv.color || "#000000",
      fontFamily: lv.font || appearance.bodyFont,
      fontWeight: lv.bold ? 700 : 400,
      fontStyle: lv.italic ? "italic" : "normal",
      textDecoration: lv.underline ? "underline" : "none",
    };
  };

  // ---- "What just moved?" highlight ----------------------------------------
  // Reordering/indenting shifts a row out from under the cursor, which makes it
  // easy to lose track of the field you're arranging. The row a change just
  // affected stays tinted for a beat. Re-arming the timer on every change keeps
  // it lit CONTINUOUSLY through a burst of clicks (rather than restarting a
  // flash each time), so holding ▲ reads as one moving highlight.
  const [flashCol, setFlashCol] = useState<number | null>(null);
  const flashTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flashRow = (col: number) => {
    setFlashCol(col);
    if (flashTimer.current) clearTimeout(flashTimer.current);
    flashTimer.current = setTimeout(() => setFlashCol(null), 1100);
  };
  useEffect(
    () => () => {
      if (flashTimer.current) clearTimeout(flashTimer.current);
    },
    [],
  );

  /** Apply a `pivotLevels` change AND highlight the field it acted on. */
  const applyMove = (col: number, next: number[][]) => {
    onChange({ pivotLevels: next });
    flashRow(col);
  };

  function patchLabel(col: number, patch: Partial<FieldLabel>) {
    const cur = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
    onChange({
      fieldLabels: { ...table.fieldLabels, [col]: { ...cur, ...patch } },
    });
  }

  function cycleSort(col: number) {
    const cur = table.sortDirs[col];
    const next: Record<number, "asc" | "desc"> = { ...table.sortDirs };
    if (cur === undefined) next[col] = "asc";
    else if (cur === "asc") next[col] = "desc";
    else delete next[col];
    onChange({ sortDirs: next });
  }

  // Icon button (◄►▲▼✕) — Fluent subtle, now visible (not the old /40).
  const icon =
    "flex h-6 w-6 items-center justify-center rounded text-muted transition-colors hover:bg-surface-alt hover:text-foreground disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-muted";
  // `tgl` (the B/I/U/sort toggle recipe) is now module-level — shared with LookControl.

  return (
    <div className="flex flex-col gap-4">
      {/* ---- SECTION HEADER group (the level-0 title, its own home) ---------- */}
      <section className={CARD}>
        {/* "Section Title", not "Section Header" — "header" is overloaded here
            (table header row, Word page headers, the Heading dropdown beside it);
            this group is exactly the section's TITLE: its text, Word-heading
            mapping, and look. */}
        <h2 className={GROUP_HEADER}>Section Title</h2>
        <div className="flex flex-col gap-2.5 text-sm">
          <div className="flex flex-wrap items-center gap-2">
            <span className="w-20 shrink-0 text-xs text-text-secondary">
              Title
            </span>
            <input
              type="text"
              value={table.sectionTitle}
              onChange={(e) => onChange({ sectionTitle: e.target.value })}
              placeholder="e.g. Fruit Database"
              className={`${FIELD} w-56`}
            />
          </div>
          {/* Word heading + Look share ONE row: both are compact now (a select
              and a swatch+Aa▾), so pairing them fills the width instead of
              leaving two sparse rows. Each keeps its own "All tables" badge
              stacked under its label. The custom-style input/hint wraps below. */}
          <div className="flex flex-wrap items-start gap-x-8 gap-y-2.5">
            <div className="flex flex-wrap items-center gap-2">
              <span className="flex flex-col text-xs text-text-secondary">
                Word heading
                <span className={`${BADGE} mt-0.5 self-start`}>All tables</span>
              </span>
              {/* Hybrid control: the dropdown is the safe, verified fast path
                  (built-in Heading 1–4 map as real <hN> in every paste mode);
                  "Custom style…" reveals a text input for mapping the title to
                  ANY named style in the destination template (e.g. TBL_TITLE) —
                  the legacy mso-style-name route, which maps on a Use-
                  Destination-Styles paste and needs the name to exist there. */}
              <select
                value={customHeading ? CUSTOM_HEADING : headingStyleName}
                onChange={(e) => {
                  if (e.target.value === CUSTOM_HEADING) {
                    // Start blank (= unmapped) until a real name is typed, so a
                    // half-configured custom style never silently emits.
                    setCustomHeading(true);
                    onHeadingStyleChange("");
                  } else {
                    setCustomHeading(false);
                    onHeadingStyleChange(e.target.value);
                  }
                }}
                aria-label="Word heading style for the title"
                className={FIELD}
              >
                {HEADING_OPTIONS.map((o) => (
                  <option key={o.value} value={o.value}>
                    {o.label}
                  </option>
                ))}
                <option value={CUSTOM_HEADING}>Custom style…</option>
              </select>
              {customHeading && (
                <>
                  <input
                    type="text"
                    value={headingStyleName}
                    onChange={(e) => onHeadingStyleChange(e.target.value)}
                    placeholder="Style name, e.g. TBL_TITLE"
                    maxLength={60}
                    aria-label="Custom destination style name for the title"
                    className={`${FIELD} w-52`}
                  />
                  <span className="basis-full text-[11px] leading-snug text-muted">
                    Must match a style that exists in the destination document —
                    maps on a <strong>Use Destination Styles</strong>{" "}paste.
                  </span>
                </>
              )}
            </div>
            {/* Look — the SAME shared control the per-level rows use (swatch +
                Aa▾ popover), so the title and the levels read as one system. */}
            <div className="flex items-center gap-2">
              <span className="flex flex-col text-xs text-text-secondary">
                Look
                <span className={`${BADGE} mt-0.5 self-start`}>All tables</span>
              </span>
              <LookControl
                value={title}
                onChange={onTitleChange}
                open={openPop === "title"}
                onOpenChange={(o) => setOpenPop(o ? "title" : null)}
                label={table.sectionTitle.trim() || "Title"}
              />
            </div>
          </div>
        </div>
      </section>

      {/* ---- ROWS group (STRUCTURE only: which fields, nested how, and each
             field's own label/sort). Everything that formats a whole LEVEL lives
             in the separate "Level formatting" card below. ---------------------- */}
      <section className={CARD}>
        <h2 className={GROUP_HEADER}>Rows</h2>

        {headers.length === 0 ? (
          <p className="text-sm text-muted">Paste a table to build the outline.</p>
        ) : (
          <div className="flex flex-col gap-1 text-sm text-text-secondary">
            {/* Add fields */}
            <div className={`${SUB} !mt-0 flex items-center justify-between gap-2`}>
              <span>Add fields</span>
              {showFieldFilter && (
                <div className="relative">
                  <input
                    type="text"
                    value={fieldFilter}
                    onChange={(e) => setFieldFilter(e.target.value)}
                    placeholder="Filter fields…"
                    aria-label="Filter the Add-fields pool"
                    className="h-6 w-40 rounded border border-border-strong bg-surface px-2 pr-5 text-[11px] normal-case tracking-normal text-foreground outline-none transition-colors focus:border-accent"
                  />
                  {fieldFilter && (
                    <button
                      type="button"
                      onClick={() => setFieldFilter("")}
                      aria-label="Clear filter"
                      className="absolute right-1.5 top-1/2 -translate-y-1/2 leading-none text-muted transition-colors hover:text-foreground"
                    >
                      &times;
                    </button>
                  )}
                </div>
              )}
            </div>
            <div className="flex flex-wrap items-center gap-1.5">
              {unused.length === 0 ? (
                <span className="text-xs text-muted">All fields added.</span>
              ) : visibleUnused.length === 0 ? (
                <span className="text-xs text-muted">
                  No fields match &ldquo;{fieldFilter.trim()}&rdquo;.
                </span>
              ) : (
                visibleUnused.map((col) => (
                  <button
                    key={col}
                    type="button"
                    onClick={() =>
                      applyMove(col, addField(pivotLevels, col))
                    }
                    title="Add as a new deepest level, then use ◄ ► ▲ ▼ to place it"
                    className="h-7 rounded border border-dashed border-border-strong px-2.5 text-xs text-text-secondary transition-colors hover:border-solid hover:border-accent hover:bg-accent-subtle hover:text-accent-text"
                  >
                    + {headers[col] || `Column ${col + 1}`}
                  </button>
                ))
              )}
            </div>

            {/* Structure */}
            {placed.length > 0 && (
              <>
                <div className={`${SUB} flex items-center justify-between gap-2`}>
                  <span>Rows</span>
                  <span
                    title="Word supports up to 9 indent levels"
                    className={`rounded-sm px-1.5 py-0.5 text-[10px] font-semibold normal-case tracking-normal ${
                      pivotLevels.length > MAX_LEVELS
                        ? "bg-[color:var(--danger-bg)] text-danger"
                        : pivotLevels.length === MAX_LEVELS
                          ? "bg-[color:color-mix(in_srgb,var(--warning)_16%,transparent)] text-[color:var(--warning)]"
                          : "bg-[color:color-mix(in_srgb,var(--muted)_14%,transparent)] text-muted"
                    }`}
                  >
                    {pivotLevels.length}/{MAX_LEVELS} levels
                  </span>
                </div>
                {pivotLevels.length >= MAX_LEVELS && (
                  <p
                    role="status"
                    className={`mb-1 flex items-start gap-1.5 rounded border px-2.5 py-1.5 text-[11px] leading-snug ${
                      pivotLevels.length > MAX_LEVELS
                        ? "border-[color:var(--danger-border)] bg-[color:var(--danger-bg)] text-danger"
                        : "border-[color:color-mix(in_srgb,var(--warning)_35%,transparent)] bg-[color:color-mix(in_srgb,var(--warning)_10%,transparent)] text-[color:var(--warning)]"
                    }`}
                  >
                    <span aria-hidden>&#9888;</span>
                    {pivotLevels.length > MAX_LEVELS ? (
                      <span>
                        This section has <strong>{pivotLevels.length} levels</strong>,
                        but Word (and this preview) show only <strong>9</strong>
                        {" "}— levels 10+ collapse onto level 9. Outdent (&#9668;)
                        or remove fields until you&apos;re at 9 or fewer.
                      </span>
                    ) : (
                      <span>
                        <strong>Maximum depth reached.</strong>{" "}Word supports
                        9 indent levels — new fields now stack at level 9 instead
                        of nesting deeper.
                      </span>
                    )}
                  </p>
                )}
                {/* One row per FIELD. The level OWNER (first field of a bucket)
                    carries the per-level controls (marker / heading / look); a
                    stacked sibling shows only field controls + a "shares level"
                    note, so the level-vs-field distinction reads honestly.
                    Arranged entirely with the ◄ ► ▲ ▼ ✕ buttons (drag-and-drop
                    was tried and removed — the arrows read as more intuitive and
                    are keyboard-accessible for free). gap-0 so the tree guides
                    stay continuous row-to-row. */}
                <div className="flex flex-col">
                  {placed.map(({ col, b, fi, k }) => {
                    const name = headers[col] || `Column ${col + 1}`;
                    const lf = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
                    const g = guides[fi];
                    const owner = k === 0;
                    const sample = String(sampleRow[col] ?? "").trim();
                    return (
                      <Fragment key={col}>
                        <div
                          className={`group flex min-h-[36px] items-center gap-1.5 rounded px-1 transition-colors duration-300 motion-reduce:transition-none ${
                            flashCol === col
                              ? "bg-[color:color-mix(in_srgb,var(--accent)_20%,transparent)] ring-1 ring-inset ring-[color:color-mix(in_srgb,var(--accent)_45%,transparent)]"
                              : "hover:bg-surface-alt"
                          }`}
                        >
                          {/* Explorer tree-line guides: b ancestor cells + 1 elbow */}
                          <span aria-hidden className="flex self-stretch">
                            {g.ancestors.map((cont, a) => (
                              <GuideCell key={a} kind={cont ? "line" : "blank"} />
                            ))}
                            <GuideCell kind={g.last ? "corner" : "tee"} />
                          </span>
                          {/* Level number pill (dim/blank on a stacked sibling) */}
                          <span
                            className={`flex h-[18px] w-[18px] shrink-0 items-center justify-center rounded-[3px] text-[11px] font-semibold tabular-nums ${
                              owner
                                ? "bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-muted"
                                : "text-transparent"
                            }`}
                          >
                            {b + 1}
                          </span>
                          {/* Field name — a live micro-preview in the level's REAL
                              look — plus a faint sample value from the first row. */}
                          <span className="flex min-w-0 flex-1 items-baseline gap-1.5 truncate">
                            <span
                              style={levelLook(b)}
                              title={`Renders in level ${b + 1}'s look`}
                              className="truncate text-[13px] leading-tight"
                            >
                              {name}
                            </span>
                            {sample && (
                              <span className="truncate text-[11px] text-muted">
                                {sample}
                              </span>
                            )}
                          </span>
                          {/* Per-FIELD controls: show / bold / italic / underline
                              the label, then sort — on every row. */}
                          <span className="flex shrink-0 items-center gap-0.5">
                            <button
                              type="button"
                              aria-pressed={lf.show}
                              aria-label={`Show the "${name}:" label`}
                              title={`Show/hide the "${name}:" label`}
                              onClick={() => patchLabel(col, { show: !lf.show })}
                              className={tgl(lf.show)}
                            >
                              Aa
                            </button>
                            <button
                              type="button"
                              aria-pressed={lf.bold}
                              disabled={!lf.show}
                              aria-label={`Bold the "${name}:" label`}
                              title="Bold the label"
                              onClick={() => patchLabel(col, { bold: !lf.bold })}
                              className={`${tgl(lf.bold)} font-bold`}
                            >
                              B
                            </button>
                            <button
                              type="button"
                              aria-pressed={lf.italic}
                              disabled={!lf.show}
                              aria-label={`Italicize the "${name}:" label`}
                              title="Italicize the label"
                              onClick={() => patchLabel(col, { italic: !lf.italic })}
                              className={`${tgl(lf.italic)} italic`}
                            >
                              I
                            </button>
                            <button
                              type="button"
                              aria-pressed={lf.underline}
                              disabled={!lf.show}
                              aria-label={`Underline the "${name}:" label`}
                              title="Underline the label"
                              onClick={() =>
                                patchLabel(col, { underline: !lf.underline })
                              }
                              className={`${tgl(lf.underline)} underline`}
                            >
                              U
                            </button>
                            {(() => {
                              const dir = table.sortDirs[col];
                              const glyph =
                                dir === "asc" ? "↑" : dir === "desc" ? "↓" : "↕";
                              const dirLabel =
                                dir === "asc"
                                  ? "ascending"
                                  : dir === "desc"
                                    ? "descending"
                                    : "off";
                              return (
                                <button
                                  type="button"
                                  aria-label={`Sort by ${name} (${dirLabel})`}
                                  title={`Sort groups by ${name} (currently ${dirLabel})`}
                                  onClick={() => cycleSort(col)}
                                  className={tgl(dir !== undefined)}
                                >
                                  {glyph}
                                </button>
                              );
                            })()}
                          </span>
                          {/* A stacked sibling notes which level it shares. All
                              per-LEVEL controls (marker / heading / look) live in
                              the "Level formatting" card below, so this row stays
                              purely about the FIELD. */}
                          {!owner && (
                            <span
                              className="shrink-0 border-l border-border pl-2 text-[11px] italic text-muted"
                              title={`Stacked at level ${b + 1} — its marker, heading, and look are set on level ${b + 1} under Level formatting`}
                            >
                              shares level {b + 1}
                            </span>
                          )}
                          {/* Actions: outdent / indent / remove (reorder = ⋮ handle) */}
                          <span className="flex shrink-0 items-center">
                            <button
                              type="button"
                              aria-label={`Outdent ${name}`}
                              disabled={!canOutdent(pivotLevels, fi)}
                              onClick={() =>
                                applyMove(col, outdentField(pivotLevels, fi))
                              }
                              className={icon}
                            >
                              ◄
                            </button>
                            <button
                              type="button"
                              aria-label={`Indent ${name}`}
                              disabled={!canIndent(pivotLevels, fi)}
                              onClick={() =>
                                applyMove(col, indentField(pivotLevels, fi))
                              }
                              className={icon}
                            >
                              ►
                            </button>
                            <button
                              type="button"
                              aria-label={`Move ${name} up`}
                              title="Move up"
                              disabled={fi === 0}
                              onClick={() =>
                                applyMove(col, moveField(pivotLevels, fi, -1))
                              }
                              className={icon}
                            >
                              ▲
                            </button>
                            <button
                              type="button"
                              aria-label={`Move ${name} down`}
                              title="Move down"
                              disabled={fi === placed.length - 1}
                              onClick={() =>
                                applyMove(col, moveField(pivotLevels, fi, 1))
                              }
                              className={icon}
                            >
                              ▼
                            </button>
                            <button
                              type="button"
                              aria-label={`Remove ${name}`}
                              title="Remove this field"
                              onClick={() =>
                                onChange({
                                  pivotLevels: removeField(pivotLevels, fi),
                                })
                              }
                              className={`${icon} ml-0.5`}
                            >
                              ✕
                            </button>
                          </span>
                        </div>
                      </Fragment>
                    );
                  })}
                </div>
              </>
            )}

          </div>
        )}
      </section>

      {/* ---- LEVEL FORMATTING card (per-LEVEL: one row per indent level) -----
          Deliberately SPLIT from the Rows canvas: the Rows list is about the
          FIELDS (what nests where, and how each field's own `Name:` label reads),
          while everything that formats a whole LEVEL — its marker/number, its
          Word-heading mapping, and its look — lives here, mirroring the Section
          Title card. Keeping them apart is what makes the two scopes legible:
          the row's Aa/B/I/U touch ONLY the label before the colon, whereas a
          level's Look restyles the entire line (label AND value).
          Scope is mixed on purpose and marked per column: Marker + Heading are
          per-TABLE, the Look is shared by DEPTH across ALL sections. */}
      {pivotLevels.length > 0 && (
        <section className={CARD}>
          <h2 className={GROUP_HEADER}>Level formatting</h2>
          <div className="flex flex-col gap-1 text-sm text-text-secondary">
            {/* Markers mode — governs what the MARKER column below means (it is
                hidden entirely in "off", and becomes a show/hide number toggle in
                "multilevel"), so it sits directly above that column. */}
            <div className={`${SUB} !mt-0`}>Markers</div>
            <div className="flex flex-wrap items-center gap-2">
              <select
                value={table.numbering.mode}
                onChange={(e) =>
                  onChange({
                    numbering: {
                      ...table.numbering,
                      mode: e.target.value as "off" | "custom" | "multilevel",
                    },
                  })
                }
                aria-label="Marker / numbering mode"
                className={FIELD}
              >
                <option value="off">Off (none)</option>
                <option value="custom">Custom (per level)</option>
                <option value="multilevel">Multilevel numbers</option>
              </select>
              {table.numbering.mode === "multilevel" && (
                <>
                  <span className="text-xs text-text-secondary">Start</span>
                  <input
                    type="text"
                    inputMode="decimal"
                    value={startInput}
                    onChange={(e) => {
                      const cleaned = e.target.value.replace(/[^0-9.]/g, "");
                      setStartInput(cleaned);
                      if (/^\d+(\.\d+)*$/.test(cleaned)) {
                        onChange({
                          numbering: { ...table.numbering, start: cleaned },
                        });
                      }
                    }}
                    onBlur={() => {
                      if (!/^\d+(\.\d+)*$/.test(startInput)) {
                        setStartInput(table.numbering.start);
                      }
                    }}
                    aria-label="Starting number for the first item, e.g. 5.1"
                    title="The exact number of the first item (e.g. 5.1 → 5.1, 5.1.1)"
                    className={`${FIELD} w-20`}
                  />
                </>
              )}
              <label className="ml-1 flex items-center gap-1.5 border-l border-border pl-3 text-xs text-text-secondary">
                <input
                  type="checkbox"
                  checked={table.breakAfter[0] === true}
                  onChange={(e) =>
                    onChange({ breakAfter: e.target.checked ? [true] : [] })
                  }
                  aria-label="Blank line between top-level groups"
                  title="Add a blank line after each top-level group (this section)"
                  className="accent-[var(--accent)]"
                />
                Blank line between top-level groups
              </label>
              <label className="flex items-center gap-1.5 text-xs text-text-secondary">
                <input
                  type="checkbox"
                  checked={table.pageBreakBefore === true}
                  onChange={(e) => onChange({ pageBreakBefore: e.target.checked })}
                  aria-label="Start each top-level group on a new Word page"
                  title="Insert a Word page break before each top-level group (this section)"
                  className="accent-[var(--accent)]"
                />
                New page per group
              </label>
            </div>

            {/* One compact row per indent level. Every control stays VISIBLE (no
                progressive disclosure) so the whole format is adjustable at a
                glance, and 9 levels still fit on screen. */}
            <div className={SUB}>Per level</div>
            <div>
              <div className="flex items-center gap-2 px-1 pb-1 text-[10px] font-semibold uppercase tracking-wide text-muted">
                <span className="w-[18px] shrink-0" />
                <span className="min-w-0 flex-1">Level</span>
                {table.numbering.mode !== "off" && (
                  <span className="w-40 shrink-0">
                    {table.numbering.mode === "multilevel" ? "Number" : "Marker"}
                  </span>
                )}
                <span className="w-20 shrink-0 text-center">Heading</span>
                <span
                  className="w-24 shrink-0 text-center text-accent-text"
                  title="Shared by DEPTH across every section — changing level 2's look restyles level 2 everywhere"
                >
                  Look · all tables
                </span>
              </div>
              <div className="flex flex-col gap-0.5">
                {pivotLevels.map((bucket, i) => {
                  const label = headers[bucket[0]] || `Level ${i + 1}`;
                  // Stacked siblings share this level, so name them here too —
                  // the row is honestly "level 2 = Country + State", not just the
                  // bucket's first field.
                  const stacked = bucket
                    .slice(1)
                    .map((c) => headers[c] || `Column ${c + 1}`);
                  const idx = levelIdxForBucket(i);
                  const lv = appearance.levelStyles[idx] ?? DEFAULT_LEVEL;
                  const isHeading = table.headingLevels[i] === true;
                  const { k: headingK, auto, explicit, skips } = headingInfo[i];
                  return (
                    <div
                      key={i}
                      className="flex items-center gap-2 rounded p-1 hover:bg-surface-alt"
                    >
                      <span className="flex h-[18px] w-[18px] shrink-0 items-center justify-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] font-semibold tabular-nums text-muted">
                        {i + 1}
                      </span>
                      <span className="flex min-w-0 flex-1 items-baseline gap-1.5 truncate text-xs">
                        <span className="truncate text-text-secondary">
                          {label}
                        </span>
                        {stacked.length > 0 && (
                          <span
                            className="truncate text-[11px] text-muted"
                            title={`Level ${i + 1} also holds: ${stacked.join(", ")}`}
                          >
                            + {stacked.join(", ")}
                          </span>
                        )}
                      </span>
                      {/* Marker cell — hidden in "off": a per-level marker select
                          in "custom", or a show-number checkbox in "multilevel". */}
                      {table.numbering.mode !== "off" && (
                        <div className="w-40 shrink-0">
                          {isHeading ? (
                            // A Word-heading level's own marker/number is
                            // suppressed (Word numbers those rows on paste), so an
                            // active-looking control here would lie.
                            <span
                              className="block truncate text-xs italic text-muted"
                              title="Word supplies this level's number on paste — the app marker/number is suppressed"
                            >
                              Word numbers
                            </span>
                          ) : table.numbering.mode === "multilevel" ? (
                            <label className="flex items-center gap-1.5 text-xs text-text-secondary">
                              <input
                                type="checkbox"
                                checked={table.numbering.levels[i] !== false}
                                onChange={(e) => {
                                  const next = [...table.numbering.levels];
                                  next[i] = e.target.checked;
                                  onChange({
                                    numbering: {
                                      ...table.numbering,
                                      levels: next,
                                    },
                                  });
                                }}
                                aria-label={`Show the number on level ${i + 1}`}
                                className="accent-[var(--accent)]"
                              />
                              Show number
                            </label>
                          ) : (
                            // Split marker: TYPE (counter/symbol glyph) +
                            // DELIMITER (glued after a counter; disabled for a
                            // symbol/none type, which takes no suffix).
                            (() => {
                              const spec =
                                markerSpecs[i] ?? defaultMarkerSpec(i + 1);
                              const counter = isCounterType(spec.type);
                              return (
                                <div className="flex items-center gap-1">
                                  <select
                                    value={spec.type}
                                    onChange={(e) =>
                                      setMarkerSpec(i, {
                                        type: e.target.value as MarkerType,
                                      })
                                    }
                                    aria-label={`Marker type for level ${i + 1}`}
                                    className={`${FIELD} min-w-0 flex-1 px-1`}
                                  >
                                    {MARKER_TYPE_OPTIONS.map((o) => (
                                      <option key={o.value} value={o.value}>
                                        {o.label}
                                      </option>
                                    ))}
                                  </select>
                                  <select
                                    value={spec.delim}
                                    disabled={!counter}
                                    onChange={(e) =>
                                      setMarkerSpec(i, {
                                        delim: e.target.value as MarkerDelim,
                                      })
                                    }
                                    aria-label={`Marker delimiter for level ${i + 1}`}
                                    title={
                                      counter
                                        ? "Delimiter after the number/letter"
                                        : "No delimiter — this marker is a symbol"
                                    }
                                    className={`${FIELD} w-16 px-1 text-center disabled:opacity-40`}
                                  >
                                    {MARKER_DELIM_OPTIONS.map((o) => (
                                      <option key={o.value} value={o.value}>
                                        {o.label}
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              );
                            })()
                          )}
                        </div>
                      )}
                      {/* Word heading: checkbox + a rank chip-dropdown. Auto = the
                          contiguous rank (never skipping); picking H1–H9 overrides
                          it. Amber when a manual pick skips a rank. */}
                      <div className="flex w-20 shrink-0 items-center justify-center gap-1">
                        <input
                          type="checkbox"
                          checked={isHeading}
                          onChange={(e) => {
                            const next = [...table.headingLevels];
                            next[i] = e.target.checked;
                            onChange({ headingLevels: next });
                          }}
                          aria-label={
                            isHeading
                              ? `${label} maps to Word Heading ${headingK} — uncheck to make it body text`
                              : `Map ${label} to a Word heading (Navigation pane, collapsible)`
                          }
                          title="Map to a Word heading (Navigation pane, collapsible)"
                          className="accent-[var(--accent)]"
                        />
                        {/* Rendered always but visibility-hidden when unchecked, so
                            the column doesn't jump when a box is ticked. */}
                        <select
                          value={explicit}
                          onChange={(e) => {
                            const next = [...(table.headingRanks ?? [])];
                            next[i] = Number(e.target.value); // 0 = auto
                            onChange({ headingRanks: next });
                          }}
                          aria-label={`Word heading rank for ${label} (Auto = Heading ${auto})`}
                          title={
                            skips
                              ? `Pastes as Word "Heading ${headingK}" — skips a rank (Word flags gapped outlines)`
                              : `Pastes as Word "Heading ${headingK}". Auto follows the checked levels above; pick H1–H9 to pin it.`
                          }
                          className={`h-6 rounded-sm border px-0.5 text-[10px] font-semibold tabular-nums outline-none transition-colors ${
                            isHeading ? "" : "invisible"
                          } ${
                            skips
                              ? "border-[color:color-mix(in_srgb,var(--warning)_55%,transparent)] bg-[color:color-mix(in_srgb,var(--warning)_12%,transparent)] text-[color:var(--warning)]"
                              : "border-border-strong bg-accent-subtle text-accent-text hover:border-accent"
                          }`}
                        >
                          <optgroup label="Auto (follows levels above)">
                            <option value={0}>H{auto}</option>
                          </optgroup>
                          <optgroup label="Pin to">
                            {Array.from({ length: 9 }, (_, n) => (
                              <option key={n + 1} value={n + 1}>
                                H{n + 1}
                              </option>
                            ))}
                          </optgroup>
                        </select>
                      </div>
                      {/* Look — the SAME shared LookControl the Section Title uses
                          (swatch + Aa▾ popover: Font/Size/B/I/U). This restyles the
                          WHOLE line (label AND value), unlike the Rows list's
                          Aa/B/I/U, which touch only the label before the colon. */}
                      <div className="flex w-24 shrink-0 justify-center">
                        <LookControl
                          value={lv}
                          onChange={(patch) => appearance.onLevelChange(idx, patch)}
                          open={openPop === `lvl-${i}`}
                          onOpenChange={(o) => setOpenPop(o ? `lvl-${i}` : null)}
                          label={label}
                          bodyFontOption={appearance.bodyFont}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </section>
      )}
    </div>
  );
}

export const TableCard = memo(TableCardInner);

"use client";

import { memo, useMemo, useState } from "react";
import { headingLevel } from "@/lib/clipboard";
import { defaultMarker, type MarkerKind } from "@/lib/renderers";
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

/** Marker styles offered per indent level, with a sample label. */
const MARKER_OPTIONS: { kind: MarkerKind; label: string }[] = [
  { kind: "decimal", label: "1." },
  { kind: "paren", label: "1)" },
  { kind: "upperAlpha", label: "A." },
  { kind: "lowerAlpha", label: "a." },
  { kind: "upperRoman", label: "I." },
  { kind: "lowerRoman", label: "i." },
  { kind: "bullet", label: "• bullet" },
  { kind: "dash", label: "– dash" },
  { kind: "none", label: "None" },
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
  // Which per-level "Look" popover (by level index) is open, or null.
  const [openPop, setOpenPop] = useState<number | null>(null);
  // A plain-text filter over the Add-fields pool, only shown once it's long enough
  // to need one (a wide spreadsheet can have dozens of columns).
  const [fieldFilter, setFieldFilter] = useState("");

  const { pivotLevels } = table;
  // A body bucket's GLOBAL level-style slot = its indent depth, +1 when a Section
  // title sits above it (mirrors the renderer's walk(nodes, 2, 1)); clamp to the
  // 9-entry chart. Keeps each matrix "Look" cell driving the same data-level the
  // Word export uses, so preview and paste stay in sync.
  const titleOffset = table.sectionTitle.trim() ? 1 : 0;
  const levelIdxForBucket = (b: number) => Math.min(8, b + titleOffset);
  // The Word heading a checked level maps to: its depth offset by the TITLE's
  // heading number — but only when a title is actually emitted (mirrors the
  // renderer's `bodyHeadingBase = title ? titleLevel : 0`, clamped at 9). Shown as
  // an H-number chip in the matrix so the mapping is visible, not guessed.
  const headingBase = table.sectionTitle.trim()
    ? headingLevel(headingStyleName)
    : 0;
  const headingKForBucket = (b: number) => Math.min(headingBase + b + 1, 9);

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
    const out: { col: number; b: number; fi: number }[] = [];
    let fi = 0;
    pivotLevels.forEach((bucket, b) =>
      bucket.forEach((col) => {
        out.push({ col, b, fi });
        fi++;
      }),
    );
    return out;
  }, [pivotLevels]);

  const guides = useMemo(() => rowGuides(placed), [placed]);

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
  // Toggle button (Aa/B/I/U/sort): brand-tinted "pressed-in" when active.
  const tgl = (active: boolean) =>
    `flex h-7 min-w-7 items-center justify-center rounded border px-1.5 text-xs transition-colors disabled:opacity-40 ${
      active
        ? "border-accent bg-accent-subtle text-accent-text"
        : "border-transparent text-text-secondary hover:bg-surface-alt hover:text-foreground"
    }`;

  return (
    <div className="flex flex-col gap-4">
      {/* ---- SECTION HEADER group (the level-0 title, its own home) ---------- */}
      <section className={CARD}>
        <h2 className={GROUP_HEADER}>Section Header</h2>
        <div className="flex flex-col gap-2.5 text-sm">
          <div className="flex flex-wrap items-center gap-2">
            <span className="w-16 shrink-0 text-xs text-text-secondary">
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
          <div className="flex flex-wrap items-center gap-2">
            <span className="w-16 shrink-0 text-xs text-text-secondary">
              Heading
            </span>
            <select
              value={headingStyleName}
              onChange={(e) => onHeadingStyleChange(e.target.value)}
              aria-label="Word heading style for the title"
              className={FIELD}
            >
              {HEADING_OPTIONS.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
            <span className={BADGE}>All tables</span>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <span className="w-16 shrink-0 text-xs text-text-secondary">
              Look
            </span>
            <select
              value={title.font}
              onChange={(e) => onTitleChange({ font: e.target.value })}
              aria-label="Title font"
              className={FIELD}
            >
              {FONTS.map((f) => (
                <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
            <input
              type="text"
              inputMode="numeric"
              value={title.sizeInput}
              onChange={(e) =>
                onTitleChange({ sizeInput: e.target.value.replace(/[^0-9]/g, "") })
              }
              aria-label="Title size in points"
              className={`${FIELD} w-14 text-center`}
            />
            <div className="flex gap-0.5">
              <button
                type="button"
                aria-pressed={title.bold}
                title="Bold the title"
                onClick={() => onTitleChange({ bold: !title.bold })}
                className={`${tgl(title.bold)} font-bold`}
              >
                B
              </button>
              <button
                type="button"
                aria-pressed={title.italic}
                title="Italicize the title"
                onClick={() => onTitleChange({ italic: !title.italic })}
                className={`${tgl(title.italic)} italic`}
              >
                I
              </button>
              <button
                type="button"
                aria-pressed={title.underline}
                title="Underline the title"
                onClick={() => onTitleChange({ underline: !title.underline })}
                className={`${tgl(title.underline)} underline`}
              >
                U
              </button>
            </div>
            <input
              type="color"
              value={title.color}
              onChange={(e) => onTitleChange({ color: e.target.value })}
              aria-label="Title color"
              className="h-7 w-8 cursor-pointer rounded border border-border-strong"
            />
            <span className={BADGE}>All tables</span>
          </div>
        </div>
      </section>

      {/* ---- LEVELS group (the per-level structure + numbering + markers) ---- */}
      <section className={CARD}>
        <h2 className={GROUP_HEADER}>Levels</h2>

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
                      onChange({ pivotLevels: addField(pivotLevels, col) })
                    }
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
                  <span>Structure</span>
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
                        but Word (and this preview) show only <strong>9</strong> —
                        levels 10+ collapse onto level 9. Outdent (&#9668;) or remove
                        fields until you&apos;re at 9 or fewer.
                      </span>
                    ) : (
                      <span>
                        <strong>Maximum depth reached.</strong> Word supports 9 indent
                        levels — new fields now stack at level 9 instead of nesting
                        deeper.
                      </span>
                    )}
                  </p>
                )}
                {/* gap-0 so the vertical tree guides stay continuous row-to-row */}
                <div className="flex flex-col">
                  {placed.map(({ col, b, fi }) => {
                    const name = headers[col] || `Column ${col + 1}`;
                    const lf = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
                    const g = guides[fi];
                    return (
                      <div
                        key={col}
                        className="flex min-h-[36px] items-center gap-1.5 rounded px-1 hover:bg-surface-alt"
                      >
                        {/* Explorer tree-line guides: b ancestor cells + 1 elbow */}
                        <span aria-hidden className="flex self-stretch">
                          {g.ancestors.map((cont, a) => (
                            <GuideCell key={a} kind={cont ? "line" : "blank"} />
                          ))}
                          <GuideCell kind={g.last ? "corner" : "tee"} />
                        </span>

                        <span className="flex h-[18px] w-[18px] shrink-0 items-center justify-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] font-semibold tabular-nums text-muted">
                          {b + 1}
                        </span>
                        <span className="rounded-[3px] bg-[color:color-mix(in_srgb,var(--accent)_8%,transparent)] px-2 py-0.5 text-xs font-semibold text-foreground">
                          {name}
                        </span>
                        <span className="ml-1 flex items-center gap-0.5">
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
                        <span className="ml-auto flex items-center">
                          <button
                            type="button"
                            aria-label={`Outdent ${name}`}
                            disabled={!canOutdent(pivotLevels, fi)}
                            onClick={() =>
                              onChange({ pivotLevels: outdentField(pivotLevels, fi) })
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
                              onChange({ pivotLevels: indentField(pivotLevels, fi) })
                            }
                            className={icon}
                          >
                            ►
                          </button>
                          <button
                            type="button"
                            aria-label={`Move ${name} up`}
                            disabled={fi === 0}
                            onClick={() =>
                              onChange({ pivotLevels: moveField(pivotLevels, fi, -1) })
                            }
                            className={icon}
                          >
                            ▲
                          </button>
                          <button
                            type="button"
                            aria-label={`Move ${name} down`}
                            disabled={fi === placed.length - 1}
                            onClick={() =>
                              onChange({ pivotLevels: moveField(pivotLevels, fi, 1) })
                            }
                            className={icon}
                          >
                            ▼
                          </button>
                          <button
                            type="button"
                            aria-label={`Remove ${name}`}
                            onClick={() =>
                              onChange({ pivotLevels: removeField(pivotLevels, fi) })
                            }
                            className={`${icon} ml-0.5`}
                          >
                            ✕
                          </button>
                        </span>
                      </div>
                    );
                  })}
                </div>
              </>
            )}

            {/* Numbering */}
            {pivotLevels.length > 0 && (
              <>
                <div className={SUB}>Markers</div>
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
                      onChange={(e) =>
                        onChange({ pageBreakBefore: e.target.checked })
                      }
                      aria-label="Start each top-level group on a new Word page"
                      title="Insert a Word page break before each top-level group (this section)"
                      className="accent-[var(--accent)]"
                    />
                    New page per group
                  </label>
                </div>
              </>
            )}

            {/* Per-level matrix: one row per indent level (all PER-TABLE) — its
                marker (or "show number" toggle when numbering is on) and Word-heading
                mapping, aligned to the SAME numbered pill as Structure so a level
                reads as one object across the card. The "blank line between top-level
                groups" toggle now lives once beside Markers (no per-level Gap column);
                every column here is this-section scope. */}
            {pivotLevels.length > 0 && (
              <>
                <div className={SUB}>Per-level</div>
                <div>
                  <div>
                    <div className="flex items-center gap-2 px-1 pb-1 text-[10px] font-semibold uppercase tracking-wide text-muted">
                      <span className="w-[18px] shrink-0" />
                      <span className="min-w-0 flex-1">Level</span>
                      {table.numbering.mode !== "off" && (
                        <span className="w-32 shrink-0">
                          {table.numbering.mode === "multilevel"
                            ? "Number"
                            : "Marker"}
                        </span>
                      )}
                      <span className="w-14 shrink-0 text-center">Heading</span>
                      <span className="w-20 shrink-0 text-center text-accent-text">
                        Look
                      </span>
                    </div>
                    <div className="flex flex-col gap-0.5">
                      {pivotLevels.map((bucket, i) => {
                        const label = headers[bucket[0]] || `Level ${i + 1}`;
                        const idx = levelIdxForBucket(i);
                        const lv = appearance.levelStyles[idx] ?? DEFAULT_LEVEL;
                        const isHeading = table.headingLevels[i] === true;
                        const headingK = headingKForBucket(i);
                        return (
                          <div
                            key={i}
                            className="flex items-center gap-2 rounded p-1 hover:bg-surface-alt"
                          >
                            <span className="flex h-[18px] w-[18px] shrink-0 items-center justify-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] font-semibold tabular-nums text-muted">
                              {i + 1}
                            </span>
                            <span className="min-w-0 flex-1 truncate text-xs text-text-secondary">
                              {label}
                            </span>
                            {/* #/Mark cell — hidden in "off" mode: a per-level
                                marker select in "custom" mode, or a show-number
                                checkbox in "multilevel" mode. */}
                            {table.numbering.mode !== "off" && (
                              <div className="w-32 shrink-0">
                                {isHeading ? (
                                  // A Word-heading level's own marker/number is
                                  // suppressed (Word numbers those rows on paste),
                                  // so an active-looking control here would lie.
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
                                <select
                                  value={table.markers[i] ?? defaultMarker(i + 1)}
                                  onChange={(e) => {
                                    const next = [...table.markers];
                                    next[i] = e.target.value as MarkerKind;
                                    onChange({ markers: next });
                                  }}
                                  aria-label={`Marker for level ${i + 1}`}
                                  className={`${FIELD} w-full`}
                                >
                                  {MARKER_OPTIONS.map((o) => (
                                    <option key={o.kind} value={o.kind}>
                                      {o.label}
                                    </option>
                                  ))}
                                </select>
                              )}
                              </div>
                            )}
                            {/* Word heading: checkbox + the RESULTING heading
                                number as a chip (H2 under a Heading-1 title, etc.)
                                so the mapping is visible at a glance. */}
                            <label className="flex w-14 shrink-0 items-center justify-center gap-1">
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
                                    : `Map ${label} to Word Heading ${headingK} (Navigation pane, collapsible)`
                                }
                                title={`Maps to Word "Heading ${headingK}" (Navigation pane, collapsible)`}
                                className="accent-[var(--accent)]"
                              />
                              {isHeading && (
                                <span
                                  className="rounded-sm bg-accent-subtle px-1 text-[10px] font-semibold tabular-nums text-accent-text"
                                  title={`Pastes as Word "Heading ${headingK}"`}
                                >
                                  H{headingK}
                                </span>
                              )}
                            </label>
                            {/* Look (GLOBAL — all tables): color swatch inline +
                                font/size/bold in a popover. Tinted so the
                                this-section vs all-tables boundary stays obvious. */}
                            <div className="relative flex w-20 shrink-0 items-center justify-center gap-1 rounded bg-accent-subtle py-0.5">
                              <input
                                type="color"
                                value={lv.color}
                                onChange={(e) =>
                                  appearance.onLevelChange(idx, {
                                    color: e.target.value,
                                  })
                                }
                                aria-label={`${label} text color (all tables)`}
                                title="Text color · all tables"
                                className="h-6 w-6 cursor-pointer rounded border border-border-strong"
                              />
                              <button
                                type="button"
                                onClick={() =>
                                  setOpenPop((p) => (p === i ? null : i))
                                }
                                aria-expanded={openPop === i}
                                aria-label={`${label} font, size, bold (all tables)`}
                                title="Font · size · bold · all tables"
                                className="rounded border border-border-strong bg-surface px-1 text-[11px] font-semibold text-text-secondary transition-colors hover:border-accent hover:text-accent-text"
                              >
                                Aa▾
                              </button>
                              <Popover
                                open={openPop === i}
                                onClose={() => setOpenPop(null)}
                              >
                                <div className="flex flex-col gap-2">
                                  <div className="text-[11px] font-semibold uppercase tracking-wide text-muted">
                                    {label} look
                                    <span className="ml-1 font-normal normal-case text-accent-text">
                                      · all tables
                                    </span>
                                  </div>
                                  <label className="flex items-center justify-between gap-3">
                                    Font
                                    <select
                                      value={lv.font}
                                      onChange={(e) =>
                                        appearance.onLevelChange(idx, {
                                          font: e.target.value,
                                        })
                                      }
                                      aria-label={`${label} font`}
                                      className={FIELD}
                                    >
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
                                      value={lv.sizeInput}
                                      onChange={(e) =>
                                        appearance.onLevelChange(idx, {
                                          sizeInput: e.target.value.replace(
                                            /[^0-9]/g,
                                            "",
                                          ),
                                        })
                                      }
                                      aria-label={`${label} size in points`}
                                      className={`${FIELD} w-16`}
                                    />
                                  </label>
                                  <label className="flex items-center gap-1.5">
                                    <input
                                      type="checkbox"
                                      checked={lv.bold}
                                      onChange={(e) =>
                                        appearance.onLevelChange(idx, {
                                          bold: e.target.checked,
                                        })
                                      }
                                      className="accent-[var(--accent)]"
                                    />
                                    Bold
                                  </label>
                                </div>
                              </Popover>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              </>
            )}
          </div>
        )}
      </section>
    </div>
  );
}

export const TableCard = memo(TableCardInner);

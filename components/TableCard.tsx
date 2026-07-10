"use client";

import { memo, useMemo, useState } from "react";
import { defaultMarker, type MarkerKind } from "@/lib/renderers";
import { DEFAULT_FIELD_LABEL, type FieldLabel } from "@/lib/types";
import {
  addField,
  canIndent,
  canOutdent,
  indentField,
  moveField,
  outdentField,
  removeField,
  unusedColumns,
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

const FONTS = [
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

type Props = {
  table: TableState;
  onChange: (patch: Partial<TableState>) => void;
  /** Global Word-heading style the TITLE maps to ("" = None). Shared. */
  headingStyleName: string;
  onHeadingStyleChange: (value: string) => void;
  /** Shared title look (font/size/color + bold/italic/underline). */
  title: TitleInput;
  onTitleChange: (patch: Partial<TitleInput>) => void;
};

function TableCardInner({
  table,
  onChange,
  headingStyleName,
  onHeadingStyleChange,
  title,
  onTitleChange,
}: Props) {
  // Start-number field held as a string so it can be cleared/retyped (the card is
  // keyed by table id, so this resets per table).
  const [startInput, setStartInput] = useState(String(table.numbering.start));

  const { grid, pivotLevels } = table;

  const headers = useMemo(
    () => (grid[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  const unused = useMemo(
    () => unusedColumns(headers.length, pivotLevels),
    [headers.length, pivotLevels],
  );

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
          {headingStyleName !== "" && (
            <p className="text-xs text-muted">
              Mapped to <strong>{headingStyleName}</strong> — this section joins the
              Word outline (Navigation pane) and Word supplies its heading{" "}
              <strong>number</strong>. Your <strong>Look</strong> above
              (font/size/color/<strong>B</strong>/<em>I</em>/
              <span className="underline">U</span>) still applies on top. Note: if
              the destination heading is ALL-CAPS, Word keeps the uppercase.
            </p>
          )}
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
            <div className={`${SUB} !mt-0`}>Add fields</div>
            <div className="flex flex-wrap items-center gap-1.5">
              {unused.length === 0 ? (
                <span className="text-xs text-muted">All fields added.</span>
              ) : (
                unused.map((col) => (
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
                <div className={SUB}>Structure</div>
                <div className="flex flex-col gap-0.5">
                  {placed.map(({ col, b, fi }) => {
                    const name = headers[col] || `Column ${col + 1}`;
                    const lf = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
                    return (
                      <div
                        key={col}
                        className="flex items-center gap-1.5 rounded p-1 hover:bg-surface-alt"
                        style={{ marginLeft: `${b * 1.25}rem` }}
                      >
                        <span className="flex h-[18px] w-[18px] items-center justify-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] font-semibold tabular-nums text-muted">
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
                <div className={SUB}>Numbering</div>
                <div className="flex flex-wrap items-center gap-2">
                  <select
                    value={table.numbering.mode}
                    onChange={(e) =>
                      onChange({
                        numbering: {
                          ...table.numbering,
                          mode: e.target.value as "off" | "multilevel",
                        },
                      })
                    }
                    aria-label="Multilevel numbering mode"
                    className={FIELD}
                  >
                    <option value="off">Off</option>
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
                </div>
              </>
            )}

            {/* Markers (hidden when multilevel numbering replaces them) */}
            {pivotLevels.length > 0 && table.numbering.mode !== "multilevel" && (
              <>
                <div className={SUB}>Markers</div>
                <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5">
                  {pivotLevels.map((_, i) => (
                    <label
                      key={i}
                      className="flex items-center gap-1.5 text-xs text-text-secondary"
                    >
                      Lv {i + 1}
                      <select
                        value={table.markers[i] ?? defaultMarker(i + 1)}
                        onChange={(e) => {
                          const next = [...table.markers];
                          next[i] = e.target.value as MarkerKind;
                          onChange({ markers: next });
                        }}
                        aria-label={`Marker for indent level ${i + 1}`}
                        className={FIELD}
                      >
                        {MARKER_OPTIONS.map((o) => (
                          <option key={o.kind} value={o.kind}>
                            {o.label}
                          </option>
                        ))}
                      </select>
                    </label>
                  ))}
                </div>
              </>
            )}

            {/* Number level (only when numbering is on) */}
            {pivotLevels.length > 0 && table.numbering.mode === "multilevel" && (
              <>
                <div className={SUB}>Number level</div>
                <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5">
                  {pivotLevels.map((_, i) => (
                    <label
                      key={i}
                      className="flex items-center gap-1.5 text-xs text-text-secondary"
                    >
                      <input
                        type="checkbox"
                        checked={table.numbering.levels[i] !== false}
                        onChange={(e) => {
                          const next = [...table.numbering.levels];
                          next[i] = e.target.checked;
                          onChange({
                            numbering: { ...table.numbering, levels: next },
                          });
                        }}
                        aria-label={`Show the number on indent level ${i + 1}`}
                        className="accent-[var(--accent)]"
                      />
                      Lv {i + 1}
                    </label>
                  ))}
                </div>
              </>
            )}

            {/* Word heading */}
            {pivotLevels.length > 0 && (
              <>
                <div className={SUB}>Word heading (nav pane, collapsible)</div>
                <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5">
                  {pivotLevels.map((bucket, i) => (
                    <label
                      key={i}
                      className="flex items-center gap-1.5 text-xs text-text-secondary"
                    >
                      <input
                        type="checkbox"
                        checked={table.headingLevels[i] === true}
                        onChange={(e) => {
                          const next = [...table.headingLevels];
                          next[i] = e.target.checked;
                          onChange({ headingLevels: next });
                        }}
                        aria-label={`Make indent level ${i + 1} a Word heading`}
                        className="accent-[var(--accent)]"
                      />
                      {headers[bucket[0]] || `Lv ${i + 1}`}
                    </label>
                  ))}
                </div>
              </>
            )}

            {/* Blank line after */}
            {pivotLevels.length > 0 && (
              <>
                <div className={SUB}>Blank line after each</div>
                <select
                  value={table.breakAfter.findIndex((b) => b) + 1}
                  onChange={(e) => {
                    const lvl = Number(e.target.value);
                    onChange({
                      breakAfter:
                        lvl === 0
                          ? []
                          : pivotLevels.map((_, i) => i === lvl - 1),
                    });
                  }}
                  aria-label="Add a blank line after each group at this level"
                  className={`${FIELD} w-fit`}
                >
                  <option value={0}>(none)</option>
                  {pivotLevels.map((bucket, i) => (
                    <option key={i} value={i + 1}>
                      {headers[bucket[0]] || `Level ${i + 1}`}
                    </option>
                  ))}
                </select>
              </>
            )}
          </div>
        )}
      </section>
    </div>
  );
}

export const TableCard = memo(TableCardInner);

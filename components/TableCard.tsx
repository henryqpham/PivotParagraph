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

type Props = {
  table: TableState;
  onChange: (patch: Partial<TableState>) => void;
};

function TableCardInner({ table, onChange }: Props) {
  // Start-number field held as a string so it can be cleared/retyped (the card is
  // keyed by table id, so this resets per table); committed to numbering.start only
  // when it parses to a valid dotted decimal, and re-synced to the committed value
  // on blur if left empty/invalid.
  const [startInput, setStartInput] = useState(String(table.numbering.start));

  const { grid, pivotLevels } = table;

  // Header row, for the field names.
  const headers = useMemo(
    () => (grid[0] ? grid[0].map((c) => (c == null ? "" : String(c))) : []),
    [grid],
  );

  // Fields not yet placed in any indent bucket (the "Add fields" pool).
  const unused = useMemo(
    () => unusedColumns(headers.length, pivotLevels),
    [headers.length, pivotLevels],
  );

  // The placed fields flattened, each tagged with its bucket b, within-bucket k,
  // and flat index fi (the argument the structure helpers expect).
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

  // Toggle part of one field's label look (keyed by grid column).
  function patchLabel(col: number, patch: Partial<FieldLabel>) {
    const cur = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
    onChange({
      fieldLabels: { ...table.fieldLabels, [col]: { ...cur, ...patch } },
    });
  }

  // Cycle one field's sort: off → asc → desc → off (keyed by grid column, so it
  // persists across remove/re-add). Off removes the key; asc/desc set it.
  function cycleSort(col: number) {
    const cur = table.sortDirs[col];
    const next: Record<number, "asc" | "desc"> = { ...table.sortDirs };
    if (cur === undefined) next[col] = "asc";
    else if (cur === "asc") next[col] = "desc";
    else delete next[col];
    onChange({ sortDirs: next });
  }

  const btn =
    "px-1 text-foreground/40 transition-colors hover:text-foreground disabled:opacity-20 disabled:hover:text-foreground/40";
  // Per-field label toggle (Aa / B / I / U): highlighted when active.
  const tgl = (active: boolean) =>
    `flex h-5 w-5 items-center justify-center rounded border text-[10px] transition-colors disabled:opacity-25 ${
      active
        ? "border-foreground/40 bg-foreground/10 text-foreground"
        : "border-foreground/15 text-foreground/50 hover:bg-foreground/5"
    }`;

  return (
    <div className="flex flex-col gap-3 rounded-xl border border-foreground/15 p-4">
      <div className="flex flex-wrap items-center gap-4">
        <label className="flex items-center gap-2 text-sm text-foreground/60">
          Section title
          <input
            type="text"
            value={table.sectionTitle}
            onChange={(e) => onChange({ sectionTitle: e.target.value })}
            placeholder="e.g. Fruit Database"
            className="w-48 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
          />
        </label>
        {/* Numbering: app-drawn static multilevel numbers (5, 5.1, 5.1.1) as plain
            body text. When on, the number replaces the per-level marker. */}
        <label className="flex items-center gap-1.5 text-sm text-foreground/60">
          Numbering
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
            className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
          >
            <option value="off">Off</option>
            <option value="multilevel">Multilevel numbers</option>
          </select>
        </label>
        {table.numbering.mode === "multilevel" && (
          <label className="flex items-center gap-1.5 text-sm text-foreground/60">
            Start
            <input
              type="text"
              inputMode="decimal"
              value={startInput}
              onChange={(e) => {
                const cleaned = e.target.value.replace(/[^0-9.]/g, "");
                setStartInput(cleaned);
                // Commit a valid dotted-decimal (e.g. "5", "5.1", "5.1.2"); it is
                // the exact number of the first top-level item.
                if (/^\d+(\.\d+)*$/.test(cleaned)) {
                  onChange({
                    numbering: { ...table.numbering, start: cleaned },
                  });
                }
              }}
              onBlur={() => {
                // Left empty/invalid: restore the committed value.
                if (!/^\d+(\.\d+)*$/.test(startInput)) {
                  setStartInput(table.numbering.start);
                }
              }}
              aria-label="Starting number for the first item, e.g. 5.1"
              title="The exact number of the first item (e.g. 5.1 → 5.1, 5.1.1, 5.1.1.1)"
              className="w-20 rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
            />
          </label>
        )}
      </div>

      {headers.length > 0 && (
        <div className="flex flex-col gap-3 text-sm text-foreground/70">
          {/* Add fields: the unused pool. Clicking appends a new deepest level. */}
          <div className="flex flex-col gap-1.5">
            <span className="text-foreground/60">
              Add fields{" "}
              <span className="text-muted">
                (click to add to the outline)
              </span>
              :
            </span>
            <div className="flex flex-wrap items-center gap-1.5">
              {unused.length === 0 ? (
                <span className="text-xs text-muted">
                  All fields added.
                </span>
              ) : (
                unused.map((col) => (
                  <button
                    key={col}
                    type="button"
                    onClick={() =>
                      onChange({ pivotLevels: addField(pivotLevels, col) })
                    }
                    className="rounded-md border border-foreground/15 px-2 py-1 text-xs text-foreground/60 transition-colors hover:bg-foreground/5"
                  >
                    + {headers[col] || `Column ${col + 1}`}
                  </button>
                ))
              )}
            </div>
          </div>

          {/* Structure: ordered indent buckets. ◄ ► change level (stack = same
              indent), ▲ ▼ reorder, ✕ remove. */}
          {placed.length > 0 && (
            <div className="flex flex-col gap-1.5">
              <span className="text-foreground/60">
                Structure{" "}
                <span className="text-muted">
                  (◄ ► set indent — stack fields at one level to show them
                  together)
                </span>
                :
              </span>
              <div className="flex flex-col gap-1">
                {placed.map(({ col, b, fi }) => {
                  const name = headers[col] || `Column ${col + 1}`;
                  const lf = table.fieldLabels[col] ?? DEFAULT_FIELD_LABEL;
                  return (
                    <div
                      key={col}
                      className="flex items-center gap-1.5"
                      style={{ marginLeft: `${b * 1.25}rem` }}
                    >
                      <span className="flex h-4 w-4 items-center justify-center rounded bg-foreground/10 text-[10px] font-semibold tabular-nums text-foreground/60">
                        {b + 1}
                      </span>
                      <span className="rounded bg-foreground/5 px-2 py-0.5 text-xs text-foreground">
                        {name}
                      </span>
                      {/* Label look: show/hide "Field:", bold/italic/underline it. */}
                      <span className="flex items-center gap-0.5">
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
                      </span>
                      {/* Sort: off ↕ → asc ↑ → desc ↓ → off, ordering the sibling
                          groups at this field's level. */}
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
                            title={`Sort groups by ${name} (currently ${dirLabel}; click to cycle off / ascending / descending)`}
                            onClick={() => cycleSort(col)}
                            className={tgl(dir !== undefined)}
                          >
                            {glyph}
                          </button>
                        );
                      })()}
                      <span className="flex items-center leading-none">
                        <button
                          type="button"
                          aria-label={`Outdent ${name}`}
                          disabled={!canOutdent(pivotLevels, fi)}
                          onClick={() =>
                            onChange({
                              pivotLevels: outdentField(pivotLevels, fi),
                            })
                          }
                          className={btn}
                        >
                          ◄
                        </button>
                        <button
                          type="button"
                          aria-label={`Indent ${name}`}
                          disabled={!canIndent(pivotLevels, fi)}
                          onClick={() =>
                            onChange({
                              pivotLevels: indentField(pivotLevels, fi),
                            })
                          }
                          className={btn}
                        >
                          ►
                        </button>
                        <button
                          type="button"
                          aria-label={`Move ${name} up`}
                          disabled={fi === 0}
                          onClick={() =>
                            onChange({
                              pivotLevels: moveField(pivotLevels, fi, -1),
                            })
                          }
                          className={btn}
                        >
                          ▲
                        </button>
                        <button
                          type="button"
                          aria-label={`Move ${name} down`}
                          disabled={fi === placed.length - 1}
                          onClick={() =>
                            onChange({
                              pivotLevels: moveField(pivotLevels, fi, 1),
                            })
                          }
                          className={btn}
                        >
                          ▼
                        </button>
                        <button
                          type="button"
                          aria-label={`Remove ${name}`}
                          onClick={() =>
                            onChange({
                              pivotLevels: removeField(pivotLevels, fi),
                            })
                          }
                          className={`${btn} ml-0.5`}
                        >
                          ✕
                        </button>
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Markers: one picker per indent level (bucket). Multilevel numbering
              replaces every marker, so swap the pickers for a note when it's on. */}
          {pivotLevels.length > 0 && (
            <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5 text-sm text-foreground/70">
              <span className="text-foreground/60">
                Markers{" "}
                <span className="text-muted">(per indent level)</span>:
              </span>
              {table.numbering.mode === "multilevel" ? (
                <span className="text-xs text-muted">
                  Replaced by Numbering — the number prefixes each group.
                </span>
              ) : (
                pivotLevels.map((_, i) => (
                  <label
                    key={i}
                    className="flex items-center gap-1.5 text-xs text-foreground/60"
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
                      className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
                    >
                      {MARKER_OPTIONS.map((o) => (
                        <option key={o.kind} value={o.kind}>
                          {o.label}
                        </option>
                      ))}
                    </select>
                  </label>
                ))
              )}
            </div>
          )}

          {/* Number level: per-level show/hide of the multilevel number (only when
              numbering is on). Unchecking a level leaves its line plain while the
              path still compounds by depth. */}
          {pivotLevels.length > 0 && table.numbering.mode === "multilevel" && (
            <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5 text-sm text-foreground/70">
              <span className="text-foreground/60">
                Number level{" "}
                <span className="text-muted">(show the number per level)</span>:
              </span>
              {pivotLevels.map((_, i) => (
                <label
                  key={i}
                  className="flex items-center gap-1.5 text-xs text-foreground/60"
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
                  />
                  Lv {i + 1}
                </label>
              ))}
            </div>
          )}

          {/* Word heading: map a level's rows to a destination Heading style so
              they enter the document outline (nav pane + collapsible, flush-left).
              Word supplies their number, so the app number/marker is suppressed on
              those rows. Usually just the top level or two; labelled by field name. */}
          {pivotLevels.length > 0 && (
            <div className="flex flex-wrap items-center gap-x-3 gap-y-1.5 text-sm text-foreground/70">
              <span className="text-foreground/60">
                Word heading{" "}
                <span className="text-muted">(in nav, collapsible)</span>:
              </span>
              {pivotLevels.map((bucket, i) => (
                <label
                  key={i}
                  className="flex items-center gap-1.5 text-xs text-foreground/60"
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
                  />
                  {headers[bucket[0]] || `Lv ${i + 1}`}
                </label>
              ))}
            </div>
          )}

          {/* Blank line after: pick which level's groups get a trailing blank line
              (a separator → a gap before the next group at that level). */}
          {pivotLevels.length > 0 && (
            <label className="flex flex-wrap items-center gap-1.5 text-sm text-foreground/70">
              <span className="text-foreground/60">Blank line after each</span>
              <select
                value={table.breakAfter.findIndex((b) => b) + 1}
                onChange={(e) => {
                  const lvl = Number(e.target.value); // 0 = none, else the level
                  onChange({
                    breakAfter:
                      lvl === 0 ? [] : pivotLevels.map((_, i) => i === lvl - 1),
                  });
                }}
                aria-label="Add a blank line after each group at this level"
                className="rounded-md border border-foreground/20 px-2 py-1 text-sm text-foreground"
              >
                <option value={0}>(none)</option>
                {pivotLevels.map((bucket, i) => (
                  <option key={i} value={i + 1}>
                    {headers[bucket[0]] || `Level ${i + 1}`}
                  </option>
                ))}
              </select>
            </label>
          )}
        </div>
      )}
    </div>
  );
}

export const TableCard = memo(TableCardInner);

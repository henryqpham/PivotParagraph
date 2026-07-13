"use client";

import { useState } from "react";
import type { TableState } from "./tableModel";

/**
 * The left "Sections" rail: one row per pasted table, edited one at a time like
 * tabs — each section is its OWN independent Word document (Copy section copies
 * just the active one; nothing is stacked). A row can be selected (makes it the
 * active/edited section), reordered by **drag and drop** (grab a row and drop it
 * where you want; keyboard users grab the ⋮ handle and press ↑/↓), and removed.
 */
export function SectionsRail({
  tables,
  activeId,
  onSelect,
  onRemove,
  onReorder,
}: {
  tables: TableState[];
  activeId: string | null;
  onSelect: (id: string) => void;
  onRemove: (id: string) => void;
  /** Move the section at `from` to index `to` (drag-drop or keyboard). */
  onReorder: (from: number, to: number) => void;
}) {
  // The row currently being dragged, and the row it's hovered over (for the drop
  // indicator). Both are flat indices into `tables`, or null when idle.
  const [dragIndex, setDragIndex] = useState<number | null>(null);
  const [overIndex, setOverIndex] = useState<number | null>(null);
  // A plain-text filter, only shown once the rail gets long enough to need one.
  // Non-matching rows render as null rather than being spliced out of the array,
  // so `i` below stays the row's TRUE index into `tables` — drag-and-drop reorder
  // (which operates on those indices) stays correct even while filtered.
  const [filter, setFilter] = useState("");
  const showFilter = tables.length > 8;
  const filterLower = filter.trim().toLowerCase();
  const labelOf = (t: TableState, i: number) => t.sectionTitle.trim() || `Table ${i + 1}`;
  const matchesFilter = (t: TableState, i: number) =>
    !filterLower || labelOf(t, i).toLowerCase().includes(filterLower);
  const visibleCount = tables.reduce(
    (n, t, i) => n + (matchesFilter(t, i) ? 1 : 0),
    0,
  );

  function endDrag() {
    setDragIndex(null);
    setOverIndex(null);
  }

  function drop(to: number) {
    if (dragIndex !== null && dragIndex !== to) onReorder(dragIndex, to);
    endDrag();
  }

  // Compact icon button; hidden until the row is hovered or something inside it is
  // focused (keyboard users still reach it via focus-visible).
  const rowBtn =
    "grid h-6 w-5 shrink-0 place-items-center rounded text-xs leading-none text-muted opacity-0 transition-opacity hover:text-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100";

  return (
    <aside className="flex w-56 shrink-0 flex-col overflow-y-auto border-r border-border bg-surface-alt p-2">
      <div className="border-b border-border px-2 pb-1.5 pt-1 text-xs font-semibold uppercase tracking-wide text-muted">
        Sections
      </div>
      {showFilter && (
        <div className="relative mt-1.5 px-1">
          <input
            type="text"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="Filter sections…"
            aria-label="Filter sections by title"
            className="h-7 w-full rounded border border-border-strong bg-surface px-2 pr-6 text-xs text-foreground outline-none transition-colors focus:border-accent"
          />
          {filter && (
            <button
              type="button"
              onClick={() => setFilter("")}
              aria-label="Clear filter"
              className="absolute right-2 top-1/2 -translate-y-1/2 leading-none text-muted transition-colors hover:text-foreground"
            >
              &times;
            </button>
          )}
        </div>
      )}
      <div className="mt-1 flex flex-col gap-0.5">
        {showFilter && filterLower && visibleCount === 0 && (
          <p className="px-2 py-1 text-xs text-muted">
            No sections match &ldquo;{filter.trim()}&rdquo;.
          </p>
        )}
        {tables.map((t, i) => {
          if (!matchesFilter(t, i)) return null;
          const isActive = activeId === t.id;
          const label = labelOf(t, i);
          const isDragging = dragIndex === i;
          const isDropTarget =
            overIndex === i && dragIndex !== null && dragIndex !== i;
          // With splice-move semantics, a section dragged DOWN lands after the
          // target and one dragged UP lands before it — so the drop line sits on
          // the matching edge (bottom vs top) and matches where the row ends up.
          const dropBelow = isDropTarget && dragIndex !== null && dragIndex < i;
          return (
            <div
              key={t.id}
              draggable
              onDragStart={(e) => {
                setDragIndex(i);
                e.dataTransfer.effectAllowed = "move";
                // Firefox needs data set for a drag to initiate.
                e.dataTransfer.setData("text/plain", t.id);
              }}
              onDragOver={(e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = "move";
                if (overIndex !== i) setOverIndex(i);
              }}
              onDrop={(e) => {
                e.preventDefault();
                drop(i);
              }}
              onDragEnd={endDrag}
              className={`group relative flex h-9 cursor-grab items-center gap-1 rounded pl-1 pr-1 text-sm transition-colors active:cursor-grabbing ${
                isDragging ? "opacity-40" : ""
              } ${
                isActive
                  ? "bg-accent-subtle font-semibold text-accent-text after:absolute after:left-0 after:top-1.5 after:bottom-1.5 after:w-[3px] after:rounded-full after:bg-accent after:content-['']"
                  : "text-text-secondary hover:bg-surface"
              }`}
            >
              {isDropTarget && (
                <span
                  aria-hidden
                  className={`pointer-events-none absolute inset-x-1 h-0.5 rounded-full bg-accent ${
                    dropBelow ? "bottom-0" : "top-0"
                  }`}
                />
              )}
              {/* Drag handle — also the keyboard reorder affordance (↑/↓). */}
              <button
                type="button"
                aria-label={`Reorder ${label}. Use arrow up and down keys to move.`}
                title="Drag to reorder"
                onKeyDown={(e) => {
                  if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
                  // Swallow the arrow always (even at a boundary) so it reorders
                  // instead of scrolling the rail/page; no-op past the ends.
                  e.preventDefault();
                  if (e.key === "ArrowUp" && i > 0) onReorder(i, i - 1);
                  else if (e.key === "ArrowDown" && i < tables.length - 1)
                    onReorder(i, i + 1);
                }}
                className="grid h-6 w-4 shrink-0 cursor-grab place-items-center rounded text-xs leading-none text-muted opacity-40 transition-opacity hover:text-foreground focus-visible:opacity-100 group-hover:opacity-100"
              >
                &#8942;
              </button>
              <span className="grid h-[18px] w-[18px] shrink-0 place-items-center rounded-[3px] bg-[color:color-mix(in_srgb,var(--muted)_16%,transparent)] text-[11px] tabular-nums text-muted">
                {i + 1}
              </span>
              <button
                type="button"
                onClick={() => onSelect(t.id)}
                aria-current={isActive ? "true" : undefined}
                className="min-w-0 flex-1 truncate text-left"
                title={label}
              >
                {label}
              </button>
              <button
                type="button"
                onClick={() => onRemove(t.id)}
                aria-label={`Remove ${label}`}
                title="Remove this section"
                className={`${rowBtn} hover:text-danger`}
              >
                &times;
              </button>
            </div>
          );
        })}
      </div>
      <p className="mt-2 px-2 text-[11px] leading-snug text-muted">
        {tables.length === 0
          ? "No sections yet — paste a table to add one."
          : "Each section is its own document. Paste another table to add one; drag to reorder."}
      </p>
    </aside>
  );
}

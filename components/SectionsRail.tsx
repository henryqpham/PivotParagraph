"use client";

import type { TableState } from "./tableModel";

/**
 * The left "Sections" rail: one row per pasted table, in paste (= document) order.
 * A row can be selected (makes it the active/edited section), reordered up/down,
 * and removed. Section order IS the exported document order (`Copy all` and the
 * combined preview both follow this array), so reordering here restructures the
 * output — the reorder/remove controls appear on hover or keyboard focus to keep
 * the resting state calm.
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
  onReorder: (id: string, dir: -1 | 1) => void;
}) {
  // Compact icon button; hidden until the row is hovered or something inside it is
  // focused (keyboard users still reach it via focus-visible).
  const rowBtn =
    "grid h-6 w-5 shrink-0 place-items-center rounded text-xs leading-none text-muted opacity-0 transition-opacity hover:text-foreground focus-visible:opacity-100 group-hover:opacity-100 group-focus-within:opacity-100 disabled:opacity-0";

  return (
    <aside className="flex w-56 shrink-0 flex-col overflow-y-auto border-r border-border bg-surface-alt p-2">
      <div className="border-b border-border px-2 pb-1.5 pt-1 text-xs font-semibold uppercase tracking-wide text-muted">
        Sections
      </div>
      <div className="mt-1 flex flex-col gap-0.5">
        {tables.map((t, i) => {
          const isActive = activeId === t.id;
          const label = t.sectionTitle.trim() || `Table ${i + 1}`;
          return (
            <div
              key={t.id}
              className={`group relative flex h-9 items-center gap-1 rounded pl-3 pr-1 text-sm transition-colors ${
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
                onClick={() => onSelect(t.id)}
                aria-current={isActive ? "true" : undefined}
                className="min-w-0 flex-1 truncate text-left"
                title={label}
              >
                {label}
              </button>
              <button
                type="button"
                onClick={() => onReorder(t.id, -1)}
                disabled={i === 0}
                aria-label={`Move ${label} up`}
                title="Move section up"
                className={rowBtn}
              >
                &#9650;
              </button>
              <button
                type="button"
                onClick={() => onReorder(t.id, 1)}
                disabled={i === tables.length - 1}
                aria-label={`Move ${label} down`}
                title="Move section down"
                className={rowBtn}
              >
                &#9660;
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
          : "Paste another table to add a section. Drag order with ▲▼; that's the order Copy all stacks them."}
      </p>
    </aside>
  );
}

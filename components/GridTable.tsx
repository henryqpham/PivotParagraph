import type { Grid } from "@/lib/types";

/** Cap the rows we render so a huge paste can't freeze the tab; the note below
 *  the table says how many were hidden (no silent truncation). */
const MAX_ROWS = 500;

/**
 * A faithful, spreadsheet-style view of the RAW parsed grid — exactly what was
 * pasted, before any pivot restructuring. Row 0 is shown as the header (the app
 * treats it as the field names everywhere), the rest as data rows with a
 * row-number gutter. Wide tables (the whole reason this app exists) scroll
 * horizontally inside their own container; the header + gutter stay pinned.
 *
 * All cell text is rendered as JSX children (React escapes it), so there is no
 * injection surface even though the values are user-derived.
 */
export function GridTable({ grid }: { grid: Grid }) {
  if (!grid || grid.length === 0) {
    return (
      <p className="text-sm text-foreground/60">
        Nothing was parsed from this paste. Switch to JSON to inspect the raw
        data.
      </p>
    );
  }

  const cols = grid.reduce((max, row) => Math.max(max, row.length), 0);
  const [header = [], ...dataRows] = grid;
  const shown = dataRows.slice(0, MAX_ROWS);
  const hidden = dataRows.length - shown.length;
  const text = (c: unknown) => (c == null ? "" : String(c));

  // Header cells: pinned to the top on vertical scroll. No `z` in the base so the
  // corner cell can override it without a Tailwind ordering conflict.
  const thBase =
    "sticky top-0 whitespace-nowrap border-b border-border-strong bg-surface-alt px-2.5 py-1.5 text-left font-semibold text-foreground";
  // Row-number gutter: pinned to the left on horizontal scroll (opaque so cells
  // scroll behind it).
  const gutter =
    "sticky left-0 z-10 border-r border-border bg-surface-alt px-2 py-1.5 text-right align-top text-[11px] tabular-nums text-muted";
  const td =
    "min-w-[3.5rem] max-w-[22rem] whitespace-pre-wrap break-words border-b border-border px-2.5 py-1.5 align-top text-foreground";

  return (
    <div className="flex flex-col gap-2">
      <div className="text-sm text-foreground/60">
        {grid.length} row{grid.length === 1 ? "" : "s"} &times; {cols} column
        {cols === 1 ? "" : "s"}{" "}
        <span className="text-muted">&middot; row 1 is the header</span>
      </div>
      <div className="overflow-x-auto rounded-lg border border-border shadow-[var(--shadow-2)]">
        <table className="w-full border-collapse text-xs">
          <thead>
            <tr>
              {/* top-left corner: pinned on BOTH axes, above header + gutter */}
              <th
                aria-hidden
                className={`${thBase} left-0 z-20 border-r border-border`}
              />
              {Array.from({ length: cols }, (_, c) => (
                <th key={c} scope="col" className={thBase}>
                  {text(header[c]) || (
                    <span className="font-normal text-muted">Column {c + 1}</span>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shown.map((row, ri) => (
              <tr key={ri} className={ri % 2 ? "bg-surface-alt" : "bg-surface"}>
                <td className={gutter}>{ri + 1}</td>
                {Array.from({ length: cols }, (_, c) => (
                  <td key={c} className={td}>
                    {text(row[c])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {hidden > 0 && (
        <p className="text-xs text-muted">
          Showing the first {MAX_ROWS} of {dataRows.length} data rows.
        </p>
      )}
    </div>
  );
}

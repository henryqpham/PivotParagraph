// Per-table model + pure derivation, shared by PasteInput (the parent) and
// TableCard. Lives in its own module so both can import the type and helpers
// without a parent <-> child import cycle.

import { rowsToPivotTree } from "@/lib/mapper";
import {
  DEFAULT_NUMBERING,
  renderPivotTree,
  type MarkerKind,
  type MarkerSpec,
  type NumberingConfig,
} from "@/lib/renderers";
import type { FieldLabel, Grid } from "@/lib/types";

/** One pasted table: its grid plus the per-table pivot config the user tweaks. */
export type TableState = {
  /** Stable id from a monotonic counter; used as the React key. */
  id: string;
  grid: Grid;
  /**
   * The pivot structure as ordered INDENT BUCKETS. `pivotLevels[b]` is indent
   * level b+1 and holds one or more grid column indices, in display (stacking)
   * order. Fields stacked in one bucket render at the same indent and merge as a
   * composite group. Columns in no bucket are unused (hidden).
   */
  pivotLevels: number[][];
  /**
   * Per-level marker as split TYPE + DELIMITER (index = level − 1; sparse →
   * default cycle). Read via `resolveMarkerSpecs(t)`, which also migrates the
   * LEGACY fused `markers?: MarkerKind[]` from older sessions/exports.
   */
  markerSpecs?: MarkerSpec[];
  /**
   * LEGACY fused marker per indent level (`"1."`/`"a)"`/…). Kept optional so old
   * `localStorage` sessions and `.json` imports still load; migrated to
   * `markerSpecs` on read by `resolveMarkerSpecs`. Never written by new code.
   */
  markers?: MarkerKind[];
  /**
   * Per-field label look, keyed by grid column index (so it persists when a field
   * is removed and re-added). Absent col → the default (label shown, plain).
   */
  fieldLabels: Record<number, FieldLabel>;
  /**
   * Per-field sort direction, keyed by grid column index (like `fieldLabels`, so
   * it persists across remove/re-add). It orders the SIBLING groups at that
   * field's indent level; absent col → off (first-seen order). When a bucket
   * stacks several sorted fields, the first in bucket order wins.
   */
  sortDirs: Record<number, "asc" | "desc">;
  /**
   * "Blank line after" toggle per indent level (index = level − 1, like
   * `markers`). When on, the renderer pushes one empty spacer paragraph after
   * that level's whole node-subtree. Sparse/short → off.
   */
  breakAfter: boolean[];
  /**
   * App-drawn static multilevel numbering for this table (off by default). When
   * on, the renderer prefixes each node's first line with its compounded number
   * path (5, 5.1, 5.1.1) as plain body text and suppresses that node's marker.
   */
  numbering: NumberingConfig;
  /**
   * "Make this level a Word heading" toggle per indent level (index = level − 1,
   * like `markers`). When on, that level's rows map to a destination `Heading K`
   * style (nav pane + collapsible, flush-left), and Word supplies their number
   * (the app's number/marker is suppressed on them). Usually just the top level
   * or two; sparse/short → not a heading.
   */
  headingLevels: boolean[];
  /**
   * OPTIONAL per-level override of WHICH Word heading a heading-checked level
   * maps to (index = level − 1, like `headingLevels`; value 1–9). 0 / absent /
   * null = AUTO: one rank under the PREVIOUS heading level (the title, then each
   * heading-checked level in depth order, pinned values included) — so an auto
   * level never skips a rank and autos re-derive around any pin. An explicit
   * value serves templates whose numbering keys off a specific rank (e.g.
   * several levels all mapping to `Heading 1`) and is the only way a rank gap
   * can occur. Optional so old sessions/imports stay compatible.
   */
  headingRanks?: number[];
  /** Optional title; the one Word heading, above the nested rows. */
  sectionTitle: string;
  /**
   * Which raw-grid row is the HEADER (0-based; default 0). Bumped when banner/title
   * rows sit above the real header — everything downstream reads `bodyGrid(t)`
   * (`grid.slice(headerRow)`) so those rows are ignored. Only drops leading ROWS,
   * so every column-keyed structure (pivotLevels/fieldLabels/sortDirs) stays valid.
   */
  headerRow?: number;
  /**
   * "Start each top-level group on a new Word page" (per table, off by default).
   * When on, the renderer marks each top-level group AFTER the first with
   * `data-break`, which `buildWordHtml` turns into `page-break-before:always`.
   */
  pageBreakBefore?: boolean;
};

/**
 * One row of the GLOBAL per-depth body-row style editor, held as UI inputs (size
 * a string so it can be cleared/retyped). Shared across ALL tables and keyed by
 * DEPTH, not by table. Lives here (a module both PasteInput and TableCard already
 * import) so the owner (PasteInput) and the docked "Appearance" editor (TableCard)
 * can share the type + default without a component import cycle.
 */
export type LevelInput = {
  color: string;
  font: string;
  sizeInput: string;
  bold: boolean;
  italic?: boolean;
  underline?: boolean;
};
/** Default look for an untouched level: 11pt black, font INHERITED from the
 *  document-wide Body font ("" = inherit — the ⚙ Document popover's Body font
 *  control would otherwise be a dead knob, since every paragraph carries an
 *  inline per-level font that used to hard-default to Arial). A level's Look
 *  popover can still PIN a specific font; "" un-pins back to Body font. */
export const DEFAULT_LEVEL: LevelInput = {
  color: "#000000",
  font: "",
  sizeInput: "11",
  bold: false,
  italic: false,
  underline: false,
};

/**
 * Maximum indent depth (number of buckets). Word supports only 9 list/heading
 * levels, and the render pipeline hard-caps at 9 everywhere downstream: the
 * renderer clamps `data-level` to `Math.min(level, 9)`, `RenderedPreview` builds
 * exactly 9 `data-level` style rules, and `buildWordHtml`'s regexes only match
 * `data-level="[1-9]"` / `data-heading="[1-9]"`. So a 10th bucket would render at
 * level 9 in the preview and could be DROPPED from the Word output — the builder
 * must never create a structure deeper than this, or it diverges from the preview.
 */
export const MAX_LEVELS = 9;

// ---------------------------------------------------------------------------
// Pure helpers on the bucket structure (`number[][]`). All return a fresh array
// and never mutate. A "flattened index" `fi` is a field's position in
// `levels.flat()`. Bucket invariant: every bucket is non-empty.
// ---------------------------------------------------------------------------

/** All placed grid columns, in flat (top-to-bottom) order. */
function placedColumns(levels: number[][]): number[] {
  return levels.flat();
}

/** Grid columns NOT placed in any bucket (the "unused" pool), in column order. */
export function unusedColumns(width: number, levels: number[][]): number[] {
  const placed = new Set(placedColumns(levels));
  const cols: number[] = [];
  for (let i = 0; i < width; i++) if (!placed.has(i)) cols.push(i);
  return cols;
}

/** Locate a field by flattened index: its bucket `b` and within-bucket index `k`. */
function locate(levels: number[][], fi: number): { b: number; k: number } | null {
  let n = 0;
  for (let b = 0; b < levels.length; b++) {
    if (fi < n + levels[b].length) return { b, k: fi - n };
    n += levels[b].length;
  }
  return null;
}

/**
 * Add a column to the structure. Normally it appends a new deepest single-field
 * bucket, but once the depth cap (`MAX_LEVELS`) is reached it STACKS into the
 * deepest existing bucket instead of creating a 10th level — so every column can
 * still be placed without building a structure the preview/Word can't show.
 */
export function addField(levels: number[][], col: number): number[][] {
  if (levels.length >= MAX_LEVELS) {
    const last = levels.length - 1;
    return levels.map((bk, i) => (i === last ? [...bk, col] : bk));
  }
  return [...levels, [col]];
}

/** Remove the field at flat index `fi`; drop its bucket if it becomes empty. */
export function removeField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at) return levels;
  const { b, k } = at;
  const bucket = levels[b].filter((_, i) => i !== k);
  if (bucket.length === 0) return levels.filter((_, i) => i !== b);
  return levels.map((bk, i) => (i === b ? bucket : bk));
}

/**
 * Whether ► (indent) is allowed for the field at `fi`: it must not be first in its
 * bucket (nothing to split off), AND we must be under the depth cap — indenting
 * splits one bucket into two, so it can't push the structure past `MAX_LEVELS`.
 */
export function canIndent(levels: number[][], fi: number): boolean {
  const at = locate(levels, fi);
  return at != null && at.k > 0 && levels.length < MAX_LEVELS;
}

/**
 * Indent the field at `fi` one level deeper: split its bucket at `k` so the field
 * and everything after it in that bucket move into a NEW bucket one level deeper.
 * No-op unless `canIndent`.
 */
export function indentField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at || at.k === 0) return levels;
  const { b, k } = at;
  const head = levels[b].slice(0, k);
  const tail = levels[b].slice(k);
  return [...levels.slice(0, b), head, tail, ...levels.slice(b + 1)];
}

/** Whether ◄ (outdent) is allowed for `fi`: first of its bucket and bucket > 0. */
export function canOutdent(levels: number[][], fi: number): boolean {
  const at = locate(levels, fi);
  return at != null && at.k === 0 && at.b > 0;
}

/**
 * Outdent the field at `fi` one level shallower by merging its bucket into the
 * previous one (the field and the rest of its bucket join the previous level).
 * No-op unless `canOutdent`.
 */
export function outdentField(levels: number[][], fi: number): number[][] {
  const at = locate(levels, fi);
  if (!at || at.k !== 0 || at.b === 0) return levels;
  const { b } = at;
  const merged = [...levels[b - 1], ...levels[b]];
  return [...levels.slice(0, b - 1), merged, ...levels.slice(b + 1)];
}

/**
 * Move the field at `fi` up (`dir = -1`) or down (`dir = +1`) by swapping it with
 * its flat-order neighbor while keeping the bucket SHAPE fixed: the two fields
 * trade slots, so a field moved up adopts the upper slot's indent. Always valid
 * (a swap can't empty a bucket). No-op when the neighbor doesn't exist.
 */
export function moveField(levels: number[][], fi: number, dir: -1 | 1): number[][] {
  const flat = placedColumns(levels);
  const fj = fi + dir;
  if (fj < 0 || fj >= flat.length) return levels;
  const a = locate(levels, fi);
  const b = locate(levels, fj);
  if (!a || !b) return levels;
  return levels.map((bucket, bi) =>
    bucket.map((col, ki) => {
      if (bi === a.b && ki === a.k) return flat[fj];
      if (bi === b.b && ki === b.k) return flat[fi];
      return col;
    }),
  );
}

/**
 * Drop fully-empty columns from a freshly parsed grid — a column whose every cell
 * (header included) is blank is a spacer/padding artifact that would otherwise
 * clutter the Add-fields pool and merge as "(blank)" everywhere. Keeps column
 * ORDER and runs at INGEST, before any bucket captures column indices. Returns the
 * grid unchanged when nothing is empty, and never produces a 0-column grid.
 */
export function dropEmptyColumns(grid: Grid): Grid {
  if (grid.length === 0) return grid;
  const width = grid.reduce((m, r) => Math.max(m, r.length), 0);
  const keep: number[] = [];
  for (let c = 0; c < width; c++) {
    const nonEmpty = grid.some((row) => {
      const cell = row[c];
      return cell != null && String(cell).trim() !== "";
    });
    if (nonEmpty) keep.push(c);
  }
  if (keep.length === width || keep.length === 0) return grid;
  return grid.map((row) => keep.map((c) => row[c] ?? ""));
}

/**
 * The grid with any header-offset rows removed (rows above `headerRow`), so its row
 * 0 is the EFFECTIVE header. Read by the mapper, the builder's field list, and the
 * page-count stat, so a header offset applies everywhere at once. Only drops leading
 * ROWS — column indices (and thus pivotLevels/fieldLabels/sortDirs) stay valid.
 */
export function bodyGrid(t: TableState): Grid {
  const hr = t.headerRow ?? 0;
  return hr > 0 ? t.grid.slice(hr) : t.grid;
}

/**
 * Build a fresh, empty `TableState` for a pasted grid. The single source of the
 * default per-table config, shared by the paste handler and the example loader so
 * a new table always starts from the same shape (no fields placed, custom markers,
 * no title).
 */
export function newTable(id: string, grid: Grid): TableState {
  return {
    id,
    grid,
    pivotLevels: [],
    markerSpecs: [],
    fieldLabels: {},
    sortDirs: {},
    breakAfter: [],
    numbering: DEFAULT_NUMBERING,
    headingLevels: [],
    sectionTitle: "",
  };
}

/**
 * A small, deliberately WIDE demo grid (the shape the tool is built for): a few
 * grouping columns plus detail columns that don't fit a Word page side-by-side.
 * Used by the first-run "Try an example" button so a newcomer can see the
 * wide-table → nested-outline transform in one click, without their own
 * spreadsheet. Values are display strings (the parser returns strings).
 */
const EXAMPLE_GRID: Grid = [
  ["Region", "Country", "Product", "Units", "Revenue"],
  ["Americas", "United States", "Laptops", "120", "$240,000"],
  ["Americas", "United States", "Monitors", "90", "$90,000"],
  ["Americas", "Canada", "Laptops", "60", "$120,000"],
  ["EMEA", "Germany", "Laptops", "80", "$160,000"],
  ["EMEA", "Germany", "Monitors", "50", "$50,000"],
  ["EMEA", "France", "Tablets", "40", "$40,000"],
  ["APAC", "Japan", "Laptops", "100", "$200,000"],
  ["APAC", "Japan", "Tablets", "70", "$70,000"],
];

/**
 * A ready-to-explore example section: the wide demo grid pre-arranged into a
 * 3-level outline (Region → Country → Product), with Units/Revenue left in the
 * Add-fields pool so the pool is non-empty and teaches the "add a field" step.
 */
export function makeExampleTable(id: string): TableState {
  return {
    ...newTable(id, EXAMPLE_GRID),
    pivotLevels: [[0], [1], [2]],
    sectionTitle: "Sales by Region",
  };
}

/** Migrate one LEGACY fused `MarkerKind` to the split {type, delim} spec. */
export function migrateMarkerKind(k: MarkerKind): MarkerSpec {
  switch (k) {
    case "decimal":
      return { type: "decimal", delim: "dot" };
    case "paren":
      return { type: "decimal", delim: "paren" };
    case "upperAlpha":
      return { type: "upperAlpha", delim: "dot" };
    case "lowerAlpha":
      return { type: "lowerAlpha", delim: "dot" };
    case "upperRoman":
      return { type: "upperRoman", delim: "dot" };
    case "lowerRoman":
      return { type: "lowerRoman", delim: "dot" };
    // Symbol/none types ignore the delimiter, but default it to "dot" (like a
    // fresh level) so switching the Type dropdown to a counter yields "1." not
    // a bare "1".
    case "bullet":
      return { type: "bullet", delim: "dot" };
    case "dash":
      return { type: "dash", delim: "dot" };
    case "none":
      return { type: "none", delim: "dot" };
  }
}

/**
 * The per-level marker specs for a table: `markerSpecs` when present, else the
 * migrated LEGACY `markers`, else empty (renderer fills the default cycle). One
 * read-time normalizer so old sessions, imports, and new edits all funnel to the
 * same {type, delim}[] shape. Sparse holes are preserved (→ default per depth).
 */
export function resolveMarkerSpecs(t: TableState): MarkerSpec[] {
  // An EMPTY markerSpecs is "nothing set yet" (newTable writes `[]`), so fall
  // through to any legacy `markers` rather than treating `[]` as authoritative
  // and silently dropping a co-present legacy array (an import edge case).
  if (t.markerSpecs && t.markerSpecs.length > 0) return t.markerSpecs;
  return (t.markers ?? []).map((k) =>
    k ? migrateMarkerKind(k) : (undefined as unknown as MarkerSpec),
  );
}

/**
 * parse -> nest -> render for one table. The optional Section title maps to a
 * Word heading (in `buildWordHtml`); the body is always app-styled paragraphs.
 * Empty-guard first so a title never renders over nothing. Pure: the single source
 * used by both the card preview and the combined export, so both stay identical.
 *
 * The table's own config threads in: `sortDirs` reorders sibling groups in the
 * mapper post-pass, then `breakAfter` (per-level spacer), `numbering` (app-drawn
 * static multilevel numbers), and `headingLevels` (levels mapped to Word headings)
 * drive the renderer. `titleLevel` (the Word heading LEVEL the title maps to, 0 =
 * none) sets the body heading offset, so callers pass it from the shared style via
 * `headingLevel(headingStyleName)`.
 */
export function tableToHtml(
  t: TableState,
  titleLevel = 0,
  labelSep = ": ",
): string {
  const tree = rowsToPivotTree(bodyGrid(t), t.pivotLevels, t.sortDirs ?? {});
  if (tree.length === 0) return "";
  return renderPivotTree(
    tree,
    t.sectionTitle.trim() || undefined,
    resolveMarkerSpecs(t),
    t.fieldLabels ?? {},
    t.breakAfter ?? [],
    t.numbering ?? DEFAULT_NUMBERING,
    t.headingLevels ?? [],
    titleLevel,
    t.pageBreakBefore ?? false,
    t.headingRanks ?? [],
    labelSep,
  );
}

/**
 * A rough, at-a-glance summary of a section's shape — how many raw rows it started
 * from, how many indent levels it uses, and about how many Word pages the rendered
 * output would span. `pages` is a HEURISTIC (a flat lines-per-page constant, not
 * real text-wrapping/layout) — good enough to gauge "does this still fit a page,"
 * not a substitute for looking at the actual output.
 */
export type SectionStats = { rows: number; levels: number; pages: number };

// Usable body height on a US Letter page at buildWordHtml's fixed 1in margins
// (lib/clipboard.ts: `@page{size:8.5in 11in;margin:1in}`).
const PAGE_BODY_HEIGHT_IN = 9;
// Assumed body line height in inches (~11pt at this app's 115% line-height,
// rounded up a bit to account for text that wraps onto extra lines — a flat
// paragraph count can't see wrapping, so this errs toward more pages, not fewer).
const ASSUMED_LINE_HEIGHT_IN = 0.2;

export function estimateSectionStats(t: TableState, html: string): SectionStats {
  const rows = Math.max(0, bodyGrid(t).length - 1); // data rows, excluding header
  const levels = t.pivotLevels.length;
  const paragraphs = html === "" ? 0 : (html.match(/<p /g) ?? []).length;
  let pages =
    paragraphs === 0
      ? 0
      : Math.max(
          1,
          Math.ceil((paragraphs * ASSUMED_LINE_HEIGHT_IN) / PAGE_BODY_HEIGHT_IN),
        );
  // Each forced page break starts a fresh page, so the page count is at least the
  // number of top-level groups (breaks + 1) when "page break before" is on.
  const breaks = html === "" ? 0 : (html.match(/ data-break="1"/g) ?? []).length;
  if (breaks > 0) pages = Math.max(pages, breaks + 1);
  return { rows, levels, pages };
}

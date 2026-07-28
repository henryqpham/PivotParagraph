// Smart Arrange: a pure, client-side heuristic that PROPOSES an Excel-pivot
// arrangement (ordered indent buckets) from a freshly parsed grid. It exists so a
// newcomer can paste a wide table and get a sensible nested outline in one click,
// instead of hand-placing every field. The proposal is only a STARTING POINT the
// user then refines in the builder, so the heuristic favors robust, obvious
// signals over cleverness — per an explicit design review it deliberately does
// NOT attempt functional-dependency ("this column determines that one") detection,
// which is brittle on messy real-world data and produces confidently-wrong nesting.
//
// The whole thing is a pure function of the grid: no React, no DOM, no I/O, and it
// never throws (a bad/degenerate grid yields `null`, not an exception).

import type { Grid } from "@/lib/types";
import { MAX_LEVELS } from "@/components/tableModel";

/**
 * How a column reads to the heuristic:
 *   - `group`   — a categorical dimension you nest/indent BY (repeats, low-ish
 *                 cardinality text). These become the outline's indent levels.
 *   - `measure` — a numeric/currency column you'd report or sum (the "values"),
 *                 not something to group by. Stacked at the deepest level.
 *   - `detail`  — near-unique text (an ID, a note, free text): poor grouping keys,
 *                 so like measures they're stacked as leaf detail, not indented by.
 */
export type ColumnRole = "group" | "measure" | "detail";

export interface SmartArrangeResult {
  /**
   * Proposed ordered indent buckets of column indices, coarsest first. Each GROUP
   * column is its own bucket (its own indent level); every MEASURE/DETAIL column is
   * stacked into a single deepest "detail" bucket. Shape matches `TableState.
   * pivotLevels` so a caller can drop it straight in.
   */
  levels: number[][];
  /** Per-column role classification — every column of the grid is present. */
  roles: Record<number, ColumnRole>;
  /** Rough 0..1 confidence that this arrangement is actually useful. */
  confidence: number;
}

// A column reads as numeric only when a STRONG majority of its non-empty cells
// look like numbers — one stray "N/A" or a header-ish label shouldn't flip a
// measure column into a grouping dimension, nor vice-versa.
const NUMERIC_MAJORITY = 0.7;
// A text column is treated as near-unique "detail" (ID/free-text) once at least
// this fraction of its non-empty values are distinct — such a column makes a
// terrible grouping level (almost one group per row), so it's stacked as a leaf.
const NEAR_UNIQUE_RATIO = 0.9;

/**
 * Reduce a display string to its bare digits for a numeric test by removing the
 * punctuation spreadsheets sprinkle on numbers — currency `$`, grouping `,`,
 * percent `%`, accounting parens `()`, whitespace — plus a leading sign. We only
 * care WHETHER the cell is a number, never its value or sign, so stripping is
 * intentionally aggressive: "$1,234.50", "(1,234)", and "-5%" all collapse to
 * plain digits, while genuine text ("Germany") does not.
 */
function stripNumericNoise(value: string): string {
  return value.replace(/[$,%()\s]/g, "").replace(/^-/, "");
}

/**
 * Whether a trimmed display string reads as a number. `parseFloat` is deliberately
 * lenient (the spec calls for it), and the empty-string guard stops a blank cell
 * from counting as `0` — blanks are excluded from the numeric tally upstream, but
 * this keeps the helper honest if called directly.
 */
function looksNumeric(value: string): boolean {
  const bare = stripNumericNoise(value);
  if (bare === "") return false;
  return Number.isFinite(parseFloat(bare));
}

/** Coerce any raw grid cell to a trimmed display string; `null`/`undefined`/short
 *  rows all normalize to `""` so blanks are counted consistently. */
function cellText(cell: unknown): string {
  return String(cell ?? "").trim();
}

/**
 * Per-column signal used for classification: how many non-empty cells the column
 * has, how many DISTINCT non-empty values (cardinality), and whether it reads as
 * numeric overall. Computed once per column over the body rows.
 */
interface ColumnStats {
  nonEmpty: number;
  cardinality: number;
  numeric: boolean;
}

/**
 * Propose a nested pivot arrangement from a parsed grid, 100% client-side.
 *
 * The heuristic, over the BODY rows (`grid.slice(1)`; row 0 is the header):
 *   1. For each column measure its non-empty count, distinct-value count
 *      (cardinality), and whether a >=70% majority of non-empty cells look numeric.
 *   2. Classify: numeric → `measure`; else near-unique text (cardinality >= 90% of
 *      non-empty) → `detail`; else → `group` (a real categorical dimension).
 *   3. Keep the GROUP columns in ORIGINAL COLUMN ORDER (left-to-right = coarsest to
 *      finest, how pivot-shaped sheets are almost always authored — a far more
 *      robust hierarchy signal than distinct-count), and give each its own indent
 *      bucket. Cap the group buckets at `MAX_LEVELS - 1` so the deepest of the 9
 *      available levels is reserved for a leaf bucket; any group columns beyond the
 *      cap spill into that leaf bucket as detail.
 *   4. Stack every measure + detail column (in original column order) into that one
 *      deepest leaf bucket. With NO group columns there's nothing to indent by, so
 *      all measures/details land in a single flat bucket and confidence is low.
 *
 * Returns `null` when there's nothing sensible to propose — fewer than 2 data rows,
 * fewer than 2 columns, or no grouping dimension AND no measure (e.g. a couple of
 * all-unique ID columns). Never throws: any unexpected shape falls through to `null`.
 */
export function smartArrange(grid: Grid): SmartArrangeResult | null {
  try {
    if (!Array.isArray(grid) || grid.length === 0) return null;

    // Column count = the widest row (header included), mirroring how the ingest
    // step (`dropEmptyColumns`) measures width, so a body row shorter than the
    // header can't hide a column.
    const width = grid.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
    const bodyRows = grid.slice(1);

    // Nothing to arrange: need at least ~2 data rows to see any repetition, and at
    // least 2 columns for one to nest under another.
    if (bodyRows.length < 2 || width < 2) return null;

    // --- 1 & 2. Per-column stats + role classification ---------------------
    const roles: Record<number, ColumnRole> = {};
    const stats: ColumnStats[] = [];
    const groupCols: number[] = [];
    const measureCols: number[] = [];
    const detailCols: number[] = [];

    for (let col = 0; col < width; col++) {
      const distinct = new Set<string>();
      let nonEmpty = 0;
      let numericHits = 0;

      for (const row of bodyRows) {
        const text = cellText(row?.[col]);
        if (text === "") continue; // blanks don't count toward any signal
        nonEmpty++;
        distinct.add(text);
        if (looksNumeric(text)) numericHits++;
      }

      const numeric = nonEmpty > 0 && numericHits / nonEmpty >= NUMERIC_MAJORITY;
      stats[col] = { nonEmpty, cardinality: distinct.size, numeric };

      // Classify. Order matters: measure (numeric) wins first; then near-unique
      // text is detail; everything else — repeating categorical text, including a
      // single-value column (cardinality 1) — is a grouping dimension. An entirely
      // empty column (nonEmpty 0, cardinality 0) satisfies `0 >= 0.9 * 0` and so
      // reads as detail, harmless because it merely stacks as an empty leaf field
      // (and ingest's dropEmptyColumns removes such columns before we ever run).
      let role: ColumnRole;
      if (numeric) {
        role = "measure";
        measureCols.push(col);
      } else if (distinct.size >= NEAR_UNIQUE_RATIO * nonEmpty) {
        role = "detail";
        detailCols.push(col);
      } else {
        role = "group";
        groupCols.push(col);
      }
      roles[col] = role;
    }

    // Nothing to pivot: no grouping dimension AND no measure to report. A pile of
    // all-unique detail columns falls here too — there's no structure to propose.
    if (groupCols.length === 0 && measureCols.length === 0) return null;

    // --- 3. Group columns -> ordered indent buckets ------------------------
    // Keep the group columns in ORIGINAL COLUMN ORDER (left-to-right). Spreadsheets
    // are overwhelmingly authored coarse→fine from left to right (Region, Country,
    // Product, …), so column order is a far more robust hierarchy signal than
    // cardinality — sorting by distinct-count would nest Region → Product → Country
    // just because Product happens to have fewer values, inventing a false hierarchy
    // (exactly the failure mode an explicit design review warned about). `groupCols`
    // is already populated in column order, so no sort is needed.
    const sortedGroups = [...groupCols];

    // Reserve the deepest of the 9 levels for the measures/details leaf bucket, so
    // at most MAX_LEVELS-1 group columns get their own indent level; the rest spill
    // into that leaf bucket as detail (they still appear, just stacked not nested).
    const groupBucketCap = MAX_LEVELS - 1;
    const bucketedGroups = sortedGroups.slice(0, groupBucketCap);
    const overflowGroups = sortedGroups.slice(groupBucketCap);
    const groupBuckets: number[][] = bucketedGroups.map((col) => [col]);

    // --- 4. Deepest leaf bucket: measures + details (+ overflow groups) -----
    // Everything not given its own indent level stacks here, presented in original
    // column order so the leaf reads the way the source table did. A Set dedupes in
    // case a column ever landed in more than one list.
    const leafCols = Array.from(
      new Set<number>([...overflowGroups, ...measureCols, ...detailCols]),
    ).sort((a, b) => a - b);

    let levels: number[][];
    if (groupBuckets.length === 0) {
      // No grouping dimension at all — a single flat bucket of the measures/details
      // (reached only when there IS at least one measure, else we returned null).
      levels = [leafCols];
    } else if (leafCols.length === 0) {
      // Only grouping dimensions and nothing to stack beneath them.
      levels = groupBuckets;
    } else {
      levels = [...groupBuckets, leafCols];
    }

    // Enforce the 9-level cap end-to-end. The construction above already yields at
    // most MAX_LEVELS buckets, but clamp defensively so a future change can never
    // emit a structure deeper than the render pipeline (preview + buildWordHtml)
    // supports: fold any overflow buckets down into the deepest kept one, mirroring
    // how `addField` STACKS past the cap rather than creating a 10th level.
    if (levels.length > MAX_LEVELS) {
      const kept = levels.slice(0, MAX_LEVELS);
      const overflow = levels.slice(MAX_LEVELS).flat();
      kept[kept.length - 1] = [...kept[kept.length - 1], ...overflow];
      levels = kept;
    }
    // Drop any empty bucket (bucket invariant: every bucket is non-empty). Nothing
    // above should produce one, but guarding keeps the output well-formed no matter
    // what; a fully-empty result is treated as "nothing to propose".
    levels = levels.filter((bucket) => bucket.length > 0);
    if (levels.length === 0) return null;

    // --- 5. Confidence -----------------------------------------------------
    // Strongest when there are several grouping dimensions AND something to report
    // under them; medium with a single grouping dimension; low when there's no
    // grouping at all (a flat leaf list is barely a "pivot"). The >=2-groups /
    // no-measure case sits between: good nesting, but nothing obvious to report.
    const hasMeasure = measureCols.length > 0;
    let confidence: number;
    if (groupCols.length >= 2) confidence = hasMeasure ? 0.85 : 0.7;
    else if (groupCols.length === 1) confidence = 0.6;
    else confidence = 0.25;

    return { levels, roles, confidence };
  } catch {
    // Purely defensive: the logic above guards every access, but the contract is
    // "never throws", so any unforeseen shape resolves to "nothing to propose".
    return null;
  }
}

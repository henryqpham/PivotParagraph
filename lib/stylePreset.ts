// Style preset ("parser config"): save the app's FORMATTING once and replay it on
// the next document. A weekly user runs many different tables through the same
// house style (1.0 TITLE / 1.1 SUBHEADING / 1.1.1 TEXT — fonts, bolding, indent,
// markers/numbering, heading mapping, blank lines); a preset captures exactly
// that and nothing else.
//
// The key design decision: everything in a preset is either GLOBAL or keyed by
// LEVEL (indent depth) — never by grid column — so it transfers cleanly between
// documents whose headers have nothing in common. Structure (`pivotLevels`) and
// per-column settings (`fieldLabels`/`sortDirs`) are deliberately excluded: they
// are column-indexed, so they only make sense for a same-shaped table (that is
// the separate arrangement-reuse feature).
//
// Pure build/validate helpers only — file I/O and state application live in the
// owner (PasteInput), mirroring how lib/persistence.ts splits its duties.

import type { LevelInput, TableState } from "@/components/tableModel";
import { resolveMarkerSpecs } from "@/components/tableModel";
import type { TitleInput } from "@/components/TableCard";
import { DEFAULT_FIELD_LABEL, type FieldLabel } from "@/lib/types";
import {
  DEFAULT_NUMBERING,
  type MarkerSpec,
  type NumberingConfig,
} from "@/lib/renderers";

/** Everything a preset carries. Top half = the GLOBAL styling PasteInput owns;
 *  bottom half = the LEVEL-indexed per-table settings (applied to every section
 *  on import — the user's chosen scope). */
export type StylePresetData = {
  // Global styling (shared across sections already).
  levelStyles: LevelInput[];
  titleInput: TitleInput;
  headingStyleName: string;
  bodyFont: string;
  indentInput: string;
  labelSep?: string;
  lineSpacing?: string;
  // Level-indexed per-table settings — column-independent, so they replay on any
  // document. Sparse arrays are fine: every consumer reads them `?? default`.
  markerSpecs: MarkerSpec[];
  numbering: NumberingConfig;
  headingLevels: boolean[];
  headingRanks: number[];
  breakAfter: boolean[];
  pageBreakBefore: boolean;
  /**
   * The Rows card's per-field LABEL emphasis (show / bold / italic / underline),
   * re-keyed from grid columns to LEVEL POSITION so it transfers: entry [i][j] is
   * the label look of level i+1's j-th stacked field in the SOURCE arrangement.
   * `fieldLabels` itself is column-indexed and can't cross table shapes — this is
   * the level-keyed projection of it ("tie the preference to the row level").
   * Applied on import to each section's PLACED fields by the same [level][slot]
   * position (a slot past the source's stack depth reuses the last entry); a
   * field added after the import starts from the default label look. Optional so
   * presets saved before this field existed still load.
   */
  labelsByLevel?: FieldLabel[][];
};

/** The on-disk envelope. `kind` guards against feeding a WORKSPACE backup (or any
 *  other JSON) into the preset importer; `version` lets a future shape change
 *  invalidate old files loudly instead of half-loading them. */
export type StylePresetFile = {
  app: "pivotparagraph";
  kind: "style-preset";
  version: 1;
  exportedAt: string;
  data: StylePresetData;
};

/** The globals half of a preset, as PasteInput holds them. */
export type PresetGlobals = {
  levelStyles: LevelInput[];
  titleInput: TitleInput;
  headingStyleName: string;
  bodyFont: string;
  indentInput: string;
  labelSep: string;
  lineSpacing: string;
};

/**
 * Snapshot the current formatting as a preset file. The level-keyed half comes
 * from the ACTIVE section (the one whose formatting the user is looking at);
 * `resolveMarkerSpecs` normalizes any legacy fused markers on the way out so a
 * preset never carries the deprecated shape. With no active section the
 * level-keyed half exports as defaults — the globals alone are still worth saving.
 */
export function buildStylePreset(
  globals: PresetGlobals,
  active: TableState | null,
): StylePresetFile {
  return {
    app: "pivotparagraph",
    kind: "style-preset",
    version: 1,
    exportedAt: new Date().toISOString(),
    data: {
      levelStyles: globals.levelStyles,
      titleInput: globals.titleInput,
      headingStyleName: globals.headingStyleName,
      bodyFont: globals.bodyFont,
      indentInput: globals.indentInput,
      labelSep: globals.labelSep,
      lineSpacing: globals.lineSpacing,
      markerSpecs: active ? resolveMarkerSpecs(active) : [],
      numbering: active ? active.numbering : DEFAULT_NUMBERING,
      headingLevels: active ? active.headingLevels : [],
      headingRanks: active?.headingRanks ?? [],
      breakAfter: active ? active.breakAfter : [],
      pageBreakBefore: active?.pageBreakBefore ?? false,
      // Project the column-keyed label emphasis onto the arrangement's levels
      // (position [level][stack-slot]) so it survives a change of table shape.
      labelsByLevel: active
        ? active.pivotLevels.map((bucket) =>
            bucket.map((col) => ({
              ...DEFAULT_FIELD_LABEL,
              ...(active.fieldLabels[col] ?? {}),
            })),
          )
        : [],
    },
  };
}

/**
 * Re-apply a preset's level-keyed label emphasis to ONE table's column-keyed
 * `fieldLabels`: the field at level i, stack slot j takes entry [i][j] (a slot
 * past the source's stack depth reuses the last entry — the common case is a
 * deeper stack in the target). Fields in levels beyond the preset's depth, and
 * unplaced columns, keep their current settings. Pure; returns a fresh record.
 */
export function applyLabelsByLevel(
  table: TableState,
  labelsByLevel: FieldLabel[][] | undefined,
): Record<number, FieldLabel> {
  if (!labelsByLevel || labelsByLevel.length === 0)
    return table.fieldLabels ?? {};
  const next: Record<number, FieldLabel> = { ...(table.fieldLabels ?? {}) };
  table.pivotLevels.forEach((bucket, i) => {
    const levelLabels = labelsByLevel[i];
    if (!levelLabels || levelLabels.length === 0) return;
    bucket.forEach((col, j) => {
      const src = levelLabels[Math.min(j, levelLabels.length - 1)];
      next[col] = { ...DEFAULT_FIELD_LABEL, ...src };
    });
  });
  return next;
}

const NUMBERING_MODES = new Set(["off", "custom", "multilevel"]);

// Fallback title look for a preset missing/mangling `titleInput` — matches
// PasteInput's DEFAULT_TITLE (Calibri 11 black, Excel's default font).
const FALLBACK_TITLE: TitleInput = {
  font: "Calibri",
  sizeInput: "11",
  color: "#000000",
  bold: false,
  italic: false,
  underline: false,
};

/**
 * Validate a parsed JSON blob as a preset and normalize it to a full
 * `StylePresetData`, or return `null` when it isn't one (wrong kind, wrong
 * version, no data). Field-level reads are deliberately lenient — missing or
 * malformed fields fall back to safe defaults rather than rejecting the file —
 * because every downstream consumer already reads these shapes `?? default`; the
 * envelope check is the one strict gate.
 */
export function parseStylePreset(parsed: unknown): StylePresetData | null {
  if (!parsed || typeof parsed !== "object") return null;
  const env = parsed as Partial<StylePresetFile>;
  if (env.kind !== "style-preset" || env.version !== 1) return null;
  const d = env.data;
  if (!d || typeof d !== "object") return null;

  const arr = <T,>(v: unknown): T[] => (Array.isArray(v) ? (v as T[]) : []);
  const str = (v: unknown, fallback: string): string =>
    typeof v === "string" ? v : fallback;

  const rawNumbering = (d as StylePresetData).numbering;
  const numbering: NumberingConfig =
    rawNumbering &&
    typeof rawNumbering === "object" &&
    NUMBERING_MODES.has(rawNumbering.mode)
      ? {
          mode: rawNumbering.mode,
          start: str(rawNumbering.start, DEFAULT_NUMBERING.start),
          levels: arr<boolean>(rawNumbering.levels),
        }
      : DEFAULT_NUMBERING;

  const rawTitle = (d as StylePresetData).titleInput;
  const titleInput: TitleInput =
    rawTitle && typeof rawTitle === "object" && typeof rawTitle.font === "string"
      ? { ...FALLBACK_TITLE, ...rawTitle }
      : FALLBACK_TITLE;

  return {
    levelStyles: arr<LevelInput>((d as StylePresetData).levelStyles),
    titleInput,
    headingStyleName: str((d as StylePresetData).headingStyleName, ""),
    bodyFont: str((d as StylePresetData).bodyFont, "Calibri"),
    indentInput: str((d as StylePresetData).indentInput, "0.2"),
    labelSep: (d as StylePresetData).labelSep,
    lineSpacing: (d as StylePresetData).lineSpacing,
    markerSpecs: arr<MarkerSpec>((d as StylePresetData).markerSpecs),
    numbering,
    headingLevels: arr<boolean>((d as StylePresetData).headingLevels),
    headingRanks: arr<number>((d as StylePresetData).headingRanks),
    breakAfter: arr<boolean>((d as StylePresetData).breakAfter),
    pageBreakBefore: (d as StylePresetData).pageBreakBefore === true,
    // Lenient, shape-checked read: each entry must be an array of objects; each
    // label is normalized over the default so partial/hand-edited entries load.
    labelsByLevel: arr<unknown>((d as StylePresetData).labelsByLevel).map((lvl) =>
      arr<unknown>(lvl).map((l) => ({
        ...DEFAULT_FIELD_LABEL,
        ...(l && typeof l === "object" ? (l as Partial<FieldLabel>) : {}),
      })),
    ),
  };
}

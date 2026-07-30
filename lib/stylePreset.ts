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
    },
  };
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
  };
}

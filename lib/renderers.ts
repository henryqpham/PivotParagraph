import { DEFAULT_FIELD_LABEL, type FieldLabel, type PivotNode } from "./types";

/**
 * Escape the three markup-significant characters in user-derived text.
 *
 * `&` is replaced first so the entities introduced for `<`/`>` are not
 * double-escaped (e.g. `<` -> `&lt;`, never `&amp;lt;`).
 *
 * `"` / `'` are intentionally NOT escaped: every value we emit lands in element
 * text content, never inside an attribute (all `style="..."` attributes are
 * machine-generated constants), so the 3-character set is sufficient and keeps
 * the Word paste clean.
 */
function escapeHtml(s: string): string {
  return s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

/**
 * Build the label prefix `Field name: ` for one field, escaped and optionally
 * wrapped in `<b>`/`<i>`/`<u>` per its `FieldLabel`. Underline wraps ONLY the name
 * (so the separator isn't underlined); bold/italic wrap the whole "name<sep>"
 * label. The tags are inline runs (they survive a Word "Use Destination Styles"
 * paste). `sep` is the level's label separator — a FREE string typed by the
 * user (": " "; " " — " …), so it is escaped here before it touches markup.
 */
function wrapLabel(name: string, lf: FieldLabel, sep: string): string {
  const namePart = lf.underline
    ? `<u>${escapeHtml(name)}</u>`
    : escapeHtml(name);
  let html = `${namePart}${escapeHtml(sep)}`;
  if (lf.italic) html = `<i>${html}</i>`;
  if (lf.bold) html = `<b>${html}</b>`;
  return html;
}

/** Bijective base-26: 0 -> "a", 25 -> "z", 26 -> "aa". */
function toAlpha(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(97 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const ROMAN: [number, string][] = [
  [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
  [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
];
/** Lowercase roman numeral for n >= 1. */
function toRoman(n: number): string {
  let s = "";
  for (const [v, sym] of ROMAN) {
    while (n >= v) {
      s += sym;
      n -= v;
    }
  }
  return s;
}

/**
 * LEGACY fused marker style (type + delimiter in one token). Kept only so old
 * sessions/exports that stored `markers: MarkerKind[]` can migrate to the split
 * {type, delim} model (`migrateMarkerKind` in tableModel). New code uses
 * `MarkerType` + a free-string delimiter.
 */
export type MarkerKind =
  | "decimal" // 1.
  | "paren" // 1)
  | "upperAlpha" // A.
  | "lowerAlpha" // a.
  | "upperRoman" // I.
  | "lowerRoman" // i.
  | "bullet" // •
  | "dash" // –
  | "none"; // (no marker)

/** The counter/symbol TYPE for a marker, WITHOUT any trailing delimiter. */
export type MarkerType =
  | "decimal" // 1
  | "upperAlpha" // A
  | "lowerAlpha" // a
  | "upperRoman" // I
  | "lowerRoman" // i
  | "bullet" // •
  | "dash" // –
  | "none"; // (no marker)

/** A fully-specified marker = counter type + trailing delimiter. The delimiter
 *  is a FREE string (":" ";" ")" "--" …, ≤20 chars, escaped at compose time so
 *  it can never inject markup); legacy stored tokens "dot"/"paren"/"none" are
 *  migrated to their literals by `resolveMarkerSpecs`. Ignored for symbol/none
 *  types, which take no delimiter. */
export type MarkerSpec = { type: MarkerType; delim: string };

/** The counter/symbol GLYPH for a 0-based sibling `index` (no delimiter). */
function markerGlyph(type: MarkerType, index: number): string {
  switch (type) {
    case "decimal":
      return `${index + 1}`;
    case "upperAlpha":
      return toAlpha(index).toUpperCase();
    case "lowerAlpha":
      return toAlpha(index);
    case "upperRoman":
      return toRoman(index + 1).toUpperCase();
    case "lowerRoman":
      return toRoman(index + 1);
    case "bullet":
      return "•";
    case "dash":
      return "–";
    case "none":
      return "";
  }
}

/** Whether a type is a COUNTER (numbers/letters) that can take a delimiter, as
 *  opposed to a standalone symbol (bullet/dash) or none. */
export function isCounterType(type: MarkerType): boolean {
  return (
    type === "decimal" ||
    type === "upperAlpha" ||
    type === "lowerAlpha" ||
    type === "upperRoman" ||
    type === "lowerRoman"
  );
}

/** Compose a marker for a 0-based sibling `index`: the glyph, plus the (free-
 *  string, so ESCAPED here) delimiter when the type is a counter — a symbol/none
 *  takes no delimiter. Glyphs are machine constants; only the delim is user text. */
function composeMarker(spec: MarkerSpec, index: number): string {
  const glyph = markerGlyph(spec.type, index);
  return isCounterType(spec.type)
    ? `${glyph}${escapeHtml(spec.delim.slice(0, 20))}`
    : glyph;
}

/** Default marker TYPE for a nesting depth: 1 -> a -> i -> cycle (delimiter
 *  defaults to "dot", matching the old fused "1./a./i." cycle). */
function defaultMarkerType(depth: number): MarkerType {
  return (["decimal", "lowerAlpha", "lowerRoman"] as const)[(depth - 1) % 3];
}

/** The default marker spec for a depth (type cycle + "." delimiter). */
export function defaultMarkerSpec(depth: number): MarkerSpec {
  return { type: defaultMarkerType(depth), delim: "." };
}

/**
 * Per-table multilevel-numbering config (app-drawn STATIC numbers, not Word's).
 * - `mode: "off"` -- no markers and no numbers; body rows are plain.
 * - `mode: "custom"` -- each level shows its own chosen `markers` symbol.
 * - `mode: "multilevel"` -- the renderer prefixes each node's first line with a
 *   compounded number path (`5`, `5.1`, `5.1.1`, ...) as plain escaped body text,
 *   and suppresses that node's marker. The numbers are real text in the preview
 *   AND the Word output -- nothing becomes a Word heading.
 * - `start` -- a dotted-decimal STRING (e.g. "1", "5", "5.1") that is the exact
 *   number of the FIRST top-level item; top siblings advance its last component
 *   (Start "5.1" -> 5.1, 5.2) and deeper levels append ".<1-based>" (5.1 -> 5.1.1
 *   -> 5.1.1.1). So to nest a body under a "5.0" section heading, set Start "5.1".
 *
 * Numeric only for now; `mode` is shaped as a discriminator so letter/roman
 * styles could be added later without reshaping callers.
 *
 * `levels` (per indent level, index = level − 1, like `markers`) is a per-level
 * SHOW/HIDE of the number: the compounded path is always computed by full depth
 * (so numbers never collide), but a level whose entry is `false` renders no number
 * (its line goes plain). Sparse/short → shown. Lets you number the structural
 * levels and leave detail levels (e.g. Rationale / Notes) unnumbered.
 */
export type NumberingConfig = {
  mode: "off" | "custom" | "multilevel";
  start: string;
  levels: boolean[];
};

/** The default numbering config; shared by the renderer and TableState. */
// Mode "off" ON PURPOSE: a fresh paste carries NO preset formatting — no
// markers, no numbers — until the user picks a Markers mode themselves.
export const DEFAULT_NUMBERING: NumberingConfig = {
  mode: "off",
  start: "1",
  levels: [],
};

/**
 * Assign a multilevel number to every node on a SHOWN (numbered) level, as a clean
 * contiguous outline over ONLY the numbered levels. Returns a Map from node to its
 * display number; nodes on hidden levels are absent (they render plain).
 *
 * Hidden levels (`numbering.levels[depth-1] === false`) are TRANSPARENT to
 * numbering: a hidden node gets no number, and its children CONTINUE the numbering
 * of the nearest shown ancestor (sharing one counter and base path). So the
 * visible numbers never gap and -- critically -- never COLLIDE: naively dropping a
 * hidden segment and renumbering by local index would give two texts under two
 * different hidden parents the same "1.1"; carrying one shared counter across the
 * hidden siblings makes them 1.1, 1.2, 1.3, ... instead.
 *
 * The first SHOWN level reads the `start` string with its LAST component advanced
 * per sibling (Start "5.1" -> "5.1", "5.2"; Start "1" -> "1", "2"); each deeper
 * shown level appends `.<1-based index>` ("5.1" nests "5.1.1", then "5.1.1.1").
 * Digits + dots only, so the result is plain text with nothing to escape.
 */
function multilevelNumbers(
  nodes: PivotNode[],
  numbering: NumberingConfig,
  headingLevels: boolean[],
): Map<PivotNode, string> {
  const out = new Map<PivotNode, string>();
  const start = String(numbering.start || "1");
  // Advance the LAST dotted component of `start` by `i` (bumpLast("5.1", 1) ->
  // "5.2"; bumpLast("1", 0) -> "1"); used only at the top shown level.
  const bumpLast = (i: number): string => {
    const parts = start.split(".");
    const k = parts.length - 1;
    const last = parseInt(parts[k], 10);
    parts[k] = String((Number.isFinite(last) ? last : 0) + i);
    return parts.join(".");
  };
  const assign = (
    list: PivotNode[],
    depth: number,
    parentBase: string,
    counter: { n: number },
  ) => {
    // A level establishes a numbered BASE (its children count from 1 under it) if
    // it is shown for numbering OR mapped to a Word heading -- a heading is a
    // section boundary, so Word numbers it and the app's deeper rows must RESET
    // under each one (5.1 -> 5.1.1, 5.2 -> 5.2.1), even though the heading's own
    // number isn't drawn by the app. Only a plain hidden level (not a heading) is
    // transparent: its children continue the grandparent's sequence.
    const shown =
      headingLevels[depth - 1] === true || numbering.levels[depth - 1] !== false;
    for (const node of list) {
      if (shown) {
        const idx = counter.n++;
        // Top shown level = the Start value (last component advanced per sibling);
        // deeper levels append ".<1-based index>".
        const num =
          parentBase === "" ? bumpLast(idx) : `${parentBase}.${idx + 1}`;
        out.set(node, num);
        // A SHOWN level opens a fresh child counter scoped to its own number.
        assign(node.children, depth + 1, num, { n: 0 });
      } else {
        // Transparent level: children keep the SAME counter + base, so numbering
        // flows continuously across the hidden groups (no gap, no collision).
        assign(node.children, depth + 1, parentBase, counter);
      }
    }
  };
  assign(nodes, 1, "", { n: 0 });
  return out;
}

/**
 * A per-FIELD line-look override (this section only): when present for a
 * [level][slot] position, the field's WHOLE line (number/label/value) is wrapped
 * in an inline styled run, overriding the level's shared look for just that
 * line. Inline runs survive both Word paste modes. Values are sanitized here —
 * they come from controlled inputs, but the renderer trusts nothing.
 */
export type FieldLook = {
  font: string; // "" = inherit the level font (no font-family emitted)
  sizePt: number;
  color: string; // #rrggbb
  bold: boolean;
  italic: boolean;
  underline: boolean;
};

/** Wrap one line's content in its FieldLook run (span + b/i/u), sanitized. */
function wrapFieldLook(content: string, lk: FieldLook): string {
  const font = lk.font.replace(/[^A-Za-z0-9 _-]/g, "").trim();
  const size = Math.min(72, Math.max(1, Math.floor(lk.sizePt) || 11));
  const color = /^#[0-9a-fA-F]{6}$/.test(lk.color) ? lk.color : "#000000";
  const style =
    (font ? `font-family:'${font}';` : "") +
    `font-size:${size}pt;color:${color}`;
  let h = `<span style="${style}">${content}</span>`;
  if (lk.underline) h = `<u>${h}</u>`;
  if (lk.italic) h = `<i>${h}</i>`;
  if (lk.bold) h = `<b>${h}</b>`;
  return h;
}

/**
 * Render a pivot (nested-rows) tree as one HTML document fragment.
 *
 * The optional `title` is emitted as a distinct `<p class="ws-title">` so
 * `buildWordHtml` can map it to a Word heading style; the nested rows
 * (`<p class="ws-lvl" data-level="N">`) map to a body style. When a title is
 * present the nested data starts at level 2 beneath it, otherwise at level 1.
 *
 * Each node carries one or more `lines` (`PivotLine` fields stacked at that
 * indent level). For each line the label (`name: `) is shown/bolded/underlined
 * per `fieldLabels[col]` (default: shown, plain) and the value follows; `name`
 * and `value` are escaped separately. Only the FIRST line of a node gets the
 * level's marker (`markers[depth-1]`, falling back to the legacy cycle, counting
 * up per parent); the rest read as plain body lines. The title is never marked.
 *
 * NUMBERING (app-drawn static numbers): when `numbering.mode === "multilevel"`,
 * `multilevelNumbers` precomputes a node->number map and the FIRST line of each
 * numbered node is prefixed with it as plain escaped body text. The top shown
 * level reads the `start` string (last component advanced per sibling: Start "5.1"
 * -> "5.1", "5.2"); deeper shown levels append `.<1-based index>` ("5.1" -> "5.1.1"
 * -> "5.1.1.1"). Set Start "5.1" to nest a body under a "5.0" section heading. The number replaces (suppresses) that
 * node's per-level marker, and the leading number/marker inherits the first
 * field's bold (so bolding a category bolds its number too). The numbers are real
 * text in BOTH the preview and the
 * Word output -- nothing becomes a Word heading, so Word's Navigation pane / TOC
 * stay clean. Numbering is independent of the title (top data siblings number from
 * `start` whether or not a title exists). `numbering.levels[depth-1] === false`
 * makes a level TRANSPARENT: its line goes plain AND its children continue the
 * nearest shown level's sequence, so the visible numbers form a contiguous outline
 * of only the shown levels -- no gaps (1.0 -> 1.1, never 1.1.1 under a hidden 1.1)
 * and no collisions (see `multilevelNumbers`).
 *
 * HEADING ROWS (`headingLevels[depth-1]`): a level the user mapped to a Word
 * heading. Its FIRST line is tagged `data-heading="K"` so `buildWordHtml` maps it
 * to the destination `Heading K` style (nav pane + collapsible). K is SEQUENTIAL
 * by default: one rank under the PREVIOUS heading level — the title (counted only
 * when a title is emitted; 0 when it is not a heading), then each heading-checked
 * level in depth order, PINNED VALUES INCLUDED (clamped 9) — so checking levels 1
 * and 4 gives H2 → H3 under a Heading-1 title (never a skipped H2 → H5 outline,
 * which Word's accessibility checker flags), and pinning a level re-derives the
 * autos beneath it (pin H1 → the next auto is H2). An explicit
 * `headingRanks[depth-1]` (1–9; 0/absent = auto) OVERRIDES the rank for templates
 * whose numbering keys off a specific heading (several levels may share one
 * rank) — an explicit pin is the only way a rank gap can occur. The app's
 * number/marker is SUPPRESSED on heading rows (Word supplies the number), but the
 * number PATH still computes so deeper body rows still nest (TYPE → Word "5.1",
 * TITLE → app "5.1.1"). Usually just the top level or two.
 *
 * BLANK LINE (`breakAfter[depth-1]`): after a ticked level's OWN line(s), an
 * empty spacer paragraph (`<p ... data-level="N">&#160;</p>` -- a non-breaking
 * space, no marker/number/label) is pushed BEFORE its nested children, so both
 * the preview and the Word paste read "1. Index / blank / a. Group". A childless
 * node gets the blank before its next sibling instead (skipped on the last, so a
 * section never ends on a stray blank).
 *
 * GAP AFTER (`gapAfter[depth-1]`): the between-groups counterpart -- the same
 * spacer after a ticked level's WHOLE subtree, separating one group from the
 * next ("…8.2 x / blank / 9 Region"); between siblings only. The two options
 * compose (blank inside a group AND a gap between groups).
 *
 * The nbsp keeps Word from dropping the empty paragraph on paste; spacers are
 * emitted outside the lines loop, so they never perturb the marker/number
 * counters, and any spacers dangling at the very end of the document are trimmed.
 *
 * The nesting depth rides in a `data-level` ATTRIBUTE, not the tag name (HTML
 * only has h1-h6, and a class like `pl-3` would collide with Tailwind padding
 * utilities); `RenderedPreview` styles `[data-level="N"]`. Level is 1-based and
 * clamped at 9; deeper nodes still render at level 9.
 *
 * Only user-derived text is escaped; the `class`/`data-*` values are machine
 * constants (single digits), and the number path is digits + dots, so there is
 * no attribute-injection surface. Returns a bare fragment. Pure: input unchanged.
 */
export function renderPivotTree(
  nodes: PivotNode[],
  title?: string,
  markers: MarkerSpec[] = [],
  fieldLabels: Record<number, FieldLabel> = {},
  breakAfter: boolean[] = [],
  numbering: NumberingConfig = DEFAULT_NUMBERING,
  headingLevels: boolean[] = [],
  titleLevel = 0,
  pageBreakBefore = false,
  headingRanks: number[] = [],
  labelSeps: string[][] = [],
  gapAfter: boolean[] = [],
  fieldLooks: (FieldLook | undefined)[][] = [],
  breakBeforeSlots: boolean[][] = [],
): string {
  const numbered = numbering.mode === "multilevel";
  // Precompute each numbered node's display number (transparent hidden levels, top
  // level ".0"). Absent => this node's level is hidden or numbering is off.
  const numberOf = numbered
    ? multilevelNumbers(nodes, numbering, headingLevels)
    : null;
  // A body level mapped to a Word heading nests one under the TITLE's heading — but
  // ONLY when a title is actually emitted. `titleLevel` (from the Heading dropdown,
  // default "Heading 1") is passed even for a BLANK title, so gate the offset on
  // whether a title exists; otherwise a top body heading becomes Heading 2 with no
  // Heading 1 title above it (an orphaned/malformed Word outline).
  const bodyHeadingBase = title ? titleLevel : 0;
  // The Word heading each depth maps to, resolved once. AUTO = one rank under
  // the PREVIOUS heading level (the title, then each checked level in order —
  // pinned values included), so an auto level can never skip a rank and autos
  // re-derive around any pin; an explicit headingRanks entry (1-9) overrides.
  // Only consulted for heading-checked depths.
  const headingKByDepth: number[] = [];
  {
    let prevK = bodyHeadingBase;
    for (let d = 1; d <= 9; d++) {
      const explicit = headingRanks[d - 1];
      const k =
        typeof explicit === "number" && explicit >= 1 && explicit <= 9
          ? Math.floor(explicit)
          : Math.min(Math.max(prevK + 1, 1), 9);
      headingKByDepth[d - 1] = k;
      if (headingLevels[d - 1] === true) prevK = k;
    }
  }
  const blocks: string[] = [];
  const walk = (list: PivotNode[], level: number, depth: number) => {
    const lvl = Math.min(level, 9);
    const spec = markers[depth - 1] ?? defaultMarkerSpec(depth);
    // Heading rows: this level is mapped to a Word heading (nav + collapsible).
    // Word supplies its number, so the app's number/marker is suppressed on it; the
    // number PATH still computes (via numberOf), so deeper body rows nest correctly.
    const isHeading = headingLevels[depth - 1] === true;
    const headingK = headingKByDepth[Math.min(depth, 9) - 1];
    list.forEach((node, i) => {
      // Markers render only in "custom" mode (each level's chosen symbol); a
      // multilevel node shows its precomputed number instead, and "off" shows
      // neither.
      const m = numbering.mode === "custom" ? composeMarker(spec, i) : "";
      const num = numberOf?.get(node) ?? "";
      // LINE BREAK BEFORE (`breakAfter[depth-1]` — the field name is historical):
      // a spacer immediately BEFORE this node's line(s). Uniform positional
      // semantics — every level's "before" means "blank above my line"; a lead
      // spacer at the very start of the document is trimmed by the post-pass.
      if (breakAfter[depth - 1])
        blocks.push(`<p class="ws-lvl" data-level="${lvl}">&#160;</p>`);
      // Hanging alignment for stacked nodes: when line 0 carries a LEAD (number
      // or marker) and the node stacks more lines, line 0 hangs its lead into
      // the margin (data-hang) and the continuation lines indent to match
      // (data-cont) — so "Product" lines up with "Country", not with "1.1".
      const lead0 =
        !isHeading && (num || (!numbered && m)) ? true : false;
      const hangs = lead0 && node.lines.length > 1;
      node.lines.forEach((line, j) => {
        // BREAK BETWEEN STACKED FIELDS: a per-field flag puts a blank line
        // before THIS stacked line, inside the group ("1.1 Country / blank /
        // Product"). Slot 0 is covered by the level's own "before" above.
        if (j > 0 && breakBeforeSlots[depth - 1]?.[j] && !isHeading)
          blocks.push(`<p class="ws-lvl" data-level="${lvl}">&#160;</p>`);
        const lf = fieldLabels[line.col] ?? DEFAULT_FIELD_LABEL;
        const showLabel = lf.show && line.name !== "";
        // Per-(level, stack-slot) label separator (free string, ≤20 chars,
        // escaped inside wrapLabel); absent → the classic ": ".
        const label = showLabel
          ? wrapLabel(
              line.name,
              lf,
              (labelSeps[depth - 1]?.[j] ?? ": ").slice(0, 20),
            )
          : "";
        const value = escapeHtml(line.value);
        // First line only: the level's number (when shown) else its marker (only
        // when numbering is off). A heading row shows neither (Word numbers it);
        // a numbered-but-hidden level renders plain.
        const lead =
          j === 0 && !isHeading
            ? num
              ? `${escapeHtml(num)} `
              : !numbered && m
                ? `${m} `
                : ""
            : "";
        // The leading number/marker inherits the field's bold, so bolding a
        // category also bolds its number (one `<b>` run, survives a Word paste).
        const prefix =
          lead && showLabel && lf.bold ? `<b>${lead}</b>` : lead;
        // A heading row's FIRST line carries data-heading so buildWordHtml maps it
        // to the destination "Heading K" style.
        const headingAttr =
          j === 0 && isHeading ? ` data-heading="${headingK}"` : "";
        // The FIRST line of each top-level group AFTER the first carries data-break
        // when "page break before" is on → buildWordHtml turns it into a Word
        // page-break-before. (depth 1 = the top-level groups, regardless of a title.)
        const breakAttr =
          j === 0 && pageBreakBefore && depth === 1 && i > 0
            ? ` data-break="1"`
            : "";
        const alignAttr = hangs
          ? j === 0
            ? ` data-hang="1"`
            : ` data-cont="1"`
          : "";
        // Per-FIELD line-look override: the whole line (lead included, so the
        // number matches its line) wraps in an inline styled run.
        const lk = fieldLooks[depth - 1]?.[j];
        const body = `${prefix}${label}${value}`;
        blocks.push(
          `<p class="ws-lvl" data-level="${lvl}"${breakAttr}${headingAttr}${alignAttr}>${
            lk && !isHeading ? wrapFieldLook(body, lk) : body
          }</p>`,
        );
      });
      if (node.children.length > 0) walk(node.children, level + 1, depth + 1);
      // LINE BREAK AFTER (`gapAfter[depth-1]`): a spacer after this node's
      // WHOLE subtree ("…8.2 x / blank / 9 Region"). Emitted unconditionally —
      // the old skip-the-last-sibling guard silently no-opped on one-group-per-
      // parent levels (a sole child is always "last"); the post-pass instead
      // collapses doubled spacers and trims the document's edges.
      if (gapAfter[depth - 1])
        blocks.push(`<p class="ws-lvl" data-level="${lvl}">&#160;</p>`);
    });
  };
  if (title) {
    // Title is a distinct paragraph (ws-title) so buildWordHtml can map it to a
    // heading style; the data nests beneath it starting at level 2.
    blocks.push(`<p class="ws-title">${escapeHtml(title)}</p>`);
    walk(nodes, 2, 1);
  } else {
    walk(nodes, 1, 1);
  }
  // Spacer post-pass. Matched exactly on the spacer shape emitted above, so a
  // real data row containing "&#160;" (escaped on input) can never be touched.
  // 1) COLLAPSE runs of consecutive spacers into one — overlapping choices
  //    (this group's "after" meeting the next group's "before") read as one
  //    blank line, never a growing stack. 2) TRIM the document's edges: no
  //    leading blank (a "before" on the very first block) and no trailing one.
  const isSpacer = (b: string) =>
    /^<p class="ws-lvl" data-level="[1-9]">&#160;<\/p>$/.test(b);
  const out: string[] = [];
  for (const b of blocks) {
    if (isSpacer(b) && (out.length === 0 || isSpacer(out[out.length - 1])))
      continue;
    out.push(b);
  }
  while (out.length > 0 && isSpacer(out[out.length - 1])) out.pop();
  return out.join("\n");
}

# Roadmap

The app is **pivot-only**: paste → nest → render → Copy for Word. It began with four view modes and a Download option; those were removed to focus on the pivot. Status of the current feature set below.

| Step | What | Status |
| ---- | ---- | ------ |
| **Data model** | `PivotNode { lines: PivotLine[], children }`, `PivotLine { col, name, value }`, `FieldLabel` ([`lib/types.ts`](../lib/types.ts)) | Done |
| **Paste → raw rows** | SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts)) | Done |
| **Pivot mapper** | `rowsToPivotTree(rows, levels, sortDirs?)` — ordered indent buckets; each bucket is one level of ≥1 field merged by a composite key; a post-pass reorders sibling groups per `sortDirs` ([`lib/mapper.ts`](../lib/mapper.ts)) | Done |
| **Structure picker** | Add-fields pool + placed list with ◄ outdent / ► indent (stack = same indent) / ▲ ▼ reorder / ✕ remove (`pivotLevels: number[][]`) | Done |
| **Per-field label format** | Per-field toggles (`fieldLabels`): show/hide the `Field name:` label, bold it, underline it (label-only emphasis) | Done |
| **Per-column sorting** | Per-field ↕/↑/↓ cycle (`sortDirs: Record<col, "asc"|"desc">`) orders the sibling groups at that field's level; numeric + text aware, case-insensitive (`localeCompare`, `numeric:true`); off keeps first-seen order | Done |
| **Markers** | Per-level `<select>` (`1./1)/A./a./I./i./•/–/None`); first field of a multi-field level marked, restarting per parent | Done |
| **Blank line after group** | Per-level checkboxes (`breakAfter: boolean[]`); when on, a nbsp spacer paragraph follows that level's whole subtree (survives a Word paste; no marker/number/label) | Done |
| **Title heading** | Optional Section title → the chosen Word `Heading 1–4` (`mso-style-name`, "5.0") at the correct `mso-outline-level` on a Use-Destination-Styles paste (level threaded as `titleLevel` via `headingLevel(name)`); the user's full title Look (font/size/color/bold/italic/underline) applies inline on top, so it pastes exactly as previewed — Word supplies only the number/outline (and any all-caps) | Done |
| **Multilevel numbering (app-drawn static numbers)** | Per-table `numbering: { mode, start, levels }` prefixes each node's first line with a static number from a chosen `start` (a dotted-decimal string = the exact first item number; `5.1` → `5.1.1` → `5.1.1.1`) as plain body text and suppresses that node's marker; a per-level checkbox (`numbering.levels`) hides a level and makes it transparent, so the numbers follow the gap (`5.1` → `5.1.1`, `5.1.2`) with no gaps or collisions. NOT Word auto-numbering — nothing becomes a Word heading, so the Nav pane / TOC stay clean and the preview shows the real numbers | Done |
| **Word heading rows** | `headingLevels: boolean[]` (per level, labelled by field name) maps a level's rows to a destination `Heading K` style — nav pane + collapsible, flush-left; Word numbers them so the app number/marker is suppressed there (the path still computes, so deeper rows nest). Scoped to the top level or two, so the nav stays clean | Done |
| **Renderer** | `renderPivotTree(nodes, title?, markers?, fieldLabels?, breakAfter?, numbering?, headingLevels?, titleLevel?)` → HTML fragment ([`lib/renderers.ts`](../lib/renderers.ts)) | Done |
| **Clipboard output** | `text/html` (+ `text/plain`) for Word ([`lib/clipboard.ts`](../lib/clipboard.ts) `buildWordHtml`) | Done |
| **Multiple tables + Sections rail** | Paste appends a table; rows in a left Sections rail within a 4-pane IDE layout (cap 100) | Done |
| **Shared per-level styling** | One "Heading levels" panel → `HeadingStyle = { levels: LevelStyle[] }`: Level 1 = title, Levels 2-9 = nested rows (color/font/size/bold each, default all the same; "Reset levels") + one Body font | Done |
| **Combined export** | **Copy all** — every table as one Word doc | Done |
| **Malformed-paste handling** | Graceful empty/parse-error states | Done |

## Removed (was built, then cut to focus on pivot)

- The other three layouts: **Grouped by field**, **Fields as bullets**, **A/B/C/D sections** (with `lib/numbering.ts`, the `Section`/`Subsection`/`Body` model, `rowsToTree`/`rowsToAttributeSections`/`rowsToGroupedSections`, `renderTree`/`renderBody`).
- **Download for Word** / **Download all** (`.doc` export) — clipboard-only now.

## Possible next steps

- **Pivot aggregation** — counts/sums per group (e.g. `Brazil (3)` or a totals summary).
- **Per-table heading styling** — currently styling is global; per-table would need per-instance class scoping or inline styles.
- **`.docx` export** (currently out of scope).

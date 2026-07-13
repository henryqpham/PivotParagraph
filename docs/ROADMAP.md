# Roadmap

The app is **pivot-only**: paste → nest → render → **Copy all**. It began with four view modes and a Download option; those were removed to focus on the pivot. Status of the current feature set below.

| Step | What | Status |
| ---- | ---- | ------ |
| **Data model** | `PivotNode { lines: PivotLine[], children }`, `PivotLine { col, name, value }`, `FieldLabel` ([`lib/types.ts`](../lib/types.ts)) | Done |
| **Paste → raw rows** | SheetJS clipboard parse + JSON preview ([`lib/parser.ts`](../lib/parser.ts)) | Done |
| **Pivot mapper** | `rowsToPivotTree(rows, levels, sortDirs?)` — ordered indent buckets; each bucket is one level of ≥1 field merged by a composite key; a post-pass reorders sibling groups per `sortDirs` ([`lib/mapper.ts`](../lib/mapper.ts)) | Done |
| **Structure picker** | Add-fields pool + placed list (Windows-Explorer-style tree connector lines │ ├ └ to show nesting) with ◄ outdent / ► indent (stack = same indent) / ▲ ▼ reorder / ✕ remove — buttons only (`pivotLevels: number[][]`) | Done |
| **Per-field label format** | Per-field toggles (`fieldLabels`): show/hide the `Field name:` label, bold it, underline it (label-only emphasis) | Done |
| **Per-column sorting** | Per-field ↕/↑/↓ cycle (`sortDirs: Record<col, "asc"|"desc">`) orders the sibling groups at that field's level; numeric + text aware, case-insensitive (`localeCompare`, `numeric:true`); off keeps first-seen order | Done |
| **Markers** | A 3-mode dropdown (**Off** = none / **Custom** = per-level `<select>` `1./1)/A./a./I./i./•/–/None` / **Multilevel numbers**; default Custom); in Custom the first field of a multi-field level is marked, restarting per parent | Done |
| **Blank line between top-level groups** | A single on/off checkbox beside the Markers control (`breakAfter` = `[true]`/`[]`, no longer per-level); when on, a nbsp spacer paragraph follows each top-level group's whole subtree (survives a Word paste; no marker/number/label) | Done |
| **Title heading** | Optional Section title → the chosen Word `Heading 1–4` (`mso-style-name`, "5.0") at the correct `mso-outline-level` on a Use-Destination-Styles paste (level threaded as `titleLevel` via `headingLevel(name)`); the user's full title Look (font/size/color/bold/italic/underline) applies inline on top, so it pastes exactly as previewed — Word supplies only the number/outline (and any all-caps) | Done |
| **Multilevel numbering (app-drawn static numbers)** | Per-table `numbering: { mode, start, levels }` (the Markers control's **Multilevel** mode; `mode: "off"|"custom"|"multilevel"`) prefixes each node's first line with a static number from a chosen `start` (a dotted-decimal string = the exact first item number; `5.1` → `5.1.1` → `5.1.1.1`) as plain body text and suppresses that node's marker; a per-level checkbox (`numbering.levels`) hides a level and makes it transparent, so the numbers follow the gap (`5.1` → `5.1.1`, `5.1.2`) with no gaps or collisions. NOT Word auto-numbering — nothing becomes a Word heading, so the Nav pane / TOC stay clean and the preview shows the real numbers | Done |
| **Word heading rows** | `headingLevels: boolean[]` (per level, labelled by field name) maps a level's rows to a destination `Heading K` style — nav pane + collapsible, flush-left; Word numbers them so the app number/marker is suppressed there (the path still computes, so deeper rows nest). Scoped to the top level or two, so the nav stays clean | Done |
| **Renderer** | `renderPivotTree(nodes, title?, markers?, fieldLabels?, breakAfter?, numbering?, headingLevels?, titleLevel?)` → HTML fragment ([`lib/renderers.ts`](../lib/renderers.ts)) | Done |
| **Clipboard output** | `text/html` (+ `text/plain`) for Word ([`lib/clipboard.ts`](../lib/clipboard.ts) `buildWordHtml`) | Done |
| **Multiple tables + Sections rail** | Paste appends a table; rows in a left Sections rail within a 4-pane IDE layout (cap 100) | Done |
| **Shared per-level styling** | The per-depth nested-row look (color swatch + font/size/bold popover) is edited inline in the per-level matrix's **Look** column (global by depth, badged *All tables*), while only the document-wide **Body font**, **Indent/level**, and **Reset levels** sit in a compact **⚙ Document** popover in the top command band; a shared dependency-free `Popover` (`components/Popover.tsx`) backs both the Look and Document popovers. `PasteInput` owns the state, passes a slimmed `appearance` prop (`{ levelStyles, onLevelChange }`) → `HeadingStyle = { levels: LevelStyle[] }` (per-depth look, color/font/size/bold each, default all the same) + one Body font + indent step, and drives body font / indent / reset itself; `LevelInput`/`DEFAULT_LEVEL` in `tableModel.ts`, `FONTS` exported from `TableCard` | Done |
| **Combined export** | **Copy all** — every table as one Word doc | Done |
| **Malformed-paste handling** | Graceful empty/parse-error states | Done |
| **Session persistence** | The whole workspace (every `TableState` + shared styling) auto-saves to `localStorage` (versioned key, [`lib/persistence.ts`](../lib/persistence.ts)) — debounced + flushed on tab-hide, rehydrated on mount (id counter re-seeded). Local-only; nothing uploaded | Done |
| **Combined “All sections” preview** | The Preview tab’s **This section / All sections** toggle renders every section stacked on its own paper, from the same `tableToHtml` fragments `Copy all` concatenates (so preview and export can’t diverge); `RenderedPreview` takes `sections?: string[]` | Done |
| **Copy-all confirmation** | After `Copy all`, a visible + `aria-live` banner reports the real contributing-section count and the decisive Word paste step (**Keep Source Formatting** vs **Use Destination Styles**, the latter only when a heading is mapped); friendly failure + “add fields first” states | Done |
| **Reorder sections** | Per-row **▲▼** in the Sections rail ([`components/SectionsRail.tsx`](../components/SectionsRail.tsx)) swap neighbours in `tables[]`; rail order = the order `Copy all` stacks them | Done |
| **Undo / safe delete** | **Remove section** and **Clear all** stash a snapshot and offer a transient **Undo** (focus-grabbing); **Clear all** is a two-step inline confirm | Done |
| **First-run activation** | Onboarding empty state (value prop + 1-2-3 steps + **Try an example** → `makeExampleTable`) and **paste-anywhere** (a window `paste` listener ingests a copied table unless a text field is focused, so the paste zone no longer needs focus) | Done |

## Removed (was built, then cut to focus on pivot)

- The other three layouts: **Grouped by field**, **Fields as bullets**, **A/B/C/D sections** (with `lib/numbering.ts`, the `Section`/`Subsection`/`Body` model, `rowsToTree`/`rowsToAttributeSections`/`rowsToGroupedSections`, `renderTree`/`renderBody`).
- **Download for Word** / **Download all** (`.doc` export) — clipboard-only now.

## Possible next steps

- **Drag-and-drop** — native HTML5 DnD for arranging fields in the Structure list and reordering sections in the rail (the ▲▼ buttons ship now and stay as the keyboard path); deferred as its own round for the drop-between-vs-into fiddliness + a11y parity.
- **Duplicate a section** — clone a configured `TableState` for same-shape tables (one per region/quarter).
- **Pivot aggregation** — counts/sums per group (e.g. `Brazil (3)` or a totals summary).
- **Per-table heading styling** — currently styling is global; per-table would need per-instance class scoping or inline styles.
- **`.docx` export** (currently out of scope).

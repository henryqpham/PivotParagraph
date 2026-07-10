# Architecture

## Core data model

The product is a recursive **pivot tree**, not a flat table (see [`lib/types.ts`](../lib/types.ts)):

```ts
PivotNode {
  lines: PivotLine[]      // the fields at one indent level (>=1, stacked)
  children: PivotNode[]   // leaf = []
}
PivotLine { col, name, value }   // label `name` (header) + `value`, kept SEPARATE + raw
```

The parser produces a raw **Grid** (`Cell[][]`); `rowsToPivotTree` turns it into a `PivotNode` tree of arbitrary depth. The pivot structure is an ordered list of **indent buckets** — each bucket is one indent level holding one or more fields. A node's `lines` are that level's fields stacked at the same indent (each keeps its label `name` and `value` separate so the renderer can show/hide/format the label); rows whose values match across all of a level's fields merge into one node (composite group key). Only the title maps to a Word heading; everything else — including any app-drawn multilevel numbers — is styled body text.

Multiple pasted tables are held as a `TableState[]` in `components/PasteInput.tsx`; each `TableState` carries its own grid, `pivotLevels` (`number[][]` — the ordered indent buckets), `markers`, `fieldLabels` (`Record<col, FieldLabel>` — per-field label show/bold/italic/underline), `sortDirs` (`Record<col, "asc"|"desc">` — per-field sort of sibling groups), `breakAfter` (`boolean[]` — per-level "blank line after"), `numbering` (`{ mode, start, levels }` — app-drawn static multilevel numbers, `levels` showing/hiding the number per indent level), `headingLevels` (`boolean[]` — levels mapped to a Word heading: nav pane + collapsible, Word-numbered), and `sectionTitle`. `components/tableModel.ts` `tableToHtml(t, titleLevel)` runs the per-table nest→render pipeline (threading `sortDirs` into the mapper, `breakAfter`/`numbering`/`headingLevels`/`titleLevel` into the renderer) and is the single source used by both the right-pane preview and the combined export, so both stay identical.

## The pivot view (how a Grid becomes a tree) — `rowsToPivotTree`

Row 0 is field names. You arrange fields into an **ordered list of indent buckets** (`pivotLevels: number[][]`): bucket 1 is the outermost level, within each group nest by bucket 2, and so on. A bucket can hold **one or more** fields — fields stacked in one bucket render at the same indentation and form a **composite group key**, so rows merge into one node only when they match across every field in that bucket. Each line is labelled `Field name: value`; per field, the label (`Field name: `) can be hidden, bolded, or underlined (`fieldLabels`), leaving just the value. A per-field **sort** (`sortDirs`) reorders the sibling groups at that field's level (ascending/descending, numeric + text aware via `localeCompare` with `numeric:true`); off keeps first-seen order. A per-level **Markers** picker chooses each level's bullet/number style; only the first field of a multi-field level is marked. A per-table **Numbering** mode instead prefixes each node's first line with an app-drawn static number from a chosen `start` (a dotted-decimal string = the exact first item number; `5.1` → `5.1.1` → `5.1.1.1`) and suppresses that node's marker; a per-level checkbox (`numbering.levels`) hides a level's number, making it TRANSPARENT — its children continue the nearest shown level's sequence, so the numbers follow the gap (`5.1` → `5.1.1`, `5.1.2`, never a number under a hidden parent) with no collisions (`multilevelNumbers`). A **blank line after each** dropdown (`breakAfter`, one chosen level labelled by its field name) pushes a spacer paragraph after each of that level's groups' whole subtree. Blank cell → `(blank)`; first-seen order preserved at every level (before sorting).

```
1. Item Category: Fruit
   a. Item Name: Apple
        i. Item Qty: 16             ← bucket 3 holds three fields,
           UID: AVASCASC               stacked at one indent
           Item Description: Lorem…
   b. Item Name: Banana
        …
2. Item Category: Meat
   …
```

Output is a `PivotNode[]` (arbitrary depth; each node carries its level's `lines`). The optional **Section title** renders as `<p class="ws-title">` and the nested rows as `<p class="ws-lvl" data-level="N">` (depth clamped at 9); a "blank line after" spacer is an extra `<p class="ws-lvl" data-level="N">&#160;</p>` after a node's subtree (an nbsp so Word keeps it; no marker/number/label). When numbering is on, a node's first line is prefixed with its compounded number as plain escaped body text (the number replaces the marker). A level the user marks as a **Word heading** tags its first line `data-heading="K"`. `buildWordHtml` maps the **title** (and any heading-marked levels) to the destination `Heading 1`/`Heading K` styles when set (a Use-Destination-Styles paste adopts them — `5.0`, `5.1` — so the heading rows enter the nav pane and become collapsible); every other body row (spacers, numbered/markered lines) uses the app's direct per-level styling, so the Word output matches the live preview. With a title the nested data starts at level 2; without one, level 1. The Structure picker builds the buckets (Add-fields pool plus ◄ outdent / ► indent / ▲ ▼ reorder / ✕ remove), the per-field label toggles (Aa show/hide, B bold, I italic, U underline), and the per-field sort (↕/↑/↓).

## Data flow

```
clipboard (text/html, else text/plain)
  -> SheetJS XLSX.read({ type: "string" })
  -> sheet_to_json({ header: 1, blankrows: false, defval: "", raw: false })   -> raw Grid  (append a TableState)
  -> tableToHtml(t, titleLevel):
       rowsToPivotTree(grid, pivotLevels, sortDirs)   (build first-seen, then sort-post-pass)
         -> renderPivotTree(tree, title?, markers, fieldLabels, breakAfter, numbering, headingLevels, titleLevel)  -> HTML fragment
  -> live preview (RenderedPreview, dangerouslySetInnerHTML; scoped [data-level] CSS)
  -> buildWordHtml + htmlToPlainText -> navigator.clipboard.write             -> paste into Word
```

Per-table **Copy for Word** runs `tableToHtml` for that one table; combined **Copy all** joins every table's fragment and runs **one** `buildWordHtml` (valid because its rewrites are global regexes and it emits a single `@page`). `renderPivotTree` escapes all user-derived text (`& < >`). The JSON view shows the raw Grid instead of the rendered tree.

## Clipboard output

`lib/clipboard.ts` wraps the rendered fragment for Word and applies the styling:
- `buildWordHtml(fragment, heading, bodyFont)` → an Office-namespaced `<html>` with a `<style>` (`@page`, body font, the mapped-title rule) and `<body>{rewritten fragment}`. **The title:** when `headingStyleName` is set it becomes `<p class="MsoTitle">` + a `mso-style-name:"<name>";mso-outline-level:<N>` rule (N = the chosen Heading N via `headingLevel(name)`, so Heading 2/3/4 land at the right outline depth on a Use-Destination-Styles paste), and `wrapTitleOnHeading` layers the user's FULL `titleStyle` Look on top as inline direct formatting — one `<span style="font-family;font-size;color;font-weight">` (the explicit `font-weight` sets bold in BOTH directions, so it can also un-bold an intrinsically-bold heading) wrapped by `<i>`/`<u>` runs — so the mapped title's appearance pastes EXACTLY as previewed. Word supplies ONLY the heading's auto-NUMBER, its outline/nav-pane entry, its paragraph spacing, and any ALL-CAPS font effect (the number and all-caps are the two things inline HTML cannot override while mapped). Blank → the app's own `titleStyle` look inline in full. **Heading-marked rows** (`data-heading="K"`) become `<p class="MsoHeadingK">` + a `mso-style-name:"Heading K";mso-outline-level:K` rule, so they enter the destination outline (nav pane + collapsible) and Word numbers them. **Every other body row** (including spacer paragraphs and any app-drawn multilevel numbers, which are already plain text in the fragment): a `<p>` with **inline** direct formatting — `heading.levels[N-1]` (color/font/size) + left-indent + compact spacing (`margin:0`, `line-height:115%`) — with `<b>`/`<i>`/`<u>` runs for the label per `fieldLabels` (underline wraps just the name, not the `: ` separator). Inline (not a CSS class) is the key: a Use-Destination-Styles paste discards class/style-name formatting it can't map (resetting unknown classes to Normal, dropping bold and bringing Normal's space-after) but keeps inline direct formatting + inline runs, so the per-level look, tight spacing, and label bold/italic/underline survive and match the live preview. The browser writes the Windows CF_HTML header automatically.
- `htmlToPlainText(fragment)` → readable plain-text fallback for the `text/plain` flavor.

`HeadingStyle = { levels: LevelStyle[]; indentStep: number; headingStyleName: string; titleStyle: TitleStyle }` is the single styling source, built once in `PasteInput` and shared by every table — `headingStyleName` is the Word style name for the title (default `Heading 1`, options `None`/`Heading 1–4`; blank/`None` = the app's own `titleStyle` look inline). `PasteInput` turns that name into a number via `headingLevel(headingStyleName)` (1–9, or 0 for None) and threads it as `titleLevel`, so a mapped title lands at its chosen `mso-outline-level` (Heading 2/3/4 give a correct outline, not a hardcoded 1) and its full `titleStyle` Look (font/size/color/bold/italic/underline) applies inline on top, while body-heading rows nest at `titleLevel + depth`. `levels` is the direct per-level look (Arial 11 black) used for body rows, `titleStyle` is the shared title look, and `indentStep` (inches) is the left-indent per nesting level. The card's `copyForWord()` and the parent's `copyAll()` both go through `tableToHtml(t)` (so the preview and both exports stay identical) and write a `ClipboardItem` with both flavors via `navigator.clipboard.write`. Export is clipboard-only (no download).

## Out of scope

- **`.docx` generation** — export is HTML-on-clipboard only.

## Pipeline diagram

```mermaid
flowchart TD
    A["User pastes Excel block"]
    B["parseClipboard: read text/html or text/plain"]
    C["SheetJS XLSX.read -> sheet_to_json -> raw Grid"]
    PV["rowsToPivotTree(pivotLevels, sortDirs) -> PivotNode[]"]
    RP["renderPivotTree(tree, title?, markers, fieldLabels, breakAfter, numbering, headingLevels, titleLevel) -> HTML fragment"]
    P["RenderedPreview (live)"]
    J["JsonPreview (raw Grid)"]
    W["buildWordHtml + htmlToPlainText"]
    X["navigator.clipboard.write -> paste into Word"]
    A --> B --> C --> PV --> RP --> P
    C -.-> J
    RP --> W --> X
```

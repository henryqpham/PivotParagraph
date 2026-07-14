# Excel &rarr; Word Pivot

Convert pasted Excel data into a **Word-ready nested outline**. Paste one or more tables copied from Excel or Google Sheets; the app restructures each into an Excel-pivot-style hierarchy and copies it to the clipboard so it pastes into Microsoft Word.

## Why

Wide Excel tables don't fit an 8.5" x 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a narrow, nested outline: you arrange fields into ordered **indent levels** (each level can stack several fields at the same indentation) and rows nest and merge by those levels, so it flows down the page.

## Status

**Working end to end:** paste → parse (SheetJS) → nest → render → live preview → **Copy section**. Paste several tables (managed in a left Sections rail, **reorderable by drag-and-drop**), configure each table's pivot, style the levels once for all tables, preview the active section, and copy it as its own standalone Word document — with an inline confirmation that reminds you which Word paste option to choose. Your whole workspace **auto-saves in the browser** (local `localStorage`, nothing uploaded), so a refresh won't lose it, and **Remove section / Clear all are undoable**. No spreadsheet handy? The empty state has a **Try an example** button. See the [roadmap](./docs/ROADMAP.md) for what's out of scope (`.docx`).

## The pivot view

Add fields from the **Add fields** pool, then shape the **Structure**, an indented list drawn with Windows-Explorer-style **tree connector lines** (│ ├ └) to show the nesting: each field sits at an indent level, and ◄/► move it shallower/deeper (depth is capped at **9 levels** — the most Word shows — with a live N/9 chip; a field added past the cap stacks into the deepest level). Stack several fields at one level to show them together at the same indent; ▲/▼ reorder and ✕ removes — the ◄ ► ▲ ▼ ✕ buttons do all the reordering and indenting. Rows nest by the levels and **merge** when their values match across a level's fields. Each line reads `Field name: value` — per field, toggle the `Field name:` label off (just the value), or **bold**/**italic**/**underline** it (`Aa`/`B`/`I`/`U`; underline covers just the name) to make a long list scannable, and **sort** the groups at that field (↕ off / ↑ ascending / ↓ descending; numeric *and* text aware, so `2` sorts before `10`). A per-level **matrix** (one row per indent level) co-locates each level's controls: a Marker/Number cell (a **Marker** select prefixing `1.`/`a.`/`i.`/etc. in **Custom** mode, or a **Show number** checkbox in **Multilevel** mode, hidden in **Off** mode), a **Word heading** checkbox, and a **Look** cell — a color swatch plus a font/size/bold popover — that sets that depth's shared body-row look. Two checkboxes beside the Markers control shape the page: **Blank line between top-level groups** adds a blank line after each top-level group, and **New page per group** starts each top-level group (after the first) on a new Word page (a `page-break-before` that carries into the paste; the preview shows a dashed rule at each break). A per-table **Markers** control (a 3-mode dropdown: **Off (none)** = no markers or numbers / **Custom (per level)** = each level's chosen symbol / **Multilevel numbers**; default Custom) picks the body's marker style; Multilevel draws static **multilevel numbers** from a chosen **Start** — the exact first number (e.g. `5.1` → `5.1`, `5.1.1`, `5.1.1.1`), so set it to `5.1` to nest a body under a `5.0` section title — as plain body text, not Word auto-numbering, so nothing becomes a Word heading and the preview shows the real numbers; when on, the number replaces that level's marker, and the matrix's **Show number** checkbox hides a level's number. Hidden levels are *transparent*: the numbers follow the gap (e.g. number the type + text but hide the title → `5.1` then `5.1.1`, `5.1.2`, never a number under a missing parent). The matrix's **Word heading** checkbox maps a level's rows to a real Word heading (so they show in the Navigation pane and collapse, flush-left) — handy on just the top level so each section is in the document outline; a small **H1/H2… chip** beside the checkbox shows exactly which Heading the level pastes as (shifting under a mapped Section title), and since Word numbers those rows, the level's Marker/Number cell shows a *Word numbers* note instead of a dead control. An optional **Section title** lives in its own **Section Header** group — a **Heading** dropdown (None / Heading 1–4) plus its own look (font / size / **B** / **I** / **U** / color); pick a Heading (1–4, each mapped at its correct outline level) and the title maps to that Word heading on a *Use Destination Styles* (or default Ctrl+V) paste while your **full title look** — font, size, color, **B**/**I**/**U** — carries through inline, so the title pastes exactly as shown in the preview; Word adds only the heading's auto-number and outline/Navigation-pane entry (plus any all-caps effect). A **Table** tab beside the live preview shows the raw pasted grid as an actual table (exactly what came from Excel, before the outline) — with a **Header row** picker to skip banner/title rows sitting above the real header (so a "Q3 Report" banner doesn't become the field names) — and a **JSON** tab inspects the same grid as raw JSON. Fully blank spacer columns are dropped automatically on paste, so they never clutter the field pool.

## Multiple tables

Paste table after table — each becomes a row in the left **Sections rail** (one edited at a time, cap 100), inside a 4-pane IDE layout: top global controls · left rail · center builder · pinned right live preview. Reorder sections by **dragging** them in the rail (a grip ⋮ + keyboard ↑/↓ as the accessible path), and **duplicate** one with the **⧉** button (a deep clone right after the original) when you have several near-identical tables — each section is its own standalone Word document, not a stacked section; the Preview tab always shows the section you're editing. Paste a table whose **headers match an already-arranged section** and a banner offers to **reuse that section's arrangement** (levels, markers, headings, labels, sort) in one click — undoable, like everything else. The shared body styling is reached right in the per-level matrix's **Look** column (a color swatch + a font / size / bold popover per depth), with only the document-wide **Body font**, **Indent/level**, and **Reset levels** in a compact **⚙ Document** popover in the top command band — all badged *All tables* so they apply to every section; each table keeps its own fields and title. Export the active section with **Copy section**, then follow the inline paste-guidance banner in Word. The whole workspace is saved in your browser (locally, nothing uploaded), so closing the tab or refreshing won't lose it.

## Stack

- Next.js 16 (App Router) + TypeScript
- SheetJS (`xlsx`, official CDN tarball) for parsing
- Tailwind CSS v4 — Microsoft **Fluent/Office** reskin (Teams-blue accent, Segoe UI)
- Client-side only — no backend, no database, nothing uploaded. Your workspace is saved to the browser's local storage (on your machine only) so it survives a refresh.

## Getting started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), click the paste zone, and press **Ctrl/Cmd + V** with a cell range copied from Excel or Google Sheets. Add fields and set their indent levels; the live preview updates. Click **Copy section** (or press **Ctrl/Cmd + Enter** from anywhere) and paste into a Word document (*Use Destination Styles* — or a plain default Ctrl+V — maps the title and any heading-marked levels to your document's Heading styles; *Keep Source Formatting* keeps the preview look instead, and *Merge Formatting* strips the heading mapping).

## Run with Docker

The app ships a `Dockerfile` (Next.js standalone build, ~305 MB image) and a `docker-compose.yml`. With [Docker](https://www.docker.com/products/docker-desktop/) installed:

```bash
docker compose up --build
```

Then open [http://localhost:3000](http://localhost:3000). Press **Ctrl/Cmd + C** in the terminal (or `docker compose down`) to stop. Everything runs client-side in your browser — nothing is uploaded (the workspace is saved only in your browser's local storage).

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack, status
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, pivot, pipeline diagram
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order & status

## Project layout

```
app/                 App Router pages (home renders the full-height app shell)
components/
  PasteInput.tsx     app shell (4-pane IDE): paste/append (+ dropEmptyColumns) + paste-anywhere, tables[] + shared styles, localStorage persistence, full undo/redo (Ctrl+Z/Y), duplicate section, Copy-section confirmation, pinned preview (Preview/Table/JSON tabs + rows/levels/~pages stat) + Copy section
  SectionsRail.tsx   left Sections rail: select / drag-reorder (grip ⋮ + keyboard ↑/↓) / duplicate (⧉) / remove per section
  TableCard.tsx      one table's center builder: Section Header group (title + heading dropdown + title look) + Levels group (structure + N/9-levels cap + per-level controls + blank-line/new-page-per-group toggles)
  tableModel.ts      TableState + tableToHtml (per-table nest->render) + bucket helpers (add/remove/indent/outdent/move/unusedColumns) + MAX_LEVELS=9 depth cap + dropEmptyColumns/bodyGrid (ingest cleanup + header-row offset) + estimateSectionStats + newTable factory + makeExampleTable/EXAMPLE_GRID
  RenderedPreview.tsx renders the pivot HTML (live preview; ws-title + [data-level] CSS, 9 level rules; dashed [data-break] page-break rule); active-section paper
  GridTable.tsx      Table view: the raw pasted grid as an HTML table (as-is, before the outline) + a Header-row picker to skip banner rows
  JsonPreview.tsx    shows the raw parsed Grid as JSON
  Popover.tsx        shared dependency-free popover (Document + Look popovers)
lib/
  persistence.ts     versioned localStorage save/load/clear for the workspace (local-only)
  types.ts           PivotNode/PivotLine + FieldLabel + raw Grid
  parser.ts          SheetJS clipboard -> Grid
  mapper.ts          rowsToPivotTree (Grid + indent buckets + sortDirs -> PivotNode[]; sort post-pass) + cellToString
  renderers.ts       renderPivotTree (tree -> HTML fragment; labels + markers + breakAfter spacers + static multilevel numbering + page-break marks) + marker/numbering helpers
  clipboard.ts       Word-friendly clipboard wrapper (buildWordHtml / htmlToPlainText; HeadingStyle / LevelStyle)
docs/                OVERVIEW, ARCHITECTURE, ROADMAP
```

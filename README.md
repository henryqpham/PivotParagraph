# Excel &rarr; Word Pivot

Convert pasted Excel data into a **Word-ready nested outline**. Paste one or more tables copied from Excel or Google Sheets; the app restructures each into an Excel-pivot-style hierarchy and copies it to the clipboard so it pastes into Microsoft Word.

## Why

Wide Excel tables don't fit an 8.5" x 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a narrow, nested outline: you arrange fields into ordered **indent levels** (each level can stack several fields at the same indentation) and rows nest and merge by those levels, so it flows down the page.

## Status

**Working end to end:** paste → parse (SheetJS) → nest → render → live preview → **Copy all**. Paste several tables (managed in a left Sections rail), configure each table's pivot, style the levels once for all tables, and copy all sections at once. See the [roadmap](./docs/ROADMAP.md) for what's out of scope (`.docx`).

## The pivot view

Add fields from the **Add fields** pool, then shape the **Structure**, an indented list drawn with Windows-Explorer-style **tree connector lines** (│ ├ └) to show the nesting: each field sits at an indent level, and ◄/► move it shallower/deeper. Stack several fields at one level to show them together at the same indent; ▲/▼ reorder and ✕ removes — the ◄ ► ▲ ▼ ✕ buttons do all the reordering and indenting. Rows nest by the levels and **merge** when their values match across a level's fields. Each line reads `Field name: value` — per field, toggle the `Field name:` label off (just the value), or **bold**/**italic**/**underline** it (`Aa`/`B`/`I`/`U`; underline covers just the name) to make a long list scannable, and **sort** the groups at that field (↕ off / ↑ ascending / ↓ descending; numeric *and* text aware, so `2` sorts before `10`). A per-level **matrix** (one row per indent level) co-locates each level's controls: a Marker/Number cell (a **Marker** select prefixing `1.`/`a.`/`i.`/etc. in **Custom** mode, or a **Show number** checkbox in **Multilevel** mode, hidden in **Off** mode), a **Word heading** checkbox, and a **Look** cell — a color swatch plus a font/size/bold popover — that sets that depth's shared body-row look. A single **Blank line between top-level groups** checkbox beside the Markers control adds a blank line after each top-level group (on/off, no longer per-level). A per-table **Markers** control (a 3-mode dropdown: **Off (none)** = no markers or numbers / **Custom (per level)** = each level's chosen symbol / **Multilevel numbers**; default Custom) picks the body's marker style; Multilevel draws static **multilevel numbers** from a chosen **Start** — the exact first number (e.g. `5.1` → `5.1`, `5.1.1`, `5.1.1.1`), so set it to `5.1` to nest a body under a `5.0` section title — as plain body text, not Word auto-numbering, so nothing becomes a Word heading and the preview shows the real numbers; when on, the number replaces that level's marker, and the matrix's **Show number** checkbox hides a level's number. Hidden levels are *transparent*: the numbers follow the gap (e.g. number the type + text but hide the title → `5.1` then `5.1.1`, `5.1.2`, never a number under a missing parent). The matrix's **Word heading** checkbox maps a level's rows to a real Word heading (so they show in the Navigation pane and collapse, flush-left) — handy on just the top level so each section is in the document outline; Word numbers those rows. An optional **Section title** lives in its own **Section Header** group — a **Heading** dropdown (None / Heading 1–4) plus its own look (font / size / **B** / **I** / **U** / color); pick a Heading (1–4, each mapped at its correct outline level) and the title maps to that Word heading on a *Use Destination Styles* paste while your **full title look** — font, size, color, **B**/**I**/**U** — carries through inline, so the title pastes exactly as shown in the preview; Word adds only the heading's auto-number and outline/Navigation-pane entry (plus any all-caps effect). A **JSON** tab beside the live preview inspects the raw parsed grid.

## Multiple tables

Paste table after table — each becomes a row in the left **Sections rail** (one edited at a time, cap 100), inside a 4-pane IDE layout: top global controls · left rail · center builder · pinned right live preview. The shared body styling is reached right in the per-level matrix's **Look** column (a color swatch + a font / size / bold popover per depth), with only the document-wide **Body font**, **Indent/level**, and **Reset levels** in a compact **⚙ Document** popover in the top command band — all badged *All tables* so they apply to every section; each table keeps its own fields and title. Export every section at once with **Copy all**.

## Stack

- Next.js 16 (App Router) + TypeScript
- SheetJS (`xlsx`, official CDN tarball) for parsing
- Tailwind CSS v4 — Microsoft **Fluent/Office** reskin (Teams-blue accent, Segoe UI)
- Client-side only — no backend, no database, nothing uploaded or persisted.

## Getting started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000), click the paste zone, and press **Ctrl/Cmd + V** with a cell range copied from Excel or Google Sheets. Add fields and set their indent levels; the live preview updates. Click **Copy all** and paste into a Word document (with *Use Destination Styles* so the title maps to your Heading 1).

## Run with Docker

The app ships a `Dockerfile` (Next.js standalone build, ~305 MB image) and a `docker-compose.yml`. With [Docker](https://www.docker.com/products/docker-desktop/) installed:

```bash
docker compose up --build
```

Then open [http://localhost:3000](http://localhost:3000). Press **Ctrl/Cmd + C** in the terminal (or `docker compose down`) to stop. Everything runs client-side inside the container — nothing is uploaded or persisted.

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack, status
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, pivot, pipeline diagram
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order & status

## Project layout

```
app/                 App Router pages (home renders the full-height app shell)
components/
  PasteInput.tsx     app shell (4-pane IDE): paste/append, tables[] + shared styles, left Sections rail, pinned preview (Preview/JSON tabs) + Copy all (all sections)
  TableCard.tsx      one table's center builder: Section Header group (title + heading dropdown + title look) + Levels group (structure + per-level controls)
  tableModel.ts      TableState + tableToHtml (per-table nest->render) + bucket helpers (add/remove/indent/outdent/move/unusedColumns)
  RenderedPreview.tsx renders the pivot HTML (live preview; ws-title + [data-level] CSS)
  JsonPreview.tsx    shows the raw parsed Grid as JSON
  Popover.tsx        shared dependency-free popover (Document + Look popovers)
lib/
  types.ts           PivotNode/PivotLine + FieldLabel + raw Grid
  parser.ts          SheetJS clipboard -> Grid
  mapper.ts          rowsToPivotTree (Grid + indent buckets + sortDirs -> PivotNode[]; sort post-pass) + cellToString
  renderers.ts       renderPivotTree (tree -> HTML fragment; labels + markers + breakAfter spacers + static multilevel numbering) + marker/numbering helpers
  clipboard.ts       Word-friendly clipboard wrapper (buildWordHtml / htmlToPlainText; HeadingStyle / LevelStyle)
docs/                OVERVIEW, ARCHITECTURE, ROADMAP
```

# PivotParagraph

Turn a wide Excel table into a **Word-ready nested outline**. Paste one or more tables copied from Excel or Google Sheets; each is restructured into an Excel-pivot-style hierarchy and copied to the clipboard so it pastes straight into Microsoft Word.

## Why

Wide Excel tables don't fit an 8.5" × 11" Word page — columns bleed off the edge and the table becomes unreadable. Rather than shrink or split the grid, this app turns one wide table into a narrow, nested outline: you arrange fields into ordered **indent levels** (each level can stack several fields at the same indentation) and rows nest and merge by those levels, so the data flows *down* the page instead of off the side.

```
Region, Country, Product, Units, Revenue   ─────▶   1. Region: Americas
(5 columns bleeding off the page)                       a. Country: United States
                                                            i. Product: Laptops
                                                               Units: 120
```

## Quick start

```bash
npm install
npm run dev
```

Open <http://localhost:3000>, click **＋ Add section** in the left rail (or just press **Ctrl/Cmd + V** anywhere with a cell range copied from Excel) — or drop an `.xlsx` / `.csv` file into the modal. A fresh paste lands **blank on purpose**: add the fields you want one by one from the pool and arrange them with the ◀ ▶ ▲ ▼ buttons. Then **Copy section** (or **Ctrl/Cmd + Enter**) and paste into Word.

## Share it as one file — no hosting

Because the app is 100% client-side, it also builds to a **single self-contained `.html`**:

```bash
npm run build:single      # → dist-single/index.html (~660 KB)
```

That one file contains everything (React, SheetJS, styles, app code) with **zero external requests**. Double-click it and it runs — no server, no install. Put it on a shared network drive and your whole team can use it; update it by replacing that one file.

> Open it in Chrome or Edge, and verify **Copy section → paste into Word** once on a colleague's machine — a managed browser policy can restrict clipboard access. Copy falls back to a legacy clipboard path automatically if the modern API is refused.

## The builder

Three cards, split by **scope** — that split is the point, so it's always clear what a control will change:

### 1. Section Definition
The title text, a **Word heading** dropdown (None / Heading 1–4 / **Custom style…** for any style name in your template, e.g. `TBL_TITLE`), and its **Look** (font / size / color / **B** / **I** / **U**).

### 2. Fields & Title Definition — *per field*
Add fields from the pool, then shape one row per field, drawn with Explorer-style tree lines (│ ├ └). Arrange with **◀** outdent · **▶** indent · **▲ ▼** reorder · **✕** remove — or flip the **Drag to reorder** toggle in the card's top-right for a ⋮ grip per row (off by default; dragging is up/down only, so ◀ ▶ stay the level controls, and ▲ ▼ hide since the grip replaces them). The row you just changed **briefly highlights** so you don't lose it mid-move.

Per field: toggle the `Field name:` label off, **bold**/*italic*/underline it (`Aa` `B` `I` `U` — these affect **only** the label, the text before the colon), and **sort** that field's groups (↕ / ↑ / ↓ — numeric *and* text aware, so `2` sorts before `10`).

Depth is capped at **9 levels** (the most Word shows) with a live N/9 chip; a field added past the cap stacks into the deepest level. Fields stacked at one level show together at the same indent and merge when their values match.

### 3. Body Text Definition — *per level*
The **Markers** mode (Off / Custom / Multilevel + a **Start** number) plus **New page per group**, then one compact row per indent level:

| Column | Scope | What it does |
| ------ | ----- | ------------ |
| **Marker** | this section | Type (`1` `a` `A` `i` `I` `•` `–` None) × a free-text separator (`.` `)` `:` `;` `--` … ≤20 chars) |
| **Label sep** | this section | The text between a field's label and its value on this level — any string (`: ` `; ` ` — ` …) |
| **Line break before** | this section | A blank line right BEFORE this row's line — on the level's first sub-row it lands above the whole group; on a stacked `2.2` sub-row it's a per-field blank BETWEEN the stacked lines (`1.1 Country` → blank → `Product`) |
| **Line break after** | this section | A blank line after each of this level's whole groups (`…8.2 x` → break → `9 Region`) |
| **Heading** | this section | Map the level to a real Word heading + an Auto/H1–H9 rank chip |
| **Look** | level rows: **all sections** · stacked fields: this section | Color swatch + **Aa▾** popover (font / size / **B** / **I** / **U**); on a stacked level every sub-row styles ITS OWN field's line, and that per-field Look fully defines the line (it can un-bold what the level look bolds) |

**Look is shared by depth across every section** — restyling level 2 restyles level 2 everywhere (it's badged accordingly). A level's Look restyles the **whole line**; the Rows card's `B`/`I`/`U` touch only the label. In *Multilevel* mode the Marker cell becomes a **Show number** toggle, and on a heading-mapped level it reads *"Word numbers"* since Word supplies that number itself. Line breaks are positional — always relative to the row's own line — and adjacent blanks collapse to a single one, so the options compose without doubling and a section never starts or ends on a stray blank.

## Preview & copy

The pinned right pane shows the active section on a **true-to-scale US-Letter page** — a proportional 1in margin frame with measured pagination (dashed page boundaries) and an honest *"Showing: your Keep-Source look"* label. Page counts are a best-effort **estimate**: Word wraps in your template's font, so a real break can shift a line. A **Table** tab shows the raw pasted grid as it came from Excel.

**Copy section** writes the active section to the clipboard as its own standalone Word document, then an inline banner reminds you which paste option to pick — *Use Destination Styles* (adopt your document's Heading styles) vs *Keep Source Formatting* (keep the preview look). Avoid *Merge Formatting*: it strips the heading mapping.

## Multiple sections

Each pasted table is a row in the left **Sections rail** — **its own standalone Word document**, not a stacked section. Drag to reorder (or grip ⋮ + ↑/↓), **⧉** to duplicate, ✕ to remove; cap 100.


Document-wide settings — **Line spacing** and **Reset levels** — live in the **⚙ Document** popover. (There is deliberately no document-wide font, separator, or indent control: the base font is **Aptos 11** — Microsoft 365's standard — with per-level pins in each Look popover; label separators are set per field in the Body Text Definition matrix; and indentation is **derived from the outline geometry** — each level's line nests under its parent's text, sized by the measured number/marker widths, with 0.25in steps when a level has no marker.) It also holds the **Style preset**: export your formatting (level looks, markers, numbering, headings, blank lines, fonts, and the Rows-card label bold/italic/underline — everything keyed by *level*, never by column) as a `.json`, then import it onto next week's completely different table and every section adopts the house style in one undoable step.

Everything is **undoable** (Ctrl+Z / Ctrl+Y), and the whole workspace **auto-saves to your browser** (local only, nothing uploaded). **Ctrl/Cmd + K** opens a command palette over every action.

## Stack

- **Next.js 16** (App Router) + **React 19** + **TypeScript**
- **SheetJS** (`xlsx`, official CDN tarball) for parsing
- **Tailwind CSS v4** — Microsoft **Fluent/Office** reskin (Teams-blurple accent, Segoe UI)
- **Vite** + `vite-plugin-singlefile` for the single-file build (build-time only)
- **Client-side only** — no backend, no API routes, no database. Nothing is uploaded.

## Docs

- [docs/OVERVIEW.md](./docs/OVERVIEW.md) — problem, goals, stack, status
- [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) — data model, pivot, pipeline
- [docs/ROADMAP.md](./docs/ROADMAP.md) — build order & status

## Project layout

```
app/                   App Router shell (layout + page); globals.css design tokens
index.html             entry for the SINGLE-FILE build only (Next ignores it)
main.tsx               Vite entry: mounts the same PasteInput + globals.css
vite.config.ts         single-file build (IIFE + inlined assets, file://-safe)
components/
  PasteInput.tsx       app shell (4-pane IDE): ingestGrid (dropEmptyColumns +
                       new sections land blank) + paste-anywhere,
                       ＋ Add-table modal, Ctrl+K palette, tables[] + shared styles,
                       localStorage persistence, undo/redo, Copy section
  SectionsRail.tsx     left rail: select / drag-reorder / duplicate (⧉) / remove
  AddTableModal.tsx    intake modal: paste / drop / Browse / example → preview
                       ("Read N×M" + full scrollable table) → Add section
  CommandPalette.tsx   Ctrl/Cmd+K fuzzy launcher; portals to <body>
  TableCard.tsx        the three builder cards: Section Title · Rows (per field) ·
                       Level formatting (per level)
  tableModel.ts        TableState + tableToHtml + bucket helpers (add/remove/
                       indent/outdent/move) + MAX_LEVELS=9 + dropEmptyColumns/
                       bodyGrid + estimateSectionStats + newTable/makeExampleTable
  RenderedPreview.tsx  true-to-scale US-Letter proxy: 1in margins, measured
                       pagination (ResizeObserver → "Page N" rules)
  GridTable.tsx        Table view: the raw pasted grid, as-is
  Popover.tsx          shared dependency-free popover (Document + Look)
lib/
  types.ts             PivotNode/PivotLine + FieldLabel + raw Grid
  parser.ts            clipboard → Grid (parseClipboard) and file → Grid (parseFile);
                       shared firstSheetToGrid
  stylePreset.ts       style preset build/validate (formatting only, level-keyed)
  mapper.ts            rowsToPivotTree (buckets → PivotNode[], sort post-pass)
  renderers.ts         renderPivotTree (tree → HTML fragment)
  clipboard.ts         buildWordHtml / htmlToPlainText / writeRichClipboard
  persistence.ts       versioned localStorage save/load/clear
docs/                  OVERVIEW, ARCHITECTURE, ROADMAP
```

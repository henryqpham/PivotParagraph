# Overview

## What this is

A web app that converts **pasted Excel data into a Word-ready nested outline**. You paste one or more tables copied from Excel (or Google Sheets); the app restructures each into an Excel-pivot-style hierarchy ("Rows area") and copies it to the clipboard so it pastes into Microsoft Word.

## The problem it solves

Wide Excel tables do not fit on an 8.5" x 11" Word page and become unreadable once columns bleed off the right edge. Instead of shrinking or splitting the grid, this app **restructures** spreadsheet data into a narrow, nested outline — you arrange fields into ordered indent levels (each level can stack several fields at the same indentation), and rows nest and merge by those levels — so it flows down a Word page instead of off the side.

## Tech stack

- **Next.js 16** (App Router) + **TypeScript**
- **SheetJS** (`xlsx`, official CDN tarball) for parsing clipboard / Excel data
- **Tailwind CSS** (v4) for styling
- **Client-side only** — no backend, no API routes, no database. All parsing runs in the browser; nothing is uploaded. The workspace is saved to the browser's own `localStorage` (on your machine only) so it survives a refresh.

## What works today

The full pivot pipeline is implemented end to end, for multiple tables:

1. **Paste → parse.** A client component captures each paste and parses it with SheetJS into a raw Grid, appending a new table. A copied table drops in on **Ctrl/Cmd + V from anywhere** on the page (unless a text field is focused), so the paste zone doesn't need focus. Tables are managed in a left Sections rail (one edited at a time, cap 100, **reorderable with ▲▼**) within a 4-pane IDE layout — top global controls, left rail, center builder, and a pinned right live preview. A first-run **onboarding** empty state (value prop + 1-2-3 steps + a **Try an example** button that loads a wide demo grid) gets a newcomer started with no spreadsheet; **Remove section** and **Clear all** are undoable (Clear all is a two-step inline confirm), and the whole workspace **auto-saves to the browser's `localStorage`** so a refresh doesn't lose it.
2. **Nest → tree.** `rowsToPivotTree` nests rows by an ordered list of **indent buckets** (`pivotLevels`) into an arbitrary-depth `PivotNode` tree. Each bucket is one indent level holding one or more fields; fields stacked in a bucket render at the same indent and rows merge by the composite of that level's values.
3. **Render → preview.** The tree is rendered to HTML and shown in the pinned right-pane live preview (with a **Table** tab beside it that shows the raw pasted grid as an actual table, and a **JSON** tab to inspect the same Grid as raw JSON). A **This section / All sections** toggle on the Preview tab switches between the section you're editing and the full stacked document `Copy all` produces (each section on its own paper page, from the same fragments the export concatenates). The **Structure** list is drawn with Windows-Explorer-style tree connector lines (│ ├ └) to show the nesting, and is reordered/indented entirely via the ◄ ► ▲ ▼ ✕ buttons. A per-level **matrix** (one row per level) sets each level's Marker/Number cell (a **Marker** select `1./a./i.`/etc. in **Custom** mode, or a **Show number** checkbox in **Multilevel** mode, hidden in **Off** mode), a **Word heading** checkbox, and a **Look** cell (a color swatch + a font/size/bold popover) for that depth's shared appearance; a single **Blank line between top-level groups** checkbox beside the Markers control adds a blank line after each top-level group (on/off, no longer per-level). Per field you can hide/bold/italic/underline the `Field name:` label and **sort** the groups at that field (↕ off / ↑ asc / ↓ desc, numeric + text aware).
4. **Copy all.** The single **Copy all** button exports every section together as one document, writing `text/html` (+ a `text/plain` fallback) to the clipboard (no download), then shows an inline + screen-reader-announced confirmation of how many sections were copied and which Word paste option to pick (**Keep Source Formatting** to keep your look, or **Use Destination Styles** to adopt the template heading). The **Section title** maps to your document's chosen Heading 1–4 ("5.0") at the right outline level on a *Use Destination Styles* paste, while the full title Look you set (font, size, color, bold, italic, underline) carries through inline on top, so the title pastes exactly as previewed — Word supplies only the heading's number and outline/nav-pane entry (and any all-caps effect). A per-table **Markers** control (a 3-mode dropdown — **Off (none)** = no markers or numbers / **Custom (per level)** = each level's chosen symbol / **Multilevel numbers**; default Custom) sets the body's marker style; the Multilevel mode draws static **multilevel numbers** from a chosen **Start** (the exact first number, e.g. `5.1` → `5.1`, `5.1.1`, `5.1.1.1`) as plain body text — not Word auto-numbering, so nothing becomes a Word heading and the numbers show in the preview verbatim (with a per-level toggle to hide levels; hidden levels are transparent, so the numbers follow the gap). The body is always styled body text.

**Styling.** The shared per-depth body look lives right in the per-level matrix's **Look** column — a color swatch plus a font/size/bold popover per depth, tinted and badged *All tables* so it applies to every section — while only the document-wide **Body font**, **Indent/level**, and **Reset levels** sit in a compact **⚙ Document** popover in the top command band. Both the Look and Document popovers share one dependency-free `Popover` component. Only the **Section title** maps to a real Word heading (one Nav-pane entry per table); the body — including any app-drawn multilevel numbers — stays out of Word's outline.

See [ARCHITECTURE.md](./ARCHITECTURE.md) for the data model, the pivot, and the pipeline, and [ROADMAP.md](./ROADMAP.md) for status.

## Out of scope

- **`.docx` generation** — export is HTML-on-clipboard only ("Copy all"). There is no download.

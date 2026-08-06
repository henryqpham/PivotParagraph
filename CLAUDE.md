@AGENTS.md
# PivotParagraph

Client-side Next.js app: paste one or more Excel/Sheets tables; each is restructured into an Excel-pivot-style **nested outline** ("Rows area" — ordered indent levels, each holding ≥1 stacked fields; rows nest and merge by level), rendered as Word-ready HTML, and copied to the clipboard per section. Solves: wide tables don't fit an 8.5x11 Word page. The title and any heading-marked levels map to real Word Heading styles; all other body rows use the app's own inline per-level styling (11pt black default, base font FIXED at Aptos — Microsoft 365's default; the document-wide font control was removed — with a per-level Look pin).

UI is a 4-pane IDE layout: top command band (Ctrl+K command palette — keyboard-only, ⚙ Document popover — global line spacing / reset / **style preset** export-import; NO indent control — indentation is DERIVED), left **Sections rail** (＋ Add section pinned on top — THE add affordance; each table = its OWN standalone Word document; drag-reorder, ⧉ duplicate, ✕ remove), center builder (three cards by SCOPE: Section Definition / **Fields & Title Definition** per-field / **Body Text Definition** per-level), pinned right preview (Preview on a true-to-scale Word page / Table). Intake is the ＋ Add-table modal (paste / drop a file / example, previewed) + paste-anywhere; a fresh paste lands BLANK (fields added one by one — auto-arrange was removed on feedback). Export is per-section **Copy section** (clipboard only). Workspace auto-saves to `localStorage` (local-only); undo/redo; first-run onboarding.

**Detailed docs live in `docs/` — OVERVIEW.md (product + features), ARCHITECTURE.md (data model, pipeline, clipboard/Word mechanics), ROADMAP.md. Read the relevant one before non-trivial changes. Keep THIS file lean: pointers and rules, not prose.**

## Commands
- `npm run dev` — dev server on localhost:3000
- `npm run build` — production build, must pass with no type errors
- `npm run build:single` — Vite single-file build → `dist-single/index.html` (one self-contained .html, runs from file://)
- `npm run lint` — ESLint check

## Stack
- Next.js 16, App Router, TypeScript
- Tailwind CSS v4 (CSS-first, no tailwind.config.js)
- SheetJS for parsing
- Vite + `vite-plugin-singlefile` — build-time only, for the single-file target

## Hard Rules
- SheetJS install is the official CDN tarball: `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`. NEVER install `xlsx` from npm (stale 0.18.5).
- Client-side only. No API routes, no backend, no database. Parsing runs in the browser.
- Files at repo root. No `src/` directory. Import alias `@/*` maps to `./*`.
- SheetJS clipboard parse: `getData('text/html')` first, `getData('text/plain')` TSV fallback. Use `sheet_to_json(ws, { header: 1, blankrows: false, defval: "", raw: false })` — keep `defval: ""` so blank cells stay empty strings and column positions stay aligned.
- Keep docs in sync: when a change alters the pivot behavior, data model, pipeline, or a user-facing control, update `README.md` and the relevant `docs/` file (OVERVIEW / ARCHITECTURE / ROADMAP) in the SAME change. Stale docs are a bug. Update this file only when a rule or pointer here becomes wrong — do NOT grow it back into full prose documentation.

## Architecture (file map — details in docs/ARCHITECTURE.md)
Pipeline per table: add (paste / file) → `parseClipboard`/`parseFile` (Grid) → `dropEmptyColumns` → append blank → `rowsToPivotTree` (`PivotNode[]`) → `renderPivotTree` (HTML fragment) → live preview + `buildWordHtml` clipboard write. `tableToHtml` is the single source for both preview and export.
- `app/page.tsx` — mounts the full-height `PasteInput` shell
- `components/PasteInput.tsx` — app shell + parent state: `tables[]`, shared level/title styles, active preview + view toggle, persistence, undo/redo, reorder/duplicate, copy confirmation, onboarding; `ingestGrid` = dropEmptyColumns → append; new sections land BLANK (the auto-arrange + header-match-reuse assists were both removed on feedback); intake via the ＋ Add-table modal + paste-anywhere (double-ingest guarded by `addModalOpenRef`); Ctrl+Enter copy; Ctrl+K command palette
- `components/SectionsRail.tsx` — left rail (select / drag + keyboard reorder / duplicate / remove)
- `components/TableCard.tsx` — center builder, THREE cards split by SCOPE (deliberate: Rows' Aa/B/I/U touch only the LABEL before the colon, a level's Look restyles the WHOLE line): **Section Definition** (title text + Word-heading dropdown incl. a **Custom style…** text input for any destination style name + shared title look) · **Fields & Title Definition** = per FIELD (Add-fields pool; ONE row per field with tree guides, plain-text field names (micro-preview + sample were removed as noise), label/sort toggles, and ◀ ▶ ▲ ▼ ✕ buttons — plus an OPT-IN drag-to-reorder toggle in the card header (`dragRows`, default off): a ⋮ grip per row, VERTICAL-only via `reorderField` (fixed indent skeleton — a drop can't indent/stack), ▲ ▼ hidden while on; `flashRow`/`applyMove` briefly tint the changed row) · **Body Text Definition** = per LEVEL (Markers mode + page-break checkbox, then a matrix with ONE SUB-ROW PER FIELD (`2.1`/`2.2` pills on stacked levels): Level | Marker split **Type + free-text separator** (≤20 chars, escaped; disabled for symbol/none; "Word numbers" note on heading levels; level-scoped, first sub-row) | **Label sep** per FIELD (`labelSepByCol[col]`, free text) | **Line break before/after** per level (`breakAfter[i]` = blank BEFORE the level's line, positional; `gapAfter[i]` = blank after the whole group; first sub-row) + a per-FIELD between-stacked-lines checkbox on sibling sub-rows (`fieldBreakBefore[col]`) | Word heading + Auto/H1–H9 rank chip (first sub-row) | **Look** = the SAME `LookControl` swatch+Aa▾ the Section Title uses — LEVEL look (*all tables*) on the first sub-row, PER-FIELD override (*this section*, `fieldLooks[col]` → inline styled run) on stacked siblings); stacked pills read `2.1`/`2.2` in BOTH cards; EVERY numbered line pastes with a hanging indent (`data-hang`/`data-cont`; hang = the lead's MEASURED `ws-lead` width — canvas in `buildWordHtml`, DOM pre-paint in the preview) and INDENTATION IS DERIVED: each level nests at its parent's text position (0.25in steps when lead-less), so `1.1.1` sits exactly under the parent's text; the ⚙ Indent-per-level control was removed
- `components/AddTableModal.tsx` — ＋ Add-table intake modal (paste / drop / Browse file / Try-example → preview "Read N×M" → `onCommit(grid)` → `ingestGrid`); replaces the retired `pasteZone`
- `components/CommandPalette.tsx` — Ctrl/Cmd+K fuzzy launcher (subsequence match, grouped, keyboard nav, kbd-chip shortcuts) over PasteInput's handlers; portals to `<body>`
- `components/tableModel.ts` — `TableState`, `tableToHtml`, bucket helpers (add/remove/indent/outdent/move), `MAX_LEVELS = 9`, `dropEmptyColumns`, `bodyGrid`, `newTable`, example grid, `estimateSectionStats`, `LevelInput`/`DEFAULT_LEVEL`
- `components/RenderedPreview.tsx` / `GridTable.tsx` — the TWO right-pane views (rendered on a true-to-scale US-Letter proxy w/ 1in margins + measured `ResizeObserver` pagination + a "Showing: your Keep-Source look" label / the raw grid). The JSON view and the Header-row picker were both removed; `headerRow?` remains a legacy field honored on read.
- `appVersion.ts` — build-time git version (`1.<commit count> · <hash>`) inlined by BOTH configs as `NEXT_PUBLIC_APP_VERSION`; shown as the fixed bottom-right badge in PasteInput
- `index.html` / `main.tsx` / `vite.config.ts` — the SINGLE-FILE build only (Next ignores them). `main.tsx` mounts the same `PasteInput` + `globals.css`; IIFE output + stripped `type="module"` + a `DOMContentLoaded` guard make it work from `file://`.
- `components/Popover.tsx` — shared popover helper
- `lib/types.ts` — `PivotNode`/`PivotLine`, `FieldLabel`, Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid (`parseClipboard`) + file → Grid (`parseFile`, `XLSX.read` arrayBuffer); shared `firstSheetToGrid`
- `lib/stylePreset.ts` — style-preset build/validate/apply: FORMATTING only (globals + LEVEL-indexed per-table settings + `labelsByLevel`, the Rows-card label emphasis projected by level and re-applied positionally via `applyLabelsByLevel`), versioned envelope; applied to all sections on import (PasteInput `exportStylePreset`/`handlePresetFile`)
- `lib/mapper.ts` — `rowsToPivotTree` (buckets → merged tree) + `sortTree` post-pass
- `lib/renderers.ts` — `renderPivotTree` (title row + `data-level`/`data-heading`/`data-break` paragraphs, markers, app-drawn multilevel numbers; escapes all user text; spacer post-pass collapses adjacent blanks + trims doc edges)
- `lib/persistence.ts` — versioned `localStorage` save/load/clear (generic, SSR-safe)
- `lib/clipboard.ts` — `buildWordHtml` Word wrapper, `HeadingStyle`/`LevelStyle`/`TitleStyle`, `headingLevel`, `htmlToPlainText`, `writeRichClipboard` (async Clipboard API → `execCommand` fallback, so Copy survives a `file://` origin)

## Key invariants
- Pivot structure: `pivotLevels: number[][]` — ordered indent buckets of grid columns; per-COLUMN `fieldLabels`/`sortDirs` (survive remove/re-add); per-TABLE `markerSpecs?`(split `{type, delim}` per level; legacy fused `markers?: MarkerKind[]` migrated on read by `resolveMarkerSpecs`)/`breakAfter` (blank BEFORE the level's line — positional semantics)/`gapAfter?` (blank after the whole group, incl. sole children)/`fieldBreakBefore?` (per-col blank between stacked lines)/`numbering`/`headingLevels`; optional `headerRow?`/`pageBreakBefore?`/`headingRanks?` (per-level Word-heading rank override, 1–9; 0/absent = auto — one rank under the previous heading, pins included, so only a pin can create a rank gap) — read via `?? default` so old sessions/imports stay compatible.
- 9-level depth cap (`MAX_LEVELS`) is enforced end-to-end: `addField` STACKS past it rather than making a 10th level, renderer clamps `data-level`, preview has exactly 9 rules, `buildWordHtml` matches only 1–9.
- Word heading mapping uses REAL `<h1>`–`<h6>` elements whose `<style>` rule declares the FULL source look (mso-style-name/outline-level + color/font/size/bold): declared props are cleanly REPLACED by the destination Heading style on a Use-Destination-Styles/default paste, undeclared ones ride along as polluting direct formatting; a Keep-Source-Formatting paste never maps named styles (rows land Normal + outline level = nav pane) and shows the declared look instead — verified empirically against desktop Word (COM paste matrix, July 2026). Every other body row gets INLINE direct formatting so it survives any paste. K 7–9 fall back to the Mso class rule.
- Level styles are GLOBAL by depth — slot = `bucket + (sectionTitle ? 1 : 0)`, matching the renderer; Marker / Word-heading are per-table. `LevelStyle` (and `TitleStyle`) carry font/size/color + **bold/italic/underline**, edited via the shared `LookControl` (swatch + Aa▾ popover) used IDENTICALLY for the Section Title and every level row — the emphasis is emitted as `<b>`/`<i>`/`<u>` runs on the body `<p>` so it survives any paste.
- Style inputs are form-controlled (hex color, allow-listed `FONTS`/spacings, clamped numbers); the free-text inputs — per-level label/marker separators (≤20 chars) — are HTML-escaped by the renderer, and the custom title style name is allow-list sanitized by `sanitizeStyleName` (letters/digits/space/`._-`) before it lands in `mso-style-name`.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the pivot convention or the data model.
- Never: install `xlsx` from npm, add a backend.

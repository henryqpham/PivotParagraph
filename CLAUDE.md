@AGENTS.md
# Excel → Word Pivot

Client-side Next.js app: paste one or more Excel/Sheets tables; each is restructured into an Excel-pivot-style **nested outline** ("Rows area" — ordered indent levels, each holding ≥1 stacked fields; rows nest and merge by level), rendered as Word-ready HTML, and copied to the clipboard per section. Solves: wide tables don't fit an 8.5x11 Word page. The title and any heading-marked levels map to real Word Heading styles; all other body rows use the app's own inline per-level styling (11pt black default, font inherited from the ⚙ Body font — Arial unless changed — with a per-level Look pin).

UI is a 4-pane IDE layout: top command band (⚙ Document popover — global body font / line spacing / label separator / indent / reset), left **Sections rail** (each table = its OWN standalone Word document; drag-reorder, ⧉ duplicate, ✕ remove), center builder, pinned right preview (Preview / Table / JSON). Export is per-section **Copy section** (clipboard only). Workspace auto-saves to `localStorage` (local-only); undo/redo; first-run onboarding + paste-anywhere.

**Detailed docs live in `docs/` — OVERVIEW.md (product + features), ARCHITECTURE.md (data model, pipeline, clipboard/Word mechanics), ROADMAP.md. Read the relevant one before non-trivial changes. Keep THIS file lean: pointers and rules, not prose.**

## Commands
- `npm run dev` — dev server on localhost:3000
- `npm run build` — production build, must pass with no type errors
- `npm run lint` — ESLint check

## Stack
- Next.js 16, App Router, TypeScript
- Tailwind CSS v4 (CSS-first, no tailwind.config.js)
- SheetJS for parsing

## Hard Rules
- SheetJS install is the official CDN tarball: `https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`. NEVER install `xlsx` from npm (stale 0.18.5).
- Client-side only. No API routes, no backend, no database. Parsing runs in the browser.
- Files at repo root. No `src/` directory. Import alias `@/*` maps to `./*`.
- SheetJS clipboard parse: `getData('text/html')` first, `getData('text/plain')` TSV fallback. Use `sheet_to_json(ws, { header: 1, blankrows: false, defval: "", raw: false })` — keep `defval: ""` so blank cells stay empty strings and column positions stay aligned.
- Keep docs in sync: when a change alters the pivot behavior, data model, pipeline, or a user-facing control, update `README.md` and the relevant `docs/` file (OVERVIEW / ARCHITECTURE / ROADMAP) in the SAME change. Stale docs are a bug. Update this file only when a rule or pointer here becomes wrong — do NOT grow it back into full prose documentation.

## Architecture (file map — details in docs/ARCHITECTURE.md)
Pipeline per table: paste → `parseClipboard` (Grid) → `rowsToPivotTree` (`PivotNode[]`) → `renderPivotTree` (HTML fragment) → live preview + `buildWordHtml` clipboard write. `tableToHtml` is the single source for both preview and export.
- `app/page.tsx` — mounts the full-height `PasteInput` shell
- `components/PasteInput.tsx` — app shell + parent state: `tables[]`, shared level/title styles, active preview + view toggle, persistence, undo/redo, reorder/duplicate, copy confirmation, onboarding, paste-anywhere, Ctrl+Enter copy, arrangement-reuse offer (`headerSignature` match on ingest)
- `components/SectionsRail.tsx` — left rail (select / drag + keyboard reorder / duplicate / remove)
- `components/TableCard.tsx` — center builder: Section Title group (title text + Word-heading dropdown incl. a **Custom style…** text input for any destination style name + shared title look; renamed from "Section Header" — "header" was overloaded) and Levels group (Add-fields pool, Structure tree with ◄ ► ▲ ▼ ✕ + label/sort toggles, Markers control + blank-line/page-break checkboxes, per-level matrix: Marker/Number (a "Word numbers" note on heading levels), Word heading + Auto/H1–H9 rank chip-dropdown, Look columns)
- `components/tableModel.ts` — `TableState`, `tableToHtml`, bucket helpers, `MAX_LEVELS = 9`, `dropEmptyColumns`, `bodyGrid`, `newTable`, example grid, `estimateSectionStats`, `LevelInput`/`DEFAULT_LEVEL`
- `components/RenderedPreview.tsx` / `GridTable.tsx` / `JsonPreview.tsx` — right-pane views (rendered paper proxy / raw grid / raw JSON). The Header-row picker was removed by request; `headerRow?` remains a legacy field honored on read.
- `components/Popover.tsx` — shared popover helper
- `lib/types.ts` — `PivotNode`/`PivotLine`, `FieldLabel`, Grid types
- `lib/parser.ts` — SheetJS clipboard → Grid
- `lib/mapper.ts` — `rowsToPivotTree` (buckets → merged tree) + `sortTree` post-pass
- `lib/renderers.ts` — `renderPivotTree` (title row + `data-level`/`data-heading`/`data-break` paragraphs, markers, app-drawn multilevel numbers; escapes all user text)
- `lib/persistence.ts` — versioned `localStorage` save/load/clear (generic, SSR-safe)
- `lib/clipboard.ts` — `buildWordHtml` Word wrapper, `HeadingStyle`/`LevelStyle`/`TitleStyle`, `headingLevel`, `htmlToPlainText`

## Key invariants
- Pivot structure: `pivotLevels: number[][]` — ordered indent buckets of grid columns; per-COLUMN `fieldLabels`/`sortDirs` (survive remove/re-add); per-TABLE `markers`/`breakAfter`/`numbering`/`headingLevels`; optional `headerRow?`/`pageBreakBefore?`/`headingRanks?` (per-level Word-heading rank override, 1–9; 0/absent = auto — one rank under the previous heading, pins included, so only a pin can create a rank gap) — read via `?? default` so old sessions/imports stay compatible.
- 9-level depth cap (`MAX_LEVELS`) is enforced end-to-end: builder stacks past it, renderer clamps `data-level`, preview has exactly 9 rules, `buildWordHtml` matches only 1–9.
- Word heading mapping uses REAL `<h1>`–`<h6>` elements whose `<style>` rule declares the FULL source look (mso-style-name/outline-level + color/font/size/bold): declared props are cleanly REPLACED by the destination Heading style on a Use-Destination-Styles/default paste, undeclared ones ride along as polluting direct formatting; a Keep-Source-Formatting paste never maps named styles (rows land Normal + outline level = nav pane) and shows the declared look instead — verified empirically against desktop Word (COM paste matrix, July 2026). Every other body row gets INLINE direct formatting so it survives any paste. K 7–9 fall back to the Mso class rule.
- Level styles are GLOBAL by depth — slot = `bucket + (sectionTitle ? 1 : 0)`, matching the renderer; Marker / Word-heading are per-table.
- All style inputs are form-controlled (hex color, allow-listed `FONTS`/separators/spacings, clamped numbers); the ONE free-text style input — the custom title style name — is allow-list sanitized by `sanitizeStyleName` (letters/digits/space/`._-`) before it lands in `mso-style-name`.

## Escalation
- Autonomous: write code, run dev/build/lint, fix type errors.
- Confirm first: adding any dependency beyond SheetJS, creating API routes, changing the pivot convention or the data model.
- Never: install `xlsx` from npm, add a backend.

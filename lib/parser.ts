import * as XLSX from "xlsx";
import type { Grid, Row } from "./types";

/**
 * Parse clipboard contents into a Grid (array of row-arrays).
 *
 * Excel and Google Sheets place an HTML `<table>` on the clipboard under the
 * `text/html` flavor; we prefer that. But many sources (VS Code, chat apps, a
 * plain web-page selection, Word) put styled-but-table-less HTML on the clipboard
 * ALONGSIDE perfectly good tab-separated `text/plain` — and SheetJS THROWS on HTML
 * with no `<table>`. So we try the HTML first but fall through to `text/plain`
 * whenever the HTML parse throws or yields nothing, instead of aborting the paste.
 *
 * Returns an empty Grid when the clipboard holds nothing parseable.
 */
export function parseClipboard(data: DataTransfer): Grid {
  const html = data.getData("text/html");
  if (html && html.trim()) {
    const grid = tryParse(html);
    if (grid.length > 0) return grid;
    // else: non-table / empty HTML — fall through to the plain-text TSV below.
  }

  const text = data.getData("text/plain");
  if (text && text.trim()) {
    // Spreadsheet plain text is TAB-separated. Force tab as the delimiter when the
    // text contains any tab, so a MULTI-column paste whose cells legitimately hold
    // commas ("Smith, John") isn't shredded by SheetJS's frequency-based delimiter
    // sniffing (which prefers comma). With no tabs we let it sniff so a genuine CSV
    // paste still splits into columns — the unavoidable tradeoff is that a
    // SINGLE-column, comma-bearing, HTML-less paste is ambiguous ("Last, First" in
    // one column vs two CSV columns) and may still split; copying from the
    // spreadsheet itself (the HTML path above) avoids the ambiguity entirely.
    return tryParse(text, text.includes("\t") ? "\t" : undefined);
  }

  return [];
}

/** Parse one string as a sheet, returning [] on any SheetJS error (so a bad HTML
 *  attempt can fall through to plain text and no raw parser error reaches the UI). */
function tryParse(input: string, fs?: string): Grid {
  try {
    return parseString(input, fs);
  } catch {
    return [];
  }
}

function parseString(input: string, fs?: string): Grid {
  const workbook = XLSX.read(
    input,
    fs ? { type: "string", FS: fs } : { type: "string" },
  );
  const sheetName = workbook.SheetNames[0];
  const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;
  if (!sheet) return [];

  // header: 1   -> rows as arrays of cells (no header inference)
  // blankrows   -> drop fully empty rows
  // defval: ""  -> keep blank cells as "" so column positions stay aligned
  // raw: false  -> return formatted display strings
  return XLSX.utils.sheet_to_json<Row>(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
    raw: false,
  });
}

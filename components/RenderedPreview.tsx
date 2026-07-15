import { headingLevel, type HeadingStyle } from "@/lib/clipboard";

/**
 * Read-only display of the ACTIVE section's rendered pivot HTML (from
 * `renderPivotTree`). Each section is its own document now, so this always shows
 * exactly one section — the same fragment `copySection` copies to the clipboard,
 * so preview and export can't diverge.
 *
 * The HTML is injected via `dangerouslySetInnerHTML`, which is safe here:
 * `renderPivotTree` escapes every user-derived value (titles, field labels), so
 * the only raw markup is machine-generated tags, markers/static numbers, and
 * constant `style` attributes -- no user value lands in an attribute, so there is
 * no injection surface.
 *
 * Tailwind v4 preflight strips heading sizes. We restore spacing with
 * arbitrary-descendant variants, and the heading look (color/font/size/weight)
 * from a scoped `<style>` so the preview matches what a "Copy section" +
 * keep-source paste will produce. The style values are sanitized in the form
 * (hex color, allow-listed font, clamped sizes).
 */
// A white "paper" surface (a Word-page proxy) that stays white in BOTH themes,
// floating on the canvas with a Fluent depth shadow.
// Generous left padding (pl-14) so the hanging H-tokens sit in a comfortable
// gutter instead of pressing against the paper's edge — closer to a real Word
// page margin.
const previewClasses =
  "ws-preview rounded-[2px] bg-white p-6 pl-14 text-sm text-[#242424] shadow-[var(--shadow-4)] " +
  "[&_.ws-title]:mt-1 [&_[data-level]]:mt-1 [&_p]:my-0.5 [&_p]:leading-tight";

export function RenderedPreview({
  html,
  emptyHint,
  headingStyle,
  bodyFont,
}: {
  /** The active section's rendered HTML fragment ("" when it has no fields). */
  html?: string;
  emptyHint?: string;
  headingStyle: HeadingStyle;
  bodyFont: string;
}) {
  if (!html) {
    return (
      <p className="text-sm text-foreground/60">
        {emptyHint ??
          "Pasted, but nothing was rendered. Switch to JSON to inspect the raw grid."}
      </p>
    );
  }

  // Approximate look from headingStyle.levels (level 1 = the title row; levels
  // 2-9 the nested rows by depth). When a Word style name is mapped the real
  // template look only shows on paste; this preview uses the per-level look.
  const FALLBACK = { color: "#000000", font: "Arial", size: 11, bold: false };
  const step = headingStyle.indentStep ?? 0.2;
  const rule = (sel: string, i: number, indentIn: string) => {
    const lv = headingStyle.levels[i] ?? FALLBACK;
    const margin = indentIn ? `;margin-left:${indentIn}in` : "";
    return `.ws-preview ${sel}{color:${lv.color};font-family:'${lv.font}';font-size:${lv.size}pt;font-weight:${lv.bold ? 700 : 400}${margin}}`;
  };
  // Title (level 1, no indent); nested rows (per-level look + (n-1)*step indent).
  // App-drawn multilevel numbers are plain text in the fragment, so the preview
  // shows them verbatim. A heading row (`data-heading`) instead sits flush-left and
  // bold (it loses the body indent in Word) with a muted "#" where Word inserts the
  // live heading number on paste.
  // The title has its OWN shared look (set in the Section Title group), not the
  // level chart -- so read it from headingStyle.titleStyle, with italic/underline.
  const t = headingStyle.titleStyle ?? {
    color: "#000000",
    font: "Arial",
    size: 11,
    bold: false,
    italic: false,
    underline: false,
  };
  const titleRule =
    `.ws-preview .ws-title{color:${t.color};font-family:'${t.font}';font-size:${t.size}pt;` +
    `font-weight:${t.bold ? 700 : 400};font-style:${t.italic ? "italic" : "normal"};` +
    `text-decoration:${t.underline ? "underline" : "none"}}`;
  // A title mapped to a Word heading (headingStyleName set) is ALSO numbered by
  // Word on paste, exactly like a body heading row — so it gets the same hash
  // placeholder. The ::before only renders if a `.ws-title` actually exists, so a
  // titleless section shows nothing.
  const titleMapped = (headingStyle.headingStyleName ?? "") !== "";
  // Heading placeholder: a muted "H2"/"H3"/… token HANGING in the left margin
  // (absolutely positioned) so the heading TEXT stays flush-left — matching Word
  // exactly, where heading styles carry no indent. A constant-width token instead
  // of repeated Markdown hashes: "####" outgrew the paper's left padding at deep
  // levels, and the token also matches the matrix's H-chips (one visual
  // language). The mapped title shows its own outline level the same way.
  const headingTag = (k: number) => `H${Math.min(Math.max(k, 1), 9)}`;
  const hangingTag = (content: string) =>
    `content:"${content}";position:absolute;right:100%;padding-right:0.35em;opacity:0.4;font-weight:400`;
  const titleN = headingLevel(headingStyle.headingStyleName ?? "");
  // Body line spacing, mirroring buildWordHtml's clamp. Unlayered rule, so it
  // wins over the Tailwind-layered `leading-tight` fallback on the paper div.
  const rawSpacing = headingStyle.lineSpacing ?? 1.15;
  const spacing = Math.min(
    3,
    Math.max(1, Number.isFinite(rawSpacing) ? rawSpacing : 1.15),
  );
  const css =
    `.ws-preview p{line-height:${spacing * 100}%}` +
    titleRule +
    Array.from({ length: 9 }, (_, i) =>
      rule(`[data-level="${i + 1}"]`, i, (i * step).toFixed(2)),
    ).join("") +
    `.ws-preview [data-heading]{margin-left:0;font-weight:700;position:relative}` +
    Array.from(
      { length: 9 },
      (_, i) =>
        `.ws-preview [data-heading="${i + 1}"]::before{${hangingTag(headingTag(i + 1))}}`,
    ).join("") +
    // A page-break row: a dashed rule + generous space above it stands in for the
    // Word page boundary the paste will insert before this top-level group.
    `.ws-preview [data-break]{margin-top:22px;border-top:1px dashed #b0b0b0;padding-top:16px}` +
    (titleMapped
      ? `.ws-preview .ws-title{position:relative}` +
        `.ws-preview .ws-title::before{${hangingTag(headingTag(titleN))}}`
      : "");

  // No explanatory footnotes under the paper — the hash placeholders and the
  // dashed page-break rule carry their meaning; prose captions were judged
  // noise (removed by request, twice — don't reintroduce).
  return (
    <div>
      {/* Scoped stylesheet; the `.ws-preview` selectors apply to the paper below. */}
      <style>{css}</style>
      <div
        aria-label="Rendered section preview"
        className={previewClasses}
        style={{ fontFamily: bodyFont }}
        dangerouslySetInnerHTML={{ __html: html }}
      />
    </div>
  );
}

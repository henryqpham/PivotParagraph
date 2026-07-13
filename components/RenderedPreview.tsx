import type { HeadingStyle } from "@/lib/clipboard";

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
const previewClasses =
  "ws-preview rounded-[2px] bg-white p-6 text-sm text-[#242424] shadow-[var(--shadow-4)] " +
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
  // The title has its OWN shared look (set in the Section Header group), not the
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
  // Word on paste, exactly like a body heading row — so it gets the same "# "
  // placeholder. The ::before only renders if a `.ws-title` actually exists, so a
  // titleless section shows nothing.
  const titleMapped = (headingStyle.headingStyleName ?? "") !== "";
  const titleRendered = html.includes('class="ws-title"');
  const css =
    titleRule +
    Array.from({ length: 9 }, (_, i) =>
      rule(`[data-level="${i + 1}"]`, i, (i * step).toFixed(2)),
    ).join("") +
    `.ws-preview [data-heading]{margin-left:0;font-weight:700}` +
    `.ws-preview [data-heading]::before{content:"# ";opacity:0.4;font-weight:400}` +
    // A page-break row: a dashed rule + generous space above it stands in for the
    // Word page boundary the paste will insert before this top-level group.
    `.ws-preview [data-break]{margin-top:22px;border-top:1px dashed #b0b0b0;padding-top:16px}` +
    (titleMapped
      ? `.ws-preview .ws-title::before{content:"# ";opacity:0.4;font-weight:400}`
      : "");

  const hasBreaks = html.includes(' data-break="1"');

  // Show the heading footnote whenever a "# " placeholder is visible: a body row
  // mapped to a Word heading, OR a title that is itself mapped to a heading. Match
  // the real emitted attribute (` data-heading="…`), not a bare substring, so a
  // cell/title whose text merely contains "data-heading" can't trigger it.
  const hasHeadings =
    (titleMapped && titleRendered) || html.includes(' data-heading="');

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
      {hasHeadings && (
        <p className="mt-2 text-xs text-muted">
          <span className="opacity-40"># </span>= Word supplies the heading number
          (e.g. 5.0, 5.1) on a <strong>Use Destination Styles</strong> paste, and
          the row joins the Navigation pane. Word applies the destination
          heading&apos;s own formatting too, so it can come out{" "}
          <strong>ALL-CAPS</strong> even though the preview shows normal case. For a
          preview that matches Word exactly, use <strong>Multilevel numbers</strong>{" "}
          instead of a Word heading (the app draws the numbers itself).
        </p>
      )}
      {hasBreaks && (
        <p className="mt-2 text-xs text-muted">
          The dashed rule marks a <strong>Word page break</strong> — each top-level
          group starts on a new page when you paste.
        </p>
      )}
    </div>
  );
}

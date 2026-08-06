"use client";

import { useLayoutEffect, useRef, useState } from "react";
import { headingLevel, type HeadingStyle } from "@/lib/clipboard";

/**
 * Read-only display of the ACTIVE section's rendered pivot HTML (from
 * `renderPivotTree`), presented on a TRUE-TO-SCALE Word page proxy. Each section
 * is its own document now, so this always shows exactly one section — the same
 * fragment `copySection` copies to the clipboard, so preview and export can't
 * diverge.
 *
 * The paper is a real US-Letter aspect (8.5 × 11in): a proportional 1in margin
 * frame (so text wraps in a ~6.5in column, close to how it will flow in Word) with
 * measured, best-effort pagination — when the content is taller than one page's
 * usable band we draw a dashed page boundary and count pages. This is an ESTIMATE,
 * not a guarantee: Word wraps in the destination template's font (usually Calibri),
 * whose metrics differ from this preview's, so a real page break can land a line or
 * two off. It answers "does this roughly fit / spill" far better than a blind
 * paragraph count could, without pretending to be pixel-exact.
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
// US-Letter geometry: the paper is 8.5in wide with 1in margins on every side, so
// each page has a 9in-tall usable content band. Expressed as fractions of the
// paper WIDTH (the one dimension we know in px), since CSS percentage padding is
// itself width-relative — 1/8.5 gives a true 1in margin, and 9/8.5 converts the
// measured content height into "pages".
const MARGIN_FRAC = 1 / 8.5; // 1in side/top/bottom margin as a fraction of width
const PAGE_BAND_FRAC = 9 / 8.5; // 9in usable band per page, ditto

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
  // Measure the paper width + rendered content height to draw page boundaries
  // and count pages, and derive the outline indentation. A LAYOUT effect (runs
  // before paint, so the derived margins never flash their CSS approximation)
  // plus a ResizeObserver for pane resizes and content changes. Hooks run
  // unconditionally, before the empty-state early return below.
  const paperRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<{ pages: number; bandPx: number }>({
    pages: 1,
    bandPx: 0,
  });

  useLayoutEffect(() => {
    const paper = paperRef.current;
    const content = contentRef.current;
    if (!paper || !content) return;
    // Derive the OUTLINE INDENTATION from the rendered geometry, mirroring
    // buildWordHtml's canvas walk: each level's line indents to its PARENT'S
    // TEXT POSITION (parent indent + the parent lead's measured width — the
    // `<span class="ws-lead">` the renderer tags), so "1.1.1" starts exactly
    // under "Name" and stacked/wrapped text lands flush with the first
    // field's text. A lead-less line advances the chain by a fixed 0.25in
    // step (Word's list convention); spacers keep their indent but leave the
    // chain alone; heading rows stay flush-left (Word owns their indent) yet
    // still advance the chain for their children. The static CSS rules are
    // only a no-JS fallback — this runs pre-paint and re-applies on every
    // ResizeObserver tick and after web-font settling, so a late font swap or
    // a zero-width measurement in a hidden pane self-heals on the next tick.
    const STEP_PX = 24; // 0.25in at CSS 96dpi
    const refineOutline = () => {
      const textPos: number[] = [];
      let hangPos = STEP_PX; // carries a hang row's margin to its cont rows
      let hidden = false;
      content.querySelectorAll<HTMLElement>("p[data-level]").forEach((row) => {
        if (hidden) return;
        const d = Number(row.dataset.level) || 1;
        const indent = d <= 1 ? 0 : (textPos[d - 2] ?? (d - 1) * STEP_PX);
        if (row.hasAttribute("data-heading")) {
          textPos[d - 1] = indent + STEP_PX;
          return; // CSS keeps headings flush-left, as Word will
        }
        if (row.hasAttribute("data-cont")) {
          row.style.marginLeft = `${hangPos.toFixed(1)}px`;
          return;
        }
        const lead = row.querySelector<HTMLElement>(".ws-lead");
        if (lead) {
          const w = lead.getBoundingClientRect().width;
          if (w <= 0) {
            hidden = true; // bail wholly — a partial chain would mislead
            return;
          }
          row.style.marginLeft = `${(indent + w).toFixed(1)}px`;
          row.style.textIndent = `-${w.toFixed(1)}px`;
          textPos[d - 1] = hangPos = indent + w;
        } else {
          row.style.marginLeft = `${indent.toFixed(1)}px`;
          if (row.textContent !== "\u00A0")
            textPos[d - 1] = indent + STEP_PX;
        }
      });
    };
    const measure = () => {
      refineOutline();
      const width = paper.clientWidth;
      const bandPx = width * PAGE_BAND_FRAC;
      const contentH = content.scrollHeight;
      const pages = bandPx > 0 ? Math.max(1, Math.ceil(contentH / bandPx)) : 1;
      setLayout((prev) =>
        prev.pages === pages && Math.abs(prev.bandPx - bandPx) < 0.5
          ? prev
          : { pages, bandPx },
      );
    };
    let disposed = false;
    measure();
    // Once web fonts settle, measure again: a lead measured against the
    // fallback font is a few px off after Aptos swaps in.
    document.fonts?.ready
      .then(() => {
        if (!disposed) measure();
      })
      .catch(() => {});
    const ro = new ResizeObserver(measure);
    ro.observe(paper);
    ro.observe(content);
    return () => {
      disposed = true;
      ro.disconnect();
    };
  }, [html, headingStyle, bodyFont]);

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
  const FALLBACK = {
    color: "#000000",
    font: "Aptos",
    size: 11,
    bold: false,
    italic: false,
    underline: false,
  };
  // First-paint / no-JS approximation only — the layout effect above derives
  // the real per-line outline indentation before anything is painted.
  const step = 0.25;
  const rule = (sel: string, i: number, indentIn: string) => {
    const lv = headingStyle.levels[i] ?? FALLBACK;
    const margin = indentIn ? `;margin-left:${indentIn}in` : "";
    return (
      `.ws-preview ${sel}{color:${lv.color};font-family:'${lv.font}';font-size:${lv.size}pt;` +
      `font-weight:${lv.bold ? 700 : 400};font-style:${lv.italic ? "italic" : "normal"};` +
      `text-decoration:${lv.underline ? "underline" : "none"}${margin}}`
    );
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
    font: "Aptos",
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
    // Hanging-alignment approximation for the same no-JS fallback; the
    // layout effect's inline margins win over these rules. After the base
    // level rules, so margin wins.
    Array.from({ length: 9 }, (_, i) => {
      const m = (i * step + 0.25).toFixed(2);
      return (
        `.ws-preview [data-level="${i + 1}"][data-cont]{margin-left:${m}in}` +
        `.ws-preview [data-level="${i + 1}"][data-hang]{margin-left:${m}in;text-indent:-0.25in}`
      );
    }).join("") +
    // Heading-mapped rows are ALWAYS bold + non-italic + non-underline (Word
    // owns their look on a mapping paste; the emitted heading rule declares only
    // bold), so neutralize any per-level italic/underline here to keep the
    // preview matching the paste. Comes AFTER the [data-level] rules to win.
    `.ws-preview [data-heading]{margin-left:0;font-weight:700;font-style:normal;text-decoration:none;position:relative}` +
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

  // Measured page-boundary overlays: a dashed rule + "Page N" flag at each usable-
  // band multiple. Absolutely positioned inside the content box (pointer-events
  // none) so they never affect the measured height. Best-effort — see the file
  // header note on why this is an estimate, not a guarantee.
  const breaks =
    layout.bandPx > 0 && layout.pages > 1
      ? Array.from({ length: layout.pages - 1 }, (_, i) => (i + 1) * layout.bandPx)
      : [];

  return (
    <div className="flex flex-col gap-2">
      {/* Honest paste-mode label. (A page-fit status chip lived beside it and
          was removed as noise — the ~N-pages stat next to the view tabs and the
          drawn page boundaries below already answer "does it fit".) */}
      <div className="flex flex-wrap items-center justify-end gap-2 px-0.5 text-xs">
        <span
          className="text-muted"
          title="The preview shows the look a Keep-Source-Formatting paste produces. A Use-Destination-Styles paste instead adopts your Word template's own Heading styles for mapped headings."
        >
          Showing: your <strong className="font-semibold">Keep-Source</strong> look
        </span>
      </div>

      {/* Scoped stylesheet; the `.ws-preview` selectors apply to the paper below. */}
      <style>{css}</style>

      {/* The Word-page proxy: white in both themes, real US-Letter aspect via a
          proportional 1in margin frame, floating on the canvas with a depth shadow. */}
      <div
        ref={paperRef}
        aria-label="Rendered section preview on a Word page"
        className="relative rounded-[2px] bg-white shadow-[var(--shadow-4)]"
        style={{
          paddingLeft: `${MARGIN_FRAC * 100}%`,
          paddingRight: `${MARGIN_FRAC * 100}%`,
          paddingTop: `${MARGIN_FRAC * 100}%`,
          paddingBottom: `${MARGIN_FRAC * 100}%`,
        }}
      >
        <div ref={contentRef} className="relative">
          {/* Measured page boundaries (best-effort). */}
          {breaks.map((top, i) => (
            <div
              key={i}
              aria-hidden
              className="pointer-events-none absolute inset-x-0 flex items-center"
              style={{ top }}
            >
              <span className="h-px flex-1 border-t border-dashed border-[#c0c0c0]" />
              <span className="ml-1 rounded-sm bg-[#f0f0f0] px-1 py-0.5 text-[9px] font-semibold uppercase tracking-wide text-[#888]">
                Page {i + 2}
              </span>
            </div>
          ))}
          <div
            className={
              "ws-preview text-sm text-[#242424] " +
              "[&_.ws-title]:mt-1 [&_[data-level]]:mt-1 [&_p]:my-0.5 [&_p]:leading-tight"
            }
            style={{ fontFamily: bodyFont }}
            dangerouslySetInnerHTML={{ __html: html }}
          />
        </div>
      </div>
    </div>
  );
}

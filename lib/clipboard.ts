// Word-output step: wrap the renderer's HTML fragment as Word-flavored HTML for
// the clipboard ("Copy section"). The title and any heading-marked body levels map
// to destination Heading styles as real <h1>-<h6> elements whose style rules
// declare the full source look (see buildWordHtml for the verified paste-mode
// mechanics); every other body paragraph is plain, directly-formatted text (any
// multilevel numbers are already plain text in the fragment). renderPivotTree
// returns a bare fragment (no <html>/<head>/<body>); the browser's ClipboardItem
// writes the Windows CF_HTML header for us.

/**
 * How one heading level should look -- one entry per depth (index 0 = level 1).
 * Shared across all tables; defaults are all the same until the user diverges a
 * level.
 */
export type LevelStyle = {
  color: string; // hex
  font: string; // font-family name
  size: number; // pt
  bold: boolean;
};

/**
 * The title's own direct look (font/size/color + bold/italic/underline), SHARED
 * across tables. Drives the `ws-title` row's APPEARANCE in BOTH cases: with no
 * heading it styles a plain paragraph; with a heading mapped the SAME look is
 * emitted as inline direct formatting layered on the heading, so the Look always
 * wins the appearance while Word still supplies the heading's number/outline/
 * spacing (and any all-caps). Kept separate from `levels` so the title is styled in
 * its own dedicated "Section Header" group, not buried in the per-level "level
 * chart".
 */
export type TitleStyle = {
  font: string;
  size: number; // pt
  color: string; // hex
  bold: boolean;
  italic: boolean;
  underline: boolean;
};

/**
 * The single styling source, shared across all tables.
 * - `levels` -- per-depth direct look (index 0 = level 1). Styles the plain body
 *   rows inline, and is the look a heading-marked level's style rule declares.
 * - `indentStep` -- left-indent added per nesting level (inches).
 * - `headingStyleName` -- optional Word style name for the TITLE ("" = None). A
 *   built-in "Heading 1-6" (all the dropdown offers) is emitted as a real `<hN>`
 *   element so the paste maps it to the destination document's heading style
 *   (see buildWordHtml for the verified paste-mode mechanics); any other name
 *   falls back to the `p.MsoTitle` + `mso-style-name` class route. Blank = the
 *   title's own direct `titleStyle` look. Body levels the user marks as Word
 *   headings map the same way (`data-heading` -> `<hK>`); every other body row
 *   is the app's direct inline look.
 */
export type HeadingStyle = {
  levels: LevelStyle[];
  indentStep: number; // inches of left-indent per nesting level
  headingStyleName: string; // Word style for the title ("" = direct look)
  titleStyle?: TitleStyle; // the title's own direct look (when not mapped)
  /** Body line spacing as a multiplier (1 / 1.15 / 1.5; default 1.15). Emitted
   *  as `line-height:N%` — Word reads that as "Multiple" paragraph spacing. */
  lineSpacing?: number;
};

/**
 * Make a string safe inside the `mso-style-name:"..."` declaration in the
 * clipboard `<style>`. A real Word style name needs only letters, digits, spaces,
 * dots, underscores, and hyphens (template styles like "TBL_TITLE" use `_`), so
 * we ALLOW-LIST that set and drop everything else. That removes every character
 * with CSS meaning -- the closing quote, `{` `}` `;`, backslash, angle brackets,
 * comment sequences, and newlines/control chars -- so a crafted name can't
 * terminate the string and inject a rule into the document copied to the
 * clipboard. Names like "Heading 1" still pass unchanged.
 */
function sanitizeStyleName(s: string): string {
  return s.replace(/[^A-Za-z0-9 ._-]/g, "").trim();
}

/**
 * The Word heading LEVEL the title maps to: the trailing number of a "Heading N"
 * name (1-9), or 1 for any other non-empty style name, or 0 when no heading is set
 * (None). Threaded through the render pipeline so BOTH the title's own
 * `mso-outline-level` AND the body-heading offset (a body level nests one under the
 * title's level: a `Heading 3` title => first body heading `Heading 4`) follow the
 * ACTUAL chosen heading instead of a hardcoded 1.
 */
export function headingLevel(name: string): number {
  const clean = sanitizeStyleName(name);
  if (clean === "") return 0;
  const m = clean.match(/(\d+)/);
  return m ? Math.min(Math.max(parseInt(m[1], 10), 1), 9) : 1;
}

// Guard if a level entry is ever absent (callers pass a full 9-entry array).
const FALLBACK_LEVEL: LevelStyle = {
  color: "#000000",
  font: "Arial",
  size: 11,
  bold: false,
};

// Title default when no `titleStyle` is supplied (matches the old level-1 look).
const FALLBACK_TITLE: TitleStyle = {
  font: "Arial",
  size: 11,
  color: "#000000",
  bold: false,
  italic: false,
  underline: false,
};

/**
 * Wrap a rendered pivot fragment as a minimal Word-flavored HTML document for the
 * clipboard.
 *
 * HEADING MAPPING — real `<h1>`–`<h6>` ELEMENTS whose `<style>` rule declares the
 * FULL source look (verified empirically against desktop Word via COM paste
 * tests, July 2026). The mechanics that matter:
 * - A "Use Destination Styles" (or default Ctrl+V) paste maps `<hK>` to the
 *   built-in `Heading K` in a blank doc AND a template (where an alias like
 *   "1.0 HEADING" still resolves). Everything DECLARED IN THE `hK` CSS RULE is
 *   treated as the SOURCE STYLE'S DEFINITION and is swapped wholesale for the
 *   destination style — clean adoption, no direct-formatting residue. Anything
 *   NOT declared in the rule falls back to Word's HTML-reader h-defaults and
 *   rides along as DIRECT formatting that pollutes the destination heading
 *   (a bare `<h1>` pastes as Heading 1 + direct bold 24pt). So each emitted
 *   `<hK>` gets a rule declaring `mso-style-name`/`mso-outline-level` PLUS the
 *   level's font/size/color and bold — making the paste adopt the destination
 *   heading exactly.
 * - A "Keep Source Formatting" paste does NOT map named styles at all (the row
 *   lands as Normal + `mso-outline-level` direct formatting — it still shows in
 *   the Navigation pane, which deceptively half-works). There the declared rule
 *   look IS the visible formatting, so declaring the preview's heading look
 *   (level font/size/color, bold, flush-left) makes a KSF paste read as a real
 *   bold heading instead of naked body text.
 * - "Merge Formatting" strips both the style mapping and the outline level;
 *   nothing can survive it. The copy-confirmation banner steers users to the
 *   other two options.
 *
 * The TITLE maps this way too when `headingStyleName` is a built-in "Heading 1–6"
 * (all the dropdown offers): it becomes `<hN>` so the paste puts it in the
 * document outline and Word supplies its heading NUMBER, while the user's full
 * title Look (font/size/color + bold/italic/underline) is emitted as inline
 * direct formatting so the appearance pastes exactly as previewed. Any OTHER
 * non-empty style name keeps the legacy `<p class="MsoTitle">` +
 * `mso-style-name:"<name>"` route (a custom name has no tag equivalent; it maps
 * on a Use-Destination-Styles paste); blank = the same direct look on a plain
 * paragraph.
 *
 * A body level the user mapped to a Word heading arrives as `data-heading="K"` and
 * becomes `<hK>` (K ≤ 6), so the paste adds just those rows to the document
 * outline (nav pane + collapsible) and Word supplies their number. K 7–9 (no HTML
 * tag exists) keep `<p class="MsoHeadingK">` with the same rule declarations —
 * the class route maps on the same Use-Destination-Styles/default pastes (and,
 * like the tag route, never under Keep Source Formatting).
 *
 * Every OTHER body paragraph uses the app's direct per-level look
 * (color/font/size/bold) + indent + compact spacing, emitted as INLINE formatting
 * on each `<p>` (and `<b>`/`<i>`/`<u>` runs for the label) rather than CSS classes. A
 * "Use Destination Styles" paste discards class/style-name formatting it can't map
 * but KEEPS inline direct formatting + inline runs, so the look, tight spacing,
 * and label bold/italic/underline survive and match the live preview.
 */
export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
  const step = heading.indentStep ?? 0.2;
  const headingName = sanitizeStyleName(heading.headingStyleName ?? "");
  const titleLevel = headingLevel(headingName); // outline level for the title
  // Body line spacing (percent). Clamped to a sane band; 115 = the old fixed value.
  const spacingRaw = heading.lineSpacing ?? 1.15;
  const spacingPct = Math.round(
    Math.min(3, Math.max(1, Number.isFinite(spacingRaw) ? spacingRaw : 1.15)) *
      100,
  );

  // Inline direct formatting for an app-styled paragraph: level index `i`, indented
  // for render-level `n` (1-based). Inline (not a CSS class) so it survives a
  // "Use Destination Styles" paste. Compact spacing = no space before/after +
  // a 1.15 line, so paragraphs sit tight like the preview.
  const directStyle = (i: number, n: number) => {
    const s = lvl(i);
    const indent = ((n - 1) * step).toFixed(2);
    return (
      `margin-top:0in;margin-bottom:0in;margin-left:${indent}in;line-height:${spacingPct}%;` +
      `color:${s.color};font-family:'${s.font}';font-size:${s.size}pt`
    );
  };
  // Bold is a <b> run (direct character formatting), which a "Use Destination
  // Styles" paste keeps -- unlike a class-level font-weight, which it discards.
  const wrapBold = (i: number, content: string) =>
    lvl(i).bold ? `<b>${content}</b>` : content;

  // The title's own direct look + bold/italic/underline runs, used for the
  // `ws-title` row when it isn't mapped to a Word heading. Underline innermost,
  // then italic, then bold (all <b>/<i>/<u> runs survive a Use-Destination paste).
  const ts = heading.titleStyle ?? FALLBACK_TITLE;
  const titleStyleAttr =
    `margin-top:0in;margin-bottom:0in;line-height:${spacingPct}%;` +
    `color:${ts.color};font-family:'${ts.font}';font-size:${ts.size}pt`;
  const wrapTitle = (content: string) => {
    let h = content;
    if (ts.underline) h = `<u>${h}</u>`;
    if (ts.italic) h = `<i>${h}</i>`;
    if (ts.bold) h = `<b>${h}</b>`;
    return h;
  };

  // When the title IS mapped to a Word heading, keep the heading class + mso rule
  // (so it enters the Word outline, gets Word's heading NUMBER, and its own
  // paragraph spacing) but make the user's full Look win the APPEARANCE: emit
  // font/size/color and a font-weight as one inline <span> run, wrapped by <i>/<u>
  // runs. Inline direct character formatting overrides the heading style's own
  // character look and survives a "Use Destination Styles" paste, so the title
  // pastes as previewed. (Two things stay Word's while mapped and can't be
  // overridden from HTML: the heading's auto-NUMBER and an all-caps heading's
  // UPPERCASE effect.)
  //
  // BOLD is special: Word treats character bold as a TOGGLE relative to the
  // paragraph STYLE's own weight, and the built-in Heading 1-9 styles are all bold.
  // So a `font-weight:bold` run on a mapped (bold) heading CANCELS to not-bold,
  // while `font-weight:normal` is ignored and the heading's bold shows through. To
  // make the B button read ABSOLUTELY (checked => bold, unchecked => not bold) we
  // therefore INVERT the weight we emit: inherit (normal) to KEEP the heading's
  // bold, and emit an explicit bold run only to CANCEL it. This assumes the mapped
  // heading is bold (true for Word's built-in Heading 1-4); a custom NON-bold
  // heading style would flip it -- use Heading = None there for absolute control.
  const wrapTitleOnHeading = (content: string) => {
    const weight = ts.bold ? "normal" : "bold"; // inverted: cancels Word's toggle
    const runStyle =
      `font-family:'${ts.font}';font-size:${ts.size}pt;` +
      `color:${ts.color};font-weight:${weight}`;
    let h = `<span style="${runStyle}">${content}</span>`;
    if (ts.underline) h = `<u>${h}</u>`;
    if (ts.italic) h = `<i>${h}</i>`;
    return h;
  };

  // The title maps to a real <hN> element only for a built-in "Heading 1"–"Heading
  // 6" name (all the dropdown offers); a real tag is what maps reliably in every
  // paste mode and document (see the function comment). Any other non-empty name
  // has no tag equivalent and keeps the p.MsoTitle route.
  const titleTagLevel = /^heading [1-6]$/i.test(headingName) ? titleLevel : 0;
  // The look a heading rule declares as its "source style definition": character
  // props only (color/font/size + bold). Declared props are REPLACED by the
  // destination style on a Use-Destination-Styles/default paste (clean adoption —
  // undeclared props would instead ride along as polluting direct formatting) and
  // ARE the visible look on a Keep-Source-Formatting paste (matching the preview's
  // bold heading rows). No paragraph props: Word supplies the destination
  // heading's spacing when mapped, its h-defaults under KSF.
  type HeadingLook = { color: string; font: string; size: number };
  // Heading levels emitted as real <hN> tags (title + body) and body levels 7–9
  // (no HTML tag exists; class route). Each K records the look its rule declares.
  const tagHeadings = new Map<number, HeadingLook>();
  const classHeadings = new Map<number, HeadingLook>();
  let titleUsesClassRoute = false; // saw a title mapped to a NON-built-in name
  const body = fragment
    // Title: a mapped Word heading when named (<hN> for a built-in Heading 1–6,
    // else the legacy class + mso rule), else the app's direct look inline.
    .replace(/<p class="ws-title">([\s\S]*?)<\/p>/g, (_m, content: string) => {
      if (!headingName)
        return `<p style="${titleStyleAttr}">${wrapTitle(content)}</p>`;
      if (titleTagLevel) {
        // The rule declares the title's own look; the inline span in
        // wrapTitleOnHeading overrides the text run either way, so this only
        // shapes the paragraph fallback under KSF.
        tagHeadings.set(titleTagLevel, {
          color: ts.color,
          font: ts.font,
          size: ts.size,
        });
        return `<h${titleTagLevel}>${wrapTitleOnHeading(content)}</h${titleTagLevel}>`;
      }
      titleUsesClassRoute = true;
      return `<p class="MsoTitle">${wrapTitleOnHeading(content)}</p>`;
    })
    // Nested rows. A `data-heading="K"` line is a level mapped to a Word heading:
    // emit it as a real <hK> (K ≤ 6) so the paste maps it to the destination
    // "Heading K" style — blank doc or template — adding it to the outline (nav +
    // collapsible) with Word supplying its number; K 7–9 keep the class + mso rule
    // (no tag exists that deep). Its rule declares the LEVEL's look (the same look
    // the preview shows for a heading row), keyed by the row's data-level.
    // Otherwise the app's per-level look, inline (+ <b> when bold). An optional
    // `data-break` attribute (a top-level group starting a new page) adds
    // `page-break-before:always` to whichever paragraph it lands on; the spacer
    // paragraph (nbsp, data-level only) passes through the else branch as an
    // ordinary inline body paragraph.
    .replace(
      /<p class="ws-lvl" data-level="([1-9])"( data-break="1")?(?: data-heading="([1-9])")?>([\s\S]*?)<\/p>/g,
      (
        _m,
        d: string,
        brk: string | undefined,
        h: string | undefined,
        content: string,
      ) => {
        const pageBreak = brk ? "page-break-before:always" : "";
        const i = Number(d) - 1;
        if (h) {
          const k = Number(h);
          const s = lvl(i);
          const look = { color: s.color, font: s.font, size: s.size };
          // The page break is added as inline direct formatting (survives a
          // Use-Destination-Styles paste).
          const style = pageBreak ? ` style="${pageBreak}"` : "";
          if (k <= 6) {
            if (!tagHeadings.has(k)) tagHeadings.set(k, look);
            return `<h${k}${style}>${content}</h${k}>`;
          }
          if (!classHeadings.has(k)) classHeadings.set(k, look);
          return `<p class="MsoHeading${k}"${style}>${content}</p>`;
        }
        const base = directStyle(i, Number(d));
        const style = pageBreak ? `${base};${pageBreak}` : base;
        return `<p style="${style}">${wrapBold(i, content)}</p>`;
      },
    );

  // Mapped-style rules: one per heading level actually used. Each declares the
  // style identity (mso-style-name + mso-outline-level) AND the full source look
  // (color/font/size + bold — bold because the preview shows heading rows bold),
  // so a mapping paste adopts the destination heading cleanly and a KSF paste
  // shows a real heading look (see the function comment). The class rules
  // (custom-named title, Heading 7–9) are the only mapping those routes have.
  // Emitted only for what the fragment actually used, so e.g. a blank title adds
  // no stray rule. Every plain body level is inline and needs no rule.
  const headingRule = (sel: string, k: number, s: HeadingLook) =>
    `${sel}{mso-style-name:"Heading ${k}";mso-outline-level:${k};` +
    `color:${s.color};font-family:'${s.font}';font-size:${s.size}pt;font-weight:bold}`;
  const titleRule = titleUsesClassRoute
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:${titleLevel}}`
    : "";
  const headingRules =
    Array.from(tagHeadings)
      .map(([k, s]) => headingRule(`h${k}`, k, s))
      .join("") +
    Array.from(classHeadings)
      .map(([k, s]) => headingRule(`p.MsoHeading${k}`, k, s))
      .join("");
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w="urn:schemas-microsoft-com:office:word" ` +
    `xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8">` +
    // The ProgId/Generator/Originator metas are what Word's own clipboard writer
    // emits; they mark the payload as WORD-DOCUMENT content rather than generic
    // web HTML, so Word offers its between-documents paste options (incl. "Use
    // Destination Styles") instead of the other-program set — without them a
    // blank destination (no conflicting in-use styles) may only offer Keep
    // Source / Merge.
    `<meta name="ProgId" content="Word.Document">` +
    `<meta name="Generator" content="Microsoft Word 15">` +
    `<meta name="Originator" content="Microsoft Word 15">` +
    `<style>` +
    `@page{size:8.5in 11in;margin:1in}` +
    // Body font fallback (so Word doesn't drop to Times New Roman); each paragraph
    // also sets its own font inline.
    `body{overflow-wrap:break-word;font-family:"${bodyFont}"}` +
    titleRule +
    headingRules +
    `</style>` +
    `</head><body>${body}</body></html>`
  );
}

/**
 * Readable plain-text fallback for the `text/plain` clipboard flavor.
 *
 * Operates on the machine-generated fragment shape (all `<p>` plus escaped
 * `& < >`). The `&#160;` nbsp emitted by "blank line after" spacers is dropped
 * first (so a spacer becomes an empty line, collapsed by the `\n{2,}` pass, not a
 * literal "&#160;"); it can't match a user's literal "&#160;" text because that
 * arrives escaped as `&amp;#160;`. Then lt/gt are unescaped BEFORE amp -- the
 * inverse of the renderer's "amp first" escape order -- so `&amp;lt;` decodes to
 * `&lt;`, never `<`.
 */
export function htmlToPlainText(fragment: string): string {
  return fragment
    .replace(/<\/p>/g, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&#160;/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

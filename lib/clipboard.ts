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
  italic: boolean;
  underline: boolean;
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
 * - `indentStep` -- LEGACY; ignored (indentation is derived per line).
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
  /** LEGACY -- ignored. Indentation is DERIVED per line from the outline
   *  geometry (each level nests under its parent's text; see buildWordHtml);
   *  kept optional so old snapshots/presets still type-check. */
  indentStep?: number;
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
  font: "Aptos",
  size: 11,
  bold: false,
  italic: false,
  underline: false,
};

// Title default when no `titleStyle` is supplied (matches the old level-1 look).
const FALLBACK_TITLE: TitleStyle = {
  font: "Aptos",
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
 * document outline and Word supplies its heading NUMBER. The title's
 * font/size/color (+ `<i>`/`<u>` runs) ride INLINE and its WEIGHT is declared in
 * the title's style RULE — COM-verified result: a KEEP-SOURCE-FORMATTING paste
 * shows exactly the preview look (inline + rule), while a USE-DESTINATION-STYLES
 * paste strips BOTH and gives the destination heading's complete look (font,
 * size, weight — Word's style recovery discards direct formatting on
 * style-mapped paragraphs). No inline weight is emitted at all: bold is an OOXML
 * toggle, and the old inverted-weight trick rendered backwards under KSF (the
 * bug this replaced). Same convention as the body heading rows: UDS = the
 * document's look, KSF = the preview's.
 * Any OTHER non-empty style name keeps the legacy `<p class="MsoTitle">` +
 * `mso-style-name:"<name>"` route with the same rule-declared weight (a custom
 * name has no tag equivalent; it maps on a Use-Destination-Styles paste);
 * blank = the full direct look, including a literal `<b>` run, on a plain
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
 * (color/font/size/bold/italic/underline) + indent + compact spacing, emitted as
 * INLINE formatting on each `<p>` (with `<b>`/`<i>`/`<u>` runs for the level look
 * AND for the label) rather than CSS classes. A
 * "Use Destination Styles" paste discards class/style-name formatting it can't map
 * but KEEPS inline direct formatting + inline runs, so the look, tight spacing,
 * and label bold/italic/underline survive and match the live preview.
 */
/**
 * Width (in INCHES) of a hang line's lead ("1.1.1.1 " / "a) ") at a given
 * font/size/weight. In the browser (every real copy runs there) a cached canvas
 * measures the exact advance width — the same font + pt size Word lays out, so
 * the hanging indent lands the continuation text flush under the first line's
 * text. Headless (tests/SSR) falls back to a per-glyph estimate (digits/letters
 * ~0.52em, space ~0.23em, punctuation ~0.27em).
 */
let leadCtx: CanvasRenderingContext2D | null | undefined;
function measureLeadIn(
  text: string,
  font: string,
  sizePt: number,
  bold: boolean,
): number {
  if (leadCtx === undefined)
    leadCtx =
      typeof document !== "undefined"
        ? (document.createElement("canvas").getContext("2d") ?? null)
        : null;
  if (leadCtx) {
    leadCtx.font = `${bold ? "bold " : ""}${sizePt}pt '${font}', 'Segoe UI', sans-serif`;
    return leadCtx.measureText(text).width / 96;
  }
  let em = 0;
  for (const ch of text)
    em += /[0-9A-Za-z\u2022\u2013]/.test(ch) ? 0.52 : ch === " " ? 0.23 : 0.27;
  return (em * sizePt) / 72;
}

export function buildWordHtml(
  fragment: string,
  heading: HeadingStyle,
  bodyFont: string,
): string {
  const lvl = (i: number) => heading.levels[i] ?? FALLBACK_LEVEL;
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
  const directStyle = (i: number) => {
    const s = lvl(i);
    // margin-left is a placeholder — the outline walk below replaces it with
    // this line's derived indent.
    return (
      `margin-top:0in;margin-bottom:0in;margin-left:0in;line-height:${spacingPct}%;` +
      `color:${s.color};font-family:'${s.font}';font-size:${s.size}pt`
    );
  };
  // Outline geometry ("legal numbering", the way Word's own multilevel lists
  // nest): each level's line indents to its PARENT'S TEXT POSITION — the
  // parent's indent plus the MEASURED width of the parent's lead — so "1.1.1"
  // starts exactly under "Name", however wide the numbers grow. A line with a
  // lead hangs it (margin = indent + lead width, first-line indent = -width),
  // so wrapped text and stacked continuation lines land flush with the text.
  // textPosIn[d-1] = inches where the most recent depth-d line's text begins;
  // document order guarantees a parent is processed before its children, so
  // reading [d-2] always sees the current branch. STEP_IN is the advance for a
  // lead-less line (numbering off / hidden number / heading) — Word's 0.25in
  // list convention. The fixed ⚙ "Indent/level" control was REMOVED for this.
  const STEP_IN = 0.25;
  const textPosIn: number[] = [];
  const indentFor = (d: number) =>
    d <= 1 ? 0 : (textPosIn[d - 2] ?? (d - 1) * STEP_IN);
  let hangW = STEP_IN; // carries a hang row's lead width to its cont rows
  // The per-level look's bold/italic/underline as nested <b>/<i>/<u> runs (direct
  // character formatting, which a "Use Destination Styles" paste keeps -- unlike a
  // class-level font-weight it would discard). Underline innermost then italic then
  // bold, matching the title's wrap order and the preview.
  const wrapLook = (i: number, content: string) => {
    const s = lvl(i);
    let h = content;
    if (s.underline) h = `<u>${h}</u>`;
    if (s.italic) h = `<i>${h}</i>`;
    if (s.bold) h = `<b>${h}</b>`;
    return h;
  };

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
  // WEIGHT is deliberately NOT in this inline run. Bold is an OOXML toggle
  // property, and the old inverted-weight trick rendered LITERALLY (backwards)
  // under Keep Source Formatting, where no style mapping happens. The title's
  // true weight lives in its STYLE RULE instead (below). COM-verified outcome:
  // KSF = exact preview look (this inline run + the rule); UDS = the destination
  // heading's complete look (Word's style recovery strips direct formatting on
  // mapped paragraphs — the inline font/size/color only survive non-mapping
  // pastes, which is exactly where they're needed).
  const wrapTitleOnHeading = (content: string) => {
    const runStyle =
      `font-family:'${ts.font}';font-size:${ts.size}pt;color:${ts.color}`;
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
  type HeadingLook = { color: string; font: string; size: number; bold: boolean };
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
        // The rule declares the title's own look INCLUDING its true weight (the
        // inline span carries font/size/color but deliberately no weight — see
        // wrapTitleOnHeading): KSF shows this rule look, UDS swaps it for the
        // destination heading's definition.
        tagHeadings.set(titleTagLevel, {
          color: ts.color,
          font: ts.font,
          size: ts.size,
          bold: ts.bold,
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
      /<p class="ws-lvl" data-level="([1-9])"( data-break="1")?(?: data-heading="([1-9])")?( data-hang="1"| data-cont="1")?>([\s\S]*?)<\/p>/g,
      (
        _m,
        d: string,
        brk: string | undefined,
        h: string | undefined,
        align: string | undefined,
        content: string,
      ) => {
        const pageBreak = brk ? "page-break-before:always" : "";
        const i = Number(d) - 1;
        if (h) {
          // A heading renders flush-left (Word owns its indent on paste) but
          // still occupies its slot in the outline chain: children indent one
          // step past where its body line would have sat.
          textPosIn[Number(d) - 1] = indentFor(Number(d)) + STEP_IN;
          const k = Number(h);
          const s = lvl(i);
          // Heading rows always declare bold (the preview shows them bold).
          const look = { color: s.color, font: s.font, size: s.size, bold: true };
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
        const indent = indentFor(Number(d));
        let base = directStyle(i);
        // Outline indentation + hanging alignment, all inline direct
        // formatting (survives any paste mode): a lead line hangs its measured
        // lead; a plain line sits at the derived indent and advances the chain
        // by STEP_IN; a spacer keeps the indent but leaves the chain alone.
        let body = content;
        if (align === ' data-hang="1"') {
          // The tagged lead (`ws-lead`, stripped from the output) is measured
          // in the font it actually renders in: a field-look override's span
          // wraps the whole line, else the level style. Bold widens ~4%, so the
          // level look, a bold field look, and a bolded label all count.
          const leadHtml =
            /<span class="ws-lead">([\s\S]*?)<\/span>/.exec(content)?.[1] ?? "";
          const leadText = leadHtml
            .replace(/<[^>]+>/g, "")
            .replace(/&#160;/g, "\u00A0")
            .replace(/&lt;/g, "<")
            .replace(/&gt;/g, ">")
            .replace(/&amp;/g, "&");
          const lookStyle =
            /^(?:<[biu]>)*<span style="([^"]*)"/.exec(content)?.[1] ?? "";
          const family =
            /font-family:'([^']+)'/.exec(lookStyle)?.[1] ?? lvl(i).font;
          const sizePt =
            Number(/font-size:([\d.]+)pt/.exec(lookStyle)?.[1]) || lvl(i).size;
          const bold =
            lvl(i).bold ||
            /^<b>/.test(content) ||
            /<span class="ws-lead"><b>/.test(content);
          hangW = leadText
            ? Math.min(2, Math.max(0.05, measureLeadIn(leadText, family, sizePt, bold)))
            : STEP_IN;
          textPosIn[Number(d) - 1] = indent + hangW;
          base = `${base.replace(/margin-left:[^;]+/, `margin-left:${(indent + hangW).toFixed(3)}in`)};text-indent:-${hangW.toFixed(3)}in`;
          body = content.replace(/<span class="ws-lead">([\s\S]*?)<\/span>/, "$1");
        } else if (align === ' data-cont="1"')
          base = base.replace(
            /margin-left:[^;]+/,
            `margin-left:${(indent + hangW).toFixed(3)}in`,
          );
        else {
          base = base.replace(
            /margin-left:[^;]+/,
            `margin-left:${indent.toFixed(3)}in`,
          );
          // Spacer paragraphs (bare &#160;) keep the indent but must not
          // advance the chain — a gap between stacked lines would otherwise
          // clobber the group's text position before its children render.
          if (content !== "&#160;")
            textPosIn[Number(d) - 1] = indent + STEP_IN;
        }
        const style = pageBreak ? `${base};${pageBreak}` : base;
        return `<p style="${style}">${wrapLook(i, body)}</p>`;
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
    `color:${s.color};font-family:'${s.font}';font-size:${s.size}pt;` +
    `font-weight:${s.bold ? "bold" : "normal"}}`;
  // The custom-name route's rule also declares the title's true weight, for the
  // same reason as the <hN> rules: KSF renders it (preview-faithful), UDS swaps
  // it for the destination style's definition.
  const titleRule = titleUsesClassRoute
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:${titleLevel};` +
      `font-weight:${ts.bold ? "bold" : "normal"}}`
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

/**
 * Put a rich (`text/html` + `text/plain`) payload on the clipboard, returning
 * whether it landed. Tries the modern async Clipboard API first, then falls back
 * to the legacy `execCommand("copy")` + copy-event route.
 *
 * The fallback is what makes the SINGLE-FILE build viable: the async API needs a
 * secure context and a clipboard permission, and a page opened by double-clicking
 * a `.html` (origin `null`, no HTTPS) is exactly the case where a browser or a
 * managed-browser policy is most likely to refuse it. The legacy path has no such
 * requirement — it only needs to run inside a user gesture, which "Copy section"
 * always is — and it is still the standard way to place HTML (not just text) on
 * the clipboard. Copy is the app's entire payoff, so it must not depend on one API.
 *
 * MUST be called synchronously from a user gesture (a click / keydown handler).
 */
export async function writeRichClipboard(
  html: string,
  plain: string,
): Promise<boolean> {
  if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([plain], { type: "text/plain" }),
        }),
      ]);
      return true;
    } catch {
      // Blocked (insecure context, denied permission, unsupported) — fall through.
    }
  }
  return legacyRichCopy(html, plain);
}

/**
 * `execCommand("copy")` with a one-shot `copy` listener that overwrites the
 * payload — the pre-async-API way to put real HTML on the clipboard. Needs a live
 * selection to copy FROM, so it briefly selects an off-screen node (positioned
 * off-canvas rather than `display:none`, which cannot be selected) and restores
 * the user's own selection afterward.
 */
function legacyRichCopy(html: string, plain: string): boolean {
  if (typeof document === "undefined") return false;
  const onCopy = (e: ClipboardEvent) => {
    e.preventDefault();
    e.clipboardData?.setData("text/html", html);
    e.clipboardData?.setData("text/plain", plain);
  };
  // Capture phase so this wins regardless of anything else listening for copy.
  document.addEventListener("copy", onCopy, true);
  const holder = document.createElement("div");
  const selection = window.getSelection();
  const previous =
    selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
  try {
    holder.textContent = plain || " ";
    holder.setAttribute(
      "style",
      "position:fixed;left:-9999px;top:0;white-space:pre;opacity:0",
    );
    document.body.appendChild(holder);
    const range = document.createRange();
    range.selectNodeContents(holder);
    selection?.removeAllRanges();
    selection?.addRange(range);
    return document.execCommand("copy");
  } catch {
    return false;
  } finally {
    document.removeEventListener("copy", onCopy, true);
    holder.remove();
    // Leave the page as we found it: drop our scratch selection, and put the
    // user's own back if they had one.
    selection?.removeAllRanges();
    if (previous) selection?.addRange(previous);
  }
}

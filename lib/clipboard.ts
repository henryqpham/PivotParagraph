// Word-output step: wrap the renderer's HTML fragment as Word-flavored HTML for
// the clipboard ("Copy for Word"). The title maps to a destination Heading style
// when a Heading style is named (so a "Use Destination Styles" paste adopts the
// document's heading); every body paragraph is plain, directly-formatted text
// (any multilevel numbers are already plain text in the fragment). renderPivotTree
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
 * The single styling source, shared across all tables.
 * - `levels` -- per-depth direct look (index 0 = level 1). Always used for the
 *   body (nested rows), and for the title when `headingStyleName` is blank.
 * - `indentStep` -- left-indent added per nesting level (inches).
 * - `headingStyleName` -- optional Word style name for the TITLE only. When set,
 *   the title gets `mso-style-name:"<headingStyleName>"`, so a "Use Destination
 *   Styles" paste adopts the destination document's heading style. Blank =
 *   the app's direct level-1 look. The body is always the app's direct look.
 */
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

export type HeadingStyle = {
  levels: LevelStyle[];
  indentStep: number; // inches of left-indent per nesting level
  headingStyleName: string; // Word style for the title ("" = direct look)
  titleStyle?: TitleStyle; // the title's own direct look (when not mapped)
};

/**
 * Make a string safe inside the `mso-style-name:"..."` declaration in the
 * clipboard `<style>`. A real Word style name needs only letters, digits, spaces,
 * dots, and hyphens, so we ALLOW-LIST that set and drop everything else. That
 * removes every character with CSS meaning -- the closing quote, `{` `}` `;`,
 * backslash, angle brackets, comment sequences, and newlines/control chars -- so a
 * crafted name can't terminate the string and inject a rule into the document
 * copied to the clipboard. Names like "Heading 1" still pass unchanged.
 */
function sanitizeStyleName(s: string): string {
  return s.replace(/[^A-Za-z0-9 .-]/g, "").trim();
}

/**
 * Whether a Heading style is configured for the title (a non-empty name after
 * sanitizing). A body level mapped to a Word heading nests one level under the
 * title's `Heading 1`, so callers pass this the same boolean the title mapping
 * branches on (it sets the body heading level: title Heading 1 → body Heading 2).
 */
export function isHeadingStyleSet(name: string): boolean {
  return sanitizeStyleName(name) !== "";
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
 * The TITLE maps to a Word heading style when `headingStyleName` is set (it
 * carries `mso-style-name:"<name>";mso-outline-level:<N>` for the chosen Heading N,
 * so a "Use Destination Styles" paste puts it in the document outline and Word
 * supplies its heading NUMBER), while the user's full title Look (font/size/color +
 * bold/italic/underline) is emitted as inline direct formatting so the appearance
 * pastes exactly as previewed; blank = the same direct look on a plain paragraph.
 *
 * A body level the user mapped to a Word heading arrives as `data-heading="K"` and
 * becomes `<p class="MsoHeadingK">` + a `mso-style-name:"Heading K"` rule, so a
 * "Use Destination Styles" paste adds just those rows to the document outline (nav
 * pane + collapsible) and Word supplies their number.
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

  // Inline direct formatting for an app-styled paragraph: level index `i`, indented
  // for render-level `n` (1-based). Inline (not a CSS class) so it survives a
  // "Use Destination Styles" paste. Compact spacing = no space before/after +
  // a 1.15 line, so paragraphs sit tight like the preview.
  const directStyle = (i: number, n: number) => {
    const s = lvl(i);
    const indent = ((n - 1) * step).toFixed(2);
    return (
      `margin-top:0in;margin-bottom:0in;margin-left:${indent}in;line-height:115%;` +
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
    `margin-top:0in;margin-bottom:0in;line-height:115%;` +
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

  // Body heading levels seen on `data-heading="K"` rows (levels the user mapped to
  // a Word heading for nav + collapsibility); each gets a mapped-style rule below.
  const bodyHeadings = new Set<number>();
  const body = fragment
    // Title: a mapped Word heading (class + mso rule) when named, else the app's
    // direct level-1 look inline.
    .replace(/<p class="ws-title">([\s\S]*?)<\/p>/g, (_m, content: string) =>
      headingName
        ? `<p class="MsoTitle">${wrapTitleOnHeading(content)}</p>`
        : `<p style="${titleStyleAttr}">${wrapTitle(content)}</p>`,
    )
    // Nested rows. A `data-heading="K"` line is a level mapped to a Word heading:
    // map it to the destination "Heading K" style so a Use-Destination-Styles paste
    // adds it to the outline (nav + collapsible) and Word supplies its number.
    // Otherwise the app's per-level look, inline (+ <b> when bold). One regex with
    // an optional heading group so the two paths never double-match; the spacer
    // paragraph (nbsp, data-level only) passes through the else branch as an
    // ordinary inline body paragraph.
    .replace(
      /<p class="ws-lvl" data-level="([1-9])"(?: data-heading="([1-9])")?>([\s\S]*?)<\/p>/g,
      (_m, d: string, h: string | undefined, content: string) => {
        if (h) {
          const k = Number(h);
          bodyHeadings.add(k);
          return `<p class="MsoHeading${k}">${content}</p>`;
        }
        const i = Number(d) - 1;
        return `<p style="${directStyle(i, Number(d))}">${wrapBold(i, content)}</p>`;
      },
    );

  // Mapped-style rules: the title (when named) + each body heading level seen.
  // Every plain body level is inline, so it needs no rule.
  const titleRule = headingName
    ? `p.MsoTitle{mso-style-name:"${headingName}";mso-outline-level:${titleLevel}}`
    : "";
  const headingRules = Array.from(bodyHeadings)
    .map(
      (k) =>
        `p.MsoHeading${k}{mso-style-name:"Heading ${k}";mso-outline-level:${k}}`,
    )
    .join("");
  return (
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:w="urn:schemas-microsoft-com:office:word" ` +
    `xmlns="http://www.w3.org/TR/REC-html40">` +
    `<head><meta charset="utf-8">` +
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

"use client";

import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactPortal,
} from "react";
import { createPortal } from "react-dom";

/**
 * One entry in the command launcher. `run` is fired on select; the palette owns
 * dismissal (it calls `onClose()` right after), so a command never has to close
 * itself. `keywords` widens the match surface WITHOUT cluttering the row — e.g. a
 * "Copy section" command can also answer to "clipboard"/"export" without those
 * words appearing in the UI. `disabled` rows stay VISIBLE (so the launcher's
 * vocabulary is stable and discoverable) but are dimmed and unselectable.
 */
export interface Command {
  id: string;
  label: string;
  hint?: string; // secondary text shown muted after the label
  shortcut?: string; // e.g. "Ctrl+Enter" — rendered as kbd chips on the right
  section?: string; // group header the command sorts under
  keywords?: string; // extra text to match against (not shown)
  disabled?: boolean;
  run: () => void; // invoked on select; the palette calls onClose() after
}

/**
 * Case-insensitive SUBSEQUENCE test: every character of `query` must appear in
 * `text`, left-to-right, but not necessarily adjacent. This is the "fuzzy" feel
 * of a Ctrl+K launcher — typing "cpsec" still finds "Copy section" — while
 * staying dependency-free and O(n). An empty query matches everything, which is
 * what makes the palette show its full menu the instant it opens.
 */
function subsequenceMatch(query: string, text: string): boolean {
  if (query.length === 0) return true;
  const q = query.toLowerCase();
  const t = text.toLowerCase();
  let qi = 0;
  for (let ti = 0; ti < t.length && qi < q.length; ti++) {
    if (t[ti] === q[qi]) qi += 1;
  }
  return qi === q.length;
}

/** A `section` bucket plus its visible commands, in display order. */
interface Group {
  section?: string;
  items: Command[];
}

/**
 * Ctrl+K-style fuzzy command launcher. Self-contained and headless about WHEN it
 * opens — the PARENT owns the Ctrl+K shortcut and passes `open` — so this file
 * only owns the overlay, filtering, and in-list keyboard navigation.
 *
 * Rendering + dismissal mirror `Popover.tsx` (portal to `document.body`, a
 * full-screen backdrop button that closes on click, Escape-to-close), but this is
 * a CENTERED modal rather than a trigger-anchored panel, so there is no imperative
 * positioning — just a fixed, horizontally-centered panel ~15vh from the top.
 *
 * Focus discipline (the reason arrows are handled on the input, not on a global
 * listener): the text input keeps focus the whole time it is open, so typing
 * always lands in the query. ArrowUp/ArrowDown/Enter/Escape are intercepted on the
 * input's `onKeyDown` and drive a `highlight` index into the FLAT visible list;
 * `aria-activedescendant` points at the highlighted `role="option"` so screen
 * readers announce the moving selection without focus ever leaving the input.
 * Handling it here (rather than a `window` keydown listener) means no global state
 * to register/tear down and no risk of stealing keys from the rest of the app.
 */
export function CommandPalette({
  open,
  onClose,
  commands,
  placeholder,
}: {
  open: boolean;
  onClose: () => void;
  commands: Command[];
  placeholder?: string;
}): ReactPortal | null {
  const [query, setQuery] = useState("");
  // Index into `flat` (the filtered list in display order), NOT into `commands`.
  const [highlight, setHighlight] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  // Filter, then bucket by `section`. Order is preserved from `commands`, except
  // the no-section group is forced first (renders with no header) per the spec.
  // `flat` is the same set flattened in display order — the single source of truth
  // for keyboard nav — and `indexOf` lets each rendered row recover its flat index
  // (for highlight/aria) without threading a running counter through the JSX.
  const { groups, flat, indexOf } = useMemo(() => {
    const q = query.trim();
    const order: (string | undefined)[] = [];
    const bySection = new Map<string | undefined, Command[]>();
    for (const cmd of commands) {
      const haystack = `${cmd.label} ${cmd.hint ?? ""} ${cmd.keywords ?? ""}`;
      if (!subsequenceMatch(q, haystack)) continue;
      const key = cmd.section && cmd.section.length > 0 ? cmd.section : undefined;
      let bucket = bySection.get(key);
      if (!bucket) {
        bucket = [];
        bySection.set(key, bucket);
        order.push(key);
      }
      bucket.push(cmd);
    }
    // Stable sort that only lifts the no-section (undefined) bucket to the front;
    // every named section keeps its first-appearance order.
    order.sort((a, b) => (a === undefined ? -1 : b === undefined ? 1 : 0));
    const groups: Group[] = order.map((key) => ({
      section: key,
      items: bySection.get(key) ?? [],
    }));
    const flat = groups.flatMap((g) => g.items);
    const indexOf = new Map<Command, number>();
    flat.forEach((cmd, i) => indexOf.set(cmd, i));
    return { groups, flat, indexOf };
  }, [commands, query]);

  // On (re)open: clear any stale query and pull focus into the input. rAF defers
  // the focus until after the portal has painted, so the caret actually lands.
  useEffect(() => {
    if (!open) return;
    // Deliberate reset-on-open: clearing the stale query when the palette (re)opens
    // is exactly the "sync to an external trigger" this rule otherwise guards.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setQuery("");
    const id = requestAnimationFrame(() => inputRef.current?.focus());
    return () => cancelAnimationFrame(id);
  }, [open]);

  // Snap the highlight back to the first ENABLED row whenever the visible set
  // changes (i.e. the query changed) or the palette opens, so a filter never
  // leaves the selection on a hidden or disabled command.
  useEffect(() => {
    if (!open) return;
    const first = flat.findIndex((c) => !c.disabled);
    // Deliberate: re-home the highlight whenever the visible set changes so a
    // filter never strands the selection on a hidden/disabled row.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setHighlight(first);
  }, [flat, open]);

  // Keep the highlighted row scrolled into view during arrow navigation.
  useEffect(() => {
    if (!open || highlight < 0) return;
    const cmd = flat[highlight];
    if (!cmd) return;
    document
      .getElementById(`command-palette-option-${cmd.id}`)
      ?.scrollIntoView({ block: "nearest" });
  }, [highlight, flat, open]);

  // Step the highlight by ±1, wrapping around and skipping disabled rows. Bounded
  // by list length so an all-disabled list can't spin forever.
  const move = (delta: number) => {
    setHighlight((cur) => {
      if (flat.length === 0) return -1;
      let next = cur < 0 ? (delta > 0 ? -1 : 0) : cur;
      for (let step = 0; step < flat.length; step++) {
        next = (next + delta + flat.length) % flat.length;
        if (!flat[next].disabled) return next;
      }
      return cur;
    });
  };

  const runCommand = (cmd: Command) => {
    if (cmd.disabled) return;
    cmd.run();
    onClose();
  };

  const onInputKeyDown = (e: ReactKeyboardEvent<HTMLInputElement>) => {
    switch (e.key) {
      case "ArrowDown":
        e.preventDefault(); // don't let the caret jump to the end of the query
        move(1);
        break;
      case "ArrowUp":
        e.preventDefault();
        move(-1);
        break;
      case "Enter": {
        e.preventDefault();
        const cmd = flat[highlight];
        if (cmd) runCommand(cmd);
        break;
      }
      case "Escape":
        e.preventDefault();
        onClose();
        break;
    }
  };

  // Hooks above always run; only the render output is gated. `typeof document`
  // guards SSR, matching Popover.tsx — `createPortal` needs a real DOM node.
  if (!open || typeof document === "undefined") return null;

  const activeCmd = highlight >= 0 ? flat[highlight] : undefined;

  return createPortal(
    <>
      {/* Dimmed backdrop; a real <button> so a click OR Enter/Space dismisses. */}
      <button
        type="button"
        aria-label="Close command palette"
        tabIndex={-1}
        onClick={onClose}
        className="fixed inset-0 z-40 cursor-default bg-black/40"
      />
      <div
        role="dialog"
        aria-modal="true"
        aria-label="Command palette"
        className="fixed left-1/2 top-[15vh] z-50 flex max-h-[70vh] w-[min(34rem,calc(100vw-2rem))] -translate-x-1/2 flex-col overflow-hidden rounded-xl border border-border-strong bg-surface text-foreground shadow-[var(--shadow-4)]"
      >
        <div className="border-b border-border p-2">
          <input
            ref={inputRef}
            type="text"
            role="combobox"
            aria-label="Search commands"
            aria-expanded
            aria-controls="command-palette-list"
            aria-autocomplete="list"
            aria-activedescendant={
              activeCmd ? `command-palette-option-${activeCmd.id}` : undefined
            }
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={onInputKeyDown}
            placeholder={placeholder ?? "Type a command…"}
            spellCheck={false}
            autoComplete="off"
            className="h-9 w-full rounded border border-border-strong bg-surface px-2.5 text-sm text-foreground outline-none transition-colors focus:border-accent"
          />
        </div>

        <div
          id="command-palette-list"
          role="listbox"
          aria-label="Commands"
          className="min-h-0 flex-1 overflow-y-auto p-1.5"
        >
          {flat.length === 0 ? (
            <div className="px-3 py-6 text-center text-sm text-muted">
              No matching commands.
            </div>
          ) : (
            groups.map((group) => (
              <div key={group.section ?? "__none"} role="group">
                {group.section && (
                  <div className="px-2 pb-1 pt-2 text-[11px] font-semibold uppercase tracking-wide text-muted">
                    {group.section}
                  </div>
                )}
                {group.items.map((cmd) => {
                  const i = indexOf.get(cmd) ?? -1;
                  const active = i === highlight;
                  return (
                    <button
                      key={cmd.id}
                      id={`command-palette-option-${cmd.id}`}
                      type="button"
                      role="option"
                      aria-selected={active}
                      aria-disabled={cmd.disabled || undefined}
                      disabled={cmd.disabled}
                      // onMouseMove (not onMouseEnter) so the highlight only
                      // follows DELIBERATE mouse movement — a keyboard-driven
                      // highlight sliding under a stationary cursor won't get
                      // hijacked.
                      onMouseMove={() => {
                        if (!cmd.disabled && i !== highlight) setHighlight(i);
                      }}
                      onClick={() => runCommand(cmd)}
                      className={[
                        "flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-sm",
                        cmd.disabled
                          ? "cursor-not-allowed opacity-40"
                          : active
                            ? "bg-accent-subtle text-accent-text"
                            : "text-foreground",
                      ].join(" ")}
                    >
                      <span className="min-w-0 shrink-0">{cmd.label}</span>
                      {cmd.hint && (
                        <span className="min-w-0 truncate text-xs text-muted">
                          {cmd.hint}
                        </span>
                      )}
                      {cmd.shortcut && (
                        <span className="ml-auto flex shrink-0 items-center gap-1">
                          {cmd.shortcut.split("+").map((key, ki) => (
                            <kbd
                              key={ki}
                              className="rounded border border-border-strong px-1 py-0.5 font-mono text-[11px]"
                            >
                              {key.trim()}
                            </kbd>
                          ))}
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            ))
          )}
        </div>
      </div>
    </>,
    document.body,
  );
}

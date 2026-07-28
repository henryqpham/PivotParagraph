"use client";

import { useEffect, useLayoutEffect, useRef, type ReactNode } from "react";
import { createPortal } from "react-dom";

/**
 * Minimal popover: a panel anchored under its trigger, dismissed by a click on a
 * full-screen transparent backdrop or the Escape key. The caller wraps the
 * trigger + `<Popover>` in a `position: relative` element; the panel is PORTALED
 * to `document.body` with fixed positioning so it is never clipped by a
 * scrollable/overflow ancestor (the center builder scrolls), and it FLIPS above
 * the trigger + CLAMPS horizontally when it would fall off the viewport.
 *
 * Positioning is imperative (measure the trigger + panel in a layout effect,
 * write `style.left/top` on the panel) rather than via state — no extra render,
 * and it re-runs on scroll/resize so the panel stays glued to the trigger.
 * Shared by TableCard (the "Look" editors) and PasteInput (the "Document" popover).
 */
export function Popover({
  open,
  onClose,
  children,
}: {
  open: boolean;
  onClose: () => void;
  children: ReactNode;
}) {
  // A zero-impact marker whose parentElement is the caller's relative wrapper —
  // measured to place the portaled panel next to the trigger.
  const anchorRef = useRef<HTMLSpanElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const place = () => {
      const anchor = anchorRef.current?.parentElement;
      const panel = panelRef.current;
      if (!anchor || !panel) return;
      const a = anchor.getBoundingClientRect();
      const pw = panel.offsetWidth;
      const ph = panel.offsetHeight;
      const m = 8; // viewport margin
      // Horizontal: the panel's right edge aligns to the trigger's right edge,
      // then clamp fully into the viewport.
      let left = a.right - pw;
      left = Math.max(m, Math.min(left, window.innerWidth - pw - m));
      // Vertical: below the trigger if it fits, else flip above; then clamp.
      let top = a.bottom + 4;
      if (top + ph > window.innerHeight - m && a.top - ph - 4 >= m) {
        top = a.top - ph - 4;
      }
      top = Math.max(m, Math.min(top, window.innerHeight - ph - m));
      panel.style.left = `${left}px`;
      panel.style.top = `${top}px`;
      panel.style.visibility = "visible";
    };
    place();
    // Keep the panel glued to the trigger if the page scrolls or resizes.
    window.addEventListener("scroll", place, true);
    window.addEventListener("resize", place);
    return () => {
      window.removeEventListener("scroll", place, true);
      window.removeEventListener("resize", place);
    };
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;
  return (
    <>
      <span ref={anchorRef} className="hidden" aria-hidden />
      {typeof document !== "undefined" &&
        createPortal(
          <>
            <button
              type="button"
              aria-label="Close"
              tabIndex={-1}
              onClick={onClose}
              className="fixed inset-0 z-40 cursor-default"
            />
            {/* Rendered hidden at (0,0); the layout effect measures then places
                it and flips visibility — synchronous, so there is no flash. */}
            <div
              ref={panelRef}
              style={{ position: "fixed", left: 0, top: 0, visibility: "hidden" }}
              className="z-50 w-max min-w-44 rounded-md border border-border-strong bg-surface p-3 text-sm text-foreground shadow-[var(--shadow-4)]"
            >
              {children}
            </div>
          </>,
          document.body,
        )}
    </>
  );
}

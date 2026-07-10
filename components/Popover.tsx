import type { ReactNode } from "react";

/**
 * Minimal popover: a panel anchored under its trigger, dismissed by a click on a
 * full-screen transparent backdrop (no dependency, no portal). The caller wraps the
 * trigger + <Popover> in a `position: relative` element; the panel opens below and
 * to the right. Shared by TableCard (per-level "Look" editors) and PasteInput (the
 * top-band "Document" body settings).
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
  if (!open) return null;
  return (
    <>
      <button
        type="button"
        aria-label="Close"
        tabIndex={-1}
        onClick={onClose}
        className="fixed inset-0 z-10 cursor-default"
      />
      <div className="absolute right-0 top-full z-20 mt-1 w-max min-w-44 rounded-md border border-border-strong bg-surface p-3 text-sm text-foreground shadow-[var(--shadow-4)]">
        {children}
      </div>
    </>
  );
}

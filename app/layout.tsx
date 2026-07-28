import type { Metadata } from "next";
import "./globals.css";

// No `next/font` here on purpose: the app's type stack is Segoe UI (with a
// system fallback) and Cascadia Code / Consolas for mono, both declared in
// globals.css. The previously-loaded Geist families were never referenced by any
// rule, and dropping them removes a build-time font download — which also keeps
// this layout and the single-file build (main.tsx) rendering identically.

export const metadata: Metadata = {
  title: "PivotParagraph",
  description:
    "Turn wide Excel tables into a nested, Word-ready outline that fits the page.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="h-full flex flex-col overflow-hidden">{children}</body>
    </html>
  );
}

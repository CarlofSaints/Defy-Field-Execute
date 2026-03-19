import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Defy Field Execute",
  description: "Defy Field Execute — Reporting Platform by Atomic Marketing & Perigee",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="h-full">
      <body className="min-h-full flex flex-col antialiased">{children}</body>
    </html>
  );
}

import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "ASN TOOL GM",
  description: "Scan ASN, manage packing, calculate cartons and export Excel.",
  icons: { icon: "/icon.png", apple: "/icon.png" }
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}

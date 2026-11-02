import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "AGC Courier Billing",
  description: "Courier docket and billing management"
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}

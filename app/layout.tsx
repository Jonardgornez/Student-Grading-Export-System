import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Student Grading Export System",
  description:
    "Upload grading JSON and export attendance, activities, participation, and final grades to Excel.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}

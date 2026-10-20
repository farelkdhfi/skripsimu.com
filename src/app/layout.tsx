import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";
import { Providers } from "@/components/shared/providers";
import { SiteHeader } from "@/components/shared/site-header";
import { siteConfig } from "@/config/site";

const inter = Inter({ subsets: ["latin"] });
export const metadata: Metadata = { title: `${siteConfig.name} · Format Skripsi & Presentasi dengan AI`, description: siteConfig.description };

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="id">
      <body className={`${inter.className} min-h-screen bg-white text-slate-900 antialiased`}>
        <Providers><SiteHeader />{children}</Providers>
      </body>
    </html>
  );
}
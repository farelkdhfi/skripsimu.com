import Link from "next/link";
import { GraduationCap } from "lucide-react";
import { siteConfig } from "@/config/site";

export function SiteHeader() {
  return (
    <header className="sticky top-0 z-40 border-b border-slate-100 bg-white/90 backdrop-blur">
      <div className="mx-auto grid h-16 max-w-6xl grid-cols-[1fr_auto] items-center gap-4 px-4 md:grid-cols-[1fr_auto_1fr]">
        {/* brand */}
        <Link href="/" className="flex min-w-0 items-center gap-2 text-sm font-medium text-slate-900">
            <GraduationCap fill="black" className="size-6" />
          <span className="truncate">{siteConfig.name}</span>
        </Link>

        {/* nav pill (desktop) */}
        <nav className="hidden items-center gap-1 rounded-full border border-slate-200 bg-white p-1.5 text-sm text-slate-600 md:flex">
          <Link href="/formatter" className="rounded-full px-4 py-1.5 transition hover:bg-slate-100 hover:text-slate-900">
            Formatter
          </Link>
          <Link href="/slides" className="rounded-full px-4 py-1.5 transition hover:bg-slate-100 hover:text-slate-900">
            Slides
          </Link>
        </nav>

        <div className="flex items-center justify-end gap-3 sm:gap-4">
          {/* link biasa (mobile) */}
          <nav className="flex items-center gap-3 text-sm text-slate-600 md:hidden">
            <Link href="/formatter" className="hover:text-slate-900">Formatter</Link>
            <Link href="/slides" className="hover:text-slate-900">Slides</Link>
          </nav>
          <Link
            href="/dashboard"
            className="rounded-full bg-slate-900 px-4 py-1.5 text-sm text-white transition hover:bg-slate-800"
          >
            Dashboard
          </Link>
        </div>
      </div>
    </header>
  );
}
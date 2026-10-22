import Link from "next/link";
import { GraduationCap } from "lucide-react";
import { siteConfig } from "@/config/site";

export function SiteHeader() {
  return (
    <header className="sticky top-0 z-40 border-b border-neutral-100 bg-white/90 backdrop-blur">
      <div className="mx-auto grid h-16 max-w-6xl grid-cols-[1fr_auto] items-center gap-4 px-4 md:grid-cols-[1fr_auto_1fr]">
        {/* brand */}
        <Link href="/" className="flex min-w-0 items-center gap-2 text-sm font-medium text-neutral-900">
            <GraduationCap fill="black" className="size-6" />
          <span className="truncate font-semibold">{siteConfig.name}</span>
        </Link>

        {/* nav pill (desktop) */}
        <nav className="hidden items-center gap-1 rounded-full border border-neutral-200 bg-white p-1.5 text-sm text-neutral-600 md:flex">
          <Link href="/formatter" className="rounded-full px-4 py-1.5 transition hover:bg-neutral-100 hover:text-neutral-900">
            Formatter
          </Link>
          <Link href="/slides" className="rounded-full px-4 py-1.5 transition hover:bg-neutral-100 hover:text-neutral-900">
            Slides
          </Link>
        </nav>

        <div className="flex items-center justify-end gap-3 sm:gap-4">
          {/* link biasa (mobile) */}
          <nav className="flex items-center gap-3 text-sm text-neutral-600 md:hidden">
            <Link href="/formatter" className="hover:text-neutral-900">Formatter</Link>
            <Link href="/slides" className="hover:text-neutral-900">Slides</Link>
          </nav>
          <Link
            href="/dashboard"
            className="rounded-full bg-neutral-950 px-4 py-1.5 text-sm text-white transition hover:bg-neutral-800"
          >
            Dashboard
          </Link>
        </div>
      </div>
    </header>
  );
}
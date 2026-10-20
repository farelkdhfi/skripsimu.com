import Link from "next/link";
import { ArrowRight, FileCheck2, Presentation, ShieldCheck } from "lucide-react";

const items = [
  {
    href: "/formatter",
    icon: FileCheck2,
    idx: "01",
    title: "AI Thesis Formatter",
    desc: "Rapikan margin, font, heading, daftar isi, dan daftar pustaka. Unduh DOCX siap kumpul.",
    cta: "Format skripsi",
    card: "bg-purple-100",
    blob: "bg-purple-200",
  },
  {
    href: "/slides",
    icon: Presentation,
    idx: "02",
    title: "Skripsi to Slides",
    desc: "Ubah skripsi menjadi presentasi sidang lengkap dengan catatan presenter dan grafik.",
    cta: "Buat presentasi",
    card: "bg-green-100",
    blob: "bg-green-200",
  },
];

export default function Dashboard() {
  return (
    <main className="mx-auto max-w-5xl bg-white px-4 py-14">
      <div className="relative">
        <span className="pointer-events-none absolute -top-4 right-0 hidden size-28 rounded-full bg-purple-100 sm:block" aria-hidden="true" />
        <span className="pointer-events-none absolute right-24 top-12 hidden size-6 rounded-full bg-green-200 sm:block" aria-hidden="true" />
        <div className="relative">
          <p className="text-sm text-slate-500">Dashboard</p>
          <h1 className="mt-2 max-w-xl text-3xl font-normal leading-[1.1] tracking-tight text-slate-900 sm:text-4xl">
            Apa yang ingin kamu kerjakan hari ini?
          </h1>
          <p className="mt-4 inline-flex items-center gap-2 rounded-full bg-green-100 py-1.5 pl-1.5 pr-4 text-sm text-slate-700">
            <span className="grid size-7 place-items-center rounded-full bg-green-200 text-slate-900">
              <ShieldCheck className="size-4" />
            </span>
            Dokumenmu hanya diproses selama sesi dan tidak disimpan.
          </p>
        </div>
      </div>

      <div className="mt-10 grid gap-6 md:grid-cols-2">
        {items.map(({ href, icon: Icon, idx, title, desc, cta, card, blob }) => (
          <Link key={href} href={href} className="group flex flex-col">
            <div
              className={`relative min-h-[260px] flex-1 overflow-hidden rounded-[2rem] p-6 transition duration-300 group-hover:-translate-y-1 ${card}`}
            >
              <span className="inline-block rounded-full bg-white/70 px-2.5 py-1 text-xs font-medium text-slate-700">{idx}</span>
              <span className={`absolute -right-10 -top-10 size-44 rounded-full ${blob}`} aria-hidden="true" />
              <span className="absolute right-5 top-5 grid size-14 place-items-center rounded-full bg-white text-slate-900">
                <Icon className="size-6" />
              </span>
              <h2 className="absolute bottom-6 left-6 max-w-[70%] text-xl font-medium leading-snug text-slate-900">{title}</h2>
            </div>
            <p className="mt-4 max-w-md px-1 text-sm leading-relaxed text-slate-600">{desc}</p>
            <span className="mt-4 inline-flex items-center gap-1 px-1 text-sm font-medium text-slate-900">
              {cta} <ArrowRight className="size-4 transition group-hover:translate-x-1" />
            </span>
          </Link>
        ))}
      </div>
    </main>
  );
}
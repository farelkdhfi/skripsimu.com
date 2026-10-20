"use client";

import Link from "next/link";
import { useEffect } from "react";
import type { CSSProperties, ReactNode } from "react";
import {
  FileCheck2,
  Presentation,
  ShieldCheck,
  Upload,
  Settings2,
  Download,
  Plus,
  Check,
  Ruler,
  Type,
  ListOrdered,
  BookMarked,
  StickyNote,
  LayoutGrid,
  Lock,
  EyeOff,
  type LucideIcon,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { siteConfig } from "@/config/site";

/* ---------------------------------- DATA ---------------------------------- */

const faqs = [
  ["Apakah isi skripsi saya diubah AI?", "Tidak. AI hanya mengenali struktur (bab, subbab, daftar pustaka). Format diterapkan oleh library dokumen, dan isi tetap sama."],
  ["Apakah dokumen saya disimpan?", "Tidak. Dokumen diproses sementara selama sesi dan tidak disimpan di database."],
  ["DOCX atau PDF?", "DOCX memberi hasil paling akurat. PDF dibaca sebagai teks sehingga struktur bisa kurang presisi."],
  ["Apakah grafik di slide dijamin benar?", "Grafik hanya dibuat dari angka yang tertulis di skripsi. Tetap periksa sebelum sidang."],
];

const steps = [
  { icon: Upload, n: "01", title: "Unggah", desc: "Seret file DOCX atau PDF skripsimu.", tile: "bg-purple-100 text-purple-600" },
  { icon: Settings2, n: "02", title: "Atur", desc: "Pilih aturan format atau gaya presentasi.", tile: "bg-green-100 text-green-600" },
  { icon: Download, n: "03", title: "Unduh", desc: "Cek preview, lalu unduh DOCX atau PPTX.", tile: "bg-slate-900 text-white" },
];

const tutorials = [
  {
    format: "DOCX",
    tag: "bg-purple-100 text-purple-700",
    title: "Format skripsi",
    href: "/formatter",
    cta: "Buka formatter",
    steps: [
      ["Buka halaman Formatter", "Masuk lewat tombol \"Mulai format skripsi\"."],
      ["Unggah skripsi", "Pakai DOCX untuk hasil paling akurat. PDF juga bisa, tetapi struktur bisa kurang presisi."],
      ["Pilih aturan format", "Atur margin, font, spasi, heading, daftar isi, dan daftar pustaka sesuai pedoman kampusmu."],
      ["Cek dan unduh", "Lihat preview, pastikan isinya sama, lalu unduh DOCX."],
    ],
  },
  {
    format: "PPTX",
    tag: "bg-green-100 text-green-700",
    title: "Slide sidang",
    href: "/slides",
    cta: "Buka slides",
    steps: [
      ["Buka halaman Slides", "Masuk lewat tombol \"Buat presentasi\"."],
      ["Unggah skripsi", "Gunakan file skripsi yang sama, DOCX atau PDF."],
      ["Pilih gaya dan jumlah slide", "Tentukan gaya presentasi dan jumlah slide, mulai 5 sampai 40."],
      ["Periksa dan unduh", "Cek poin kunci, catatan presenter, dan angka di grafik, lalu unduh PPTX."],
    ],
  },
];

const tips = [
  "Pakai DOCX untuk hasil paling akurat.",
  "Periksa angka di grafik sebelum sidang.",
  "AI hanya mengatur format, isi skripsimu tidak diubah.",
];

/* ---------------------------------- CLOUDS --------------------------------- */

type Bump = [number, number]; // [lebar, tinggi] tiap gundukan awan

/* Satu pola = 1440px lebar. Diulang 3x supaya loop geser -1440px mulus
   dan tetap menutup layar lebar. */
function cloudPath(pattern: Bump[], repeat: number, base: number, height: number) {
  let d = `M0 ${height} V${base}`;
  for (let r = 0; r < repeat; r++) {
    for (const [w, h] of pattern) d += ` a${w / 2} ${h} 0 0 1 ${w} 0`;
  }
  return `${d} V${height} Z`;
}

const CLOUD_W = 4320;
const EXTRA = 12; // cadangan di bawah supaya tidak ada celah saat awan naik-turun

/* delay "rise" disetel supaya awan baru naik setelah intro selesai (~3.2s) */
const clouds = [
  { fill: "#D8B4FE", h: 300, base: 150, pattern: [[300, 70], [220, 50], [260, 85], [180, 45], [280, 65], [200, 55]] as Bump[], rise: "3.2s", bob: "9s", bobDelay: "-2s", drift: "80s", dir: "normal" },
  { fill: "#F3E8FF", h: 240, base: 110, pattern: [[240, 60], [320, 80], [200, 50], [280, 70], [160, 45], [240, 55]] as Bump[], rise: "3.4s", bob: "7s", bobDelay: "-4s", drift: "62s", dir: "reverse" },
  { fill: "#DCFCE7", h: 190, base: 90,  pattern: [[360, 70], [200, 45], [280, 65], [240, 50], [360, 60]] as Bump[],             rise: "3.6s", bob: "8s", bobDelay: "-1s", drift: "50s", dir: "normal" },
  { fill: "#FFFFFF", h: 130, base: 60,  pattern: [[260, 48], [340, 58], [220, 40], [300, 52], [320, 44]] as Bump[],             rise: "3.8s", bob: "6s", bobDelay: "-3s", drift: "42s", dir: "reverse" },
].map((c) => ({ ...c, d: cloudPath(c.pattern, 3, c.base, c.h + EXTRA) }));

/* ------------------------------ ANIMATION CSS ------------------------------ */

const css = `
html { scroll-behavior: smooth; }

@keyframes pop-in {
  0%   { opacity: 0; transform: translateY(22px) scale(.88); }
  100% { opacity: 1; transform: translateY(0) scale(1); }
}
@keyframes float {
  0%, 100% { transform: translateY(0); }
  50%      { transform: translateY(-6px); }
}
@keyframes fade-up {
  from { opacity: 0; transform: translateY(16px); }
  to   { opacity: 1; transform: none; }
}

/* ---- intro: ikon + nama web -> kabut menutup -> intro hilang ---- */
@keyframes intro-out {
  0%, 74% { opacity: 1; visibility: visible; }
  100%    { opacity: 0; visibility: hidden; }
}
@keyframes content-out {
  to { opacity: 0; transform: translateY(-10px) scale(.97); }
}
@keyframes fog-in {
  from { opacity: 0;   transform: scale(.35) translate(var(--fx, 0%), var(--fy, 0%)); }
  to   { opacity: .95; transform: scale(1) translate(0, 0); }
}

/* ---- hero: awan naik dari bawah layar, lalu bergelombang ---- */
@keyframes rise { from { transform: translateY(105%); } to { transform: translateY(0); } }
@keyframes sun-rise {
  from { opacity: 0; transform: translateY(45%); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes drift { from { transform: translateX(0); } to { transform: translateX(-1440px); } }
@keyframes bob {
  0%, 100% { transform: translateY(0) scaleY(1); }
  50%      { transform: translateY(-10px) scaleY(1.05); }
}
@keyframes ripple {
  0%   { transform: scale(1);    opacity: .55; }
  100% { transform: scale(1.45); opacity: 0; }
}
@keyframes spin-slow { to { transform: rotate(360deg); } }

.pop      { animation: pop-in .9s cubic-bezier(.2,.8,.2,1) both; }
.float    { animation: float 7s ease-in-out infinite; }
.fade-up  { animation: fade-up .9s cubic-bezier(.2,.8,.2,1) both; }

.intro         { animation: intro-out 3.9s ease forwards; }
.intro-content { animation: content-out .6s ease 2.5s forwards; }
.fog           { animation: fog-in 1.4s cubic-bezier(.3,.7,.2,1) both; }

.rise     { animation: rise 1.6s cubic-bezier(.2,.8,.2,1) both; }
.sun      { animation: sun-rise 1.6s cubic-bezier(.2,.8,.2,1) both; }
.drift    { animation: drift var(--d, 60s) linear infinite var(--dir, normal); will-change: transform; }
.bob      { animation: bob var(--b, 8s) ease-in-out var(--bd, 0s) infinite; transform-origin: 50% 100%; will-change: transform; }
.ripple   { animation: ripple 6s ease-out infinite; }
.spin-slow{ animation: spin-slow 30s linear infinite; }

@media (prefers-reduced-motion: reduce) {
  html { scroll-behavior: auto; }
  .intro { display: none !important; }
  .pop, .float, .fade-up, .rise, .sun, .drift, .bob, .ripple, .spin-slow { animation: none !important; }
  .ripple { opacity: 0; }
}
`;

/* ------------------------------ SCROLL RESET ------------------------------- */
/* Memastikan halaman selalu mulai dari paling atas saat refresh, dan
   mengunci scroll selama intro berjalan (durasi sama dengan .intro). */

const INTRO_MS = 3900;

function ScrollReset() {
  useEffect(() => {
    const root = document.documentElement;
    const toTop = () =>
      window.scrollTo({ top: 0, left: 0, behavior: "instant" as ScrollBehavior });

    // 1. matikan pemulihan scroll otomatis dari browser
    const prevRestoration = history.scrollRestoration;
    try {
      history.scrollRestoration = "manual";
    } catch {}

    // 2. paksa ke atas sekarang, lalu ulangi agar menang dari pemulihan
    //    scroll yang dilakukan browser belakangan (setelah load / bfcache)
    toTop();
    const raf = requestAnimationFrame(toTop);
    const onLoad = () => toTop();
    const onPageShow = (e: PageTransitionEvent) => {
      if (e.persisted) toTop();
    };
    window.addEventListener("load", onLoad);
    window.addEventListener("pageshow", onPageShow);

    // 3. kunci scroll selama intro (dilewati jika reduced motion)
    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    let timer: number | undefined;
    if (!reduce) {
      root.style.overflow = "hidden";
      timer = window.setTimeout(() => {
        root.style.overflow = "";
        toTop();
      }, INTRO_MS);
    }

    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener("load", onLoad);
      window.removeEventListener("pageshow", onPageShow);
      if (timer) window.clearTimeout(timer);
      root.style.overflow = "";
      try {
        history.scrollRestoration = prevRestoration;
      } catch {}
    };
  }, []);

  return null;
}

/* --------------------------------- ICONS ---------------------------------- */

function DocxIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 150 190" fill="none" className={className} role="img" aria-label="Dokumen DOCX">
      <path
        d="M18 6H100L136 42V172a12 12 0 0 1-12 12H18a12 12 0 0 1-12-12V18A12 12 0 0 1 18 6Z"
        fill="#fff"
        stroke="#CBD5E1"
        strokeWidth="2.5"
      />
      <path d="M100 6V30a12 12 0 0 0 12 12H136Z" fill="#F1F5F9" />
      <rect x="22" y="24" width="54" height="22" rx="11" fill="#0F172A" />
      <text x="49" y="39" textAnchor="middle" fontSize="10" fontWeight="700" fill="#fff" fontFamily="ui-sans-serif, system-ui, sans-serif">
        DOCX
      </text>
      <rect x="22" y="64" width="92" height="8" rx="4" fill="#CBD5E1" />
      <rect x="22" y="84" width="96" height="5" rx="2.5" fill="#EEF2F7" />
      <rect x="22" y="96" width="74" height="12" rx="4" fill="#DCFCE7" />
      <rect x="22" y="99" width="70" height="5" rx="2.5" fill="#E2E8F0" />
      <rect x="22" y="116" width="96" height="5" rx="2.5" fill="#EEF2F7" />
      <rect x="22" y="128" width="62" height="5" rx="2.5" fill="#EEF2F7" />
      <rect x="22" y="140" width="80" height="5" rx="2.5" fill="#EEF2F7" />
      <circle cx="112" cy="158" r="14" fill="#E9D5FF" />
      <path d="M105.5 158.5l5 5 9-10" stroke="#9333EA" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function PptIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 230 160" fill="none" className={className} role="img" aria-label="Slide PPT">
      <rect x="4" y="4" width="222" height="152" rx="16" fill="#fff" stroke="#E9D5FF" strokeWidth="2.5" />
      <circle cx="24" cy="24" r="4.5" fill="#E9D5FF" />
      <circle cx="38" cy="24" r="4.5" fill="#BBF7D0" />
      <circle cx="52" cy="24" r="4.5" fill="#CBD5E1" />
      <rect x="164" y="14" width="46" height="20" rx="10" fill="#C084FC" />
      <text x="187" y="28" textAnchor="middle" fontSize="10" fontWeight="700" fill="#fff" fontFamily="ui-sans-serif, system-ui, sans-serif">
        PPT
      </text>
      <rect x="22" y="48" width="104" height="9" rx="4.5" fill="#E2E8F0" />
      <rect x="22" y="64" width="72" height="5" rx="2.5" fill="#EEF2F7" />
      <rect x="26" y="112" width="16" height="28" rx="4" fill="#E2E8F0" />
      <rect x="50" y="96" width="16" height="44" rx="4" fill="#E9D5FF" />
      <rect x="74" y="84" width="16" height="56" rx="4" fill="#BBF7D0" />
      <rect x="98" y="104" width="16" height="36" rx="4" fill="#0F172A" />
      <path d="M20 141H122" stroke="#E2E8F0" strokeWidth="2" strokeLinecap="round" />
      <g transform="rotate(-90 172 96)">
        <circle cx="172" cy="96" r="26" stroke="#BBF7D0" strokeWidth="13" />
        <circle cx="172" cy="96" r="26" stroke="#D8B4FE" strokeWidth="13" strokeDasharray="70 163.4" />
        <circle cx="172" cy="96" r="26" stroke="#1E293B" strokeWidth="13" strokeDasharray="42 163.4" strokeDashoffset="-70" />
      </g>
    </svg>
  );
}

function BooksIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 140 120" fill="none" className={className} role="img" aria-label="Tumpukan buku">
      <rect x="12" y="80" width="116" height="28" rx="7" fill="#E9D5FF" />
      <rect x="30" y="80" width="8" height="28" fill="#C084FC" />
      <rect x="98" y="80" width="4" height="28" fill="#C084FC" />
      <rect x="22" y="52" width="100" height="28" rx="7" fill="#BBF7D0" />
      <rect x="38" y="52" width="8" height="28" fill="#86EFAC" />
      <rect x="86" y="52" width="4" height="28" fill="#86EFAC" />
      <rect x="14" y="24" width="92" height="28" rx="7" fill="#0F172A" />
      <rect x="30" y="24" width="6" height="28" fill="#E9D5FF" />
      <rect x="74" y="24" width="4" height="28" fill="#BBF7D0" />
    </svg>
  );
}

function OrbitIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 160 160" fill="none" className={className} aria-hidden="true">
      <circle cx="80" cy="80" r="30" stroke="#D8B4FE" strokeWidth="2" />
      <circle cx="80" cy="80" r="52" stroke="#E9D5FF" strokeWidth="2" strokeDasharray="4 6" />
      <circle cx="80" cy="80" r="74" stroke="#BBF7D0" strokeWidth="2" />
      <g className="spin-slow" style={{ transformOrigin: "80px 80px" }}>
        <circle cx="132" cy="80" r="7" fill="#0F172A" />
        <circle cx="28" cy="80" r="6" fill="#BBF7D0" />
        <circle cx="80" cy="6" r="6" fill="#D8B4FE" />
      </g>
      <circle cx="80" cy="80" r="18" fill="#0F172A" />
      <path d="M72 80.5l6 6 11-12" stroke="#fff" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function IconDot({ icon: Icon, className = "" }: { icon: LucideIcon; className?: string }) {
  return (
    <span className={`grid size-10 place-items-center rounded-full ${className}`}>
      <Icon className="size-4" />
    </span>
  );
}

/* ---------------------------------- INTRO ---------------------------------- */
/* Hanya tampil di awal: ikon DOCX + slides + nama web, lalu ditutup kabut
   dan overlay memudar, baru awan dan teks hero masuk. */

function Intro() {
  return (
    <div className="intro fixed inset-0 z-50 overflow-hidden bg-white" aria-hidden="true">
      <div className="intro-content absolute inset-0 z-10 flex flex-col items-center justify-center px-4">
        <div className="flex items-center gap-3 sm:gap-6">
          <div className="pop w-[92px] sm:w-[132px]" style={{ animationDelay: "0.1s" }}>
            <div className="float">
              <DocxIcon className="-rotate-6" />
            </div>
          </div>
          <div className="pop w-[150px] sm:w-[220px]" style={{ animationDelay: "0.35s" }}>
            <div className="float" style={{ animationDelay: "-3s" }}>
              <PptIcon className="rotate-3" />
            </div>
          </div>
        </div>
        <p
          className="fade-up mt-8 text-2xl font-normal tracking-tight text-slate-900 sm:text-3xl"
          style={{ animationDelay: "0.7s" }}
        >
          {siteConfig.name}
        </p>
      </div>

      {/* kabut: tiga gumpalan menyapu masuk dan menutup layar */}
      <div
        className="fog absolute -bottom-[25vmax] -left-[20vmax] z-20 size-[75vmax] rounded-full bg-purple-100 blur-3xl"
        style={{ "--fx": "-25%", "--fy": "25%", animationDelay: "1.8s" } as CSSProperties}
      />
      <div
        className="fog absolute -bottom-[20vmax] -right-[20vmax] z-20 size-[70vmax] rounded-full bg-green-100 blur-3xl"
        style={{ "--fx": "25%", "--fy": "25%", animationDelay: "2s" } as CSSProperties}
      />
      <div className="absolute inset-0 z-30 grid place-items-center">
        <div className="fog size-[90vmax] rounded-full bg-white blur-3xl" style={{ animationDelay: "2.2s" }} />
      </div>
    </div>
  );
}

/* ------------------------------- BENTO FITUR ------------------------------- */

function FeatureCell({
  idx,
  title,
  desc,
  className = "",
  cardClass,
  chipClass,
  titleClass = "max-w-[55%]",
  children,
}: {
  idx: string;
  title: string;
  desc: string;
  className?: string;
  cardClass: string;
  chipClass: string;
  titleClass?: string;
  children: ReactNode;
}) {
  return (
    <div className={`flex flex-col ${className}`}>
      {/* card: hanya ikon SVG + judul */}
      <div className={`relative min-h-[260px] flex-1 overflow-hidden rounded-[2rem] p-6 ${cardClass}`}>
        <span className={`inline-block rounded-full px-2.5 py-1 text-xs font-medium ${chipClass}`}>{idx}</span>
        {children}
        <h3 className={`absolute bottom-6 left-6 text-xl font-medium leading-snug ${titleClass}`}>{title}</h3>
      </div>
      {/* penjelasan: tanpa card */}
      <p className="mt-4 max-w-md px-1 text-sm leading-relaxed text-slate-600">{desc}</p>
    </div>
  );
}

function Bento() {
  return (
    <section id="fitur" className="mx-auto max-w-6xl px-4 py-24">
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        {/* judul section: teks biasa, bukan card */}
        <div className="col-span-2 flex flex-col justify-end pb-4 lg:col-start-3 lg:row-start-1">
          <p className="text-sm text-slate-500">Fitur</p>
          <h2 className="mt-2 text-4xl font-normal leading-[1.1] tracking-tight text-slate-900 sm:text-5xl">
            Satu alur dari skripsi sampai sidang
          </h2>
          <p className="mt-3 max-w-md text-sm leading-relaxed text-slate-600">
            Format skripsi, buat slide, dan tetap pegang kendali atas isinya.
          </p>
        </div>

        {/* 01 Formatter */}
        <FeatureCell
          idx="01"
          title="AI Thesis Formatter"
          desc="Margin, font, spasi, heading, daftar isi, dan daftar pustaka diterapkan konsisten."
          className="col-span-2 lg:col-span-2 lg:col-start-1 lg:row-span-2 lg:row-start-1"
          cardClass="bg-purple-100 text-slate-900"
          chipClass="bg-white/70 text-slate-700"
        >
          <div className="absolute left-6 top-16 flex gap-2">
            {[Ruler, Type, ListOrdered, BookMarked].map((Icon, i) => (
              <IconDot key={i} icon={Icon} className="bg-white text-slate-900" />
            ))}
          </div>
          <div className="absolute -bottom-8 -right-3 w-[46%] max-w-[200px] rotate-6">
            <div className="float">
              <DocxIcon />
            </div>
          </div>
        </FeatureCell>

        {/* ikon saja: buku */}
        <div className="col-span-1 grid min-h-[190px] place-items-center rounded-[2rem] bg-green-100 lg:col-start-3 lg:row-start-2">
          <div className="float">
            <BooksIcon className="w-24 sm:w-28" />
          </div>
        </div>

        {/* ikon saja: orbit */}
        <div className="col-span-1 grid min-h-[190px] place-items-center rounded-[2rem] border border-slate-200 bg-white lg:col-start-4 lg:row-start-2">
          <OrbitIcon className="w-28 sm:w-32" />
        </div>

        {/* 02 Slides */}
        <FeatureCell
          idx="02"
          title="Skripsi to Slides"
          desc="Alur presentasi 5 sampai 40 slide dengan poin kunci dan catatan presenter."
          className="col-span-2 lg:col-span-3"
          cardClass="bg-green-100 text-slate-900"
          chipClass="bg-white/70 text-slate-700"
        >
          <span className="absolute right-10 top-8 size-24 rounded-full bg-purple-200 sm:right-24 sm:size-32" />
          <div className="absolute left-6 top-16 flex gap-2">
            {[Presentation, StickyNote, LayoutGrid].map((Icon, i) => (
              <IconDot key={i} icon={Icon} className="bg-white text-slate-900" />
            ))}
          </div>
          <div className="absolute -bottom-8 right-4 w-[50%] max-w-[300px] -rotate-3">
            <div className="float" style={{ animationDelay: "-3s" }}>
              <PptIcon />
            </div>
          </div>
        </FeatureCell>

        {/* 03 Privasi */}
        <FeatureCell
          idx="03"
          title="Isi tetap milikmu"
          desc="Tanpa database dan tanpa penyimpanan permanen. AI tidak mengarang substansi."
          className="col-span-2 lg:col-span-1"
          cardClass="bg-slate-900 text-white"
          chipClass="bg-white/15 text-white"
          titleClass="max-w-[80%]"
        >
          <span className="absolute right-5 top-5 grid size-14 place-items-center rounded-full bg-green-200 text-slate-900">
            <ShieldCheck className="size-6" />
          </span>
          <div className="absolute left-6 top-16 flex gap-2">
            {[Lock, EyeOff].map((Icon, i) => (
              <IconDot key={i} icon={Icon} className="bg-white/10 text-white" />
            ))}
          </div>
        </FeatureCell>
      </div>
    </section>
  );
}

/* ---------------------------------- PAGE ---------------------------------- */

export default function Landing() {
  return (
    <main className="bg-white text-slate-700">
      <style dangerouslySetInnerHTML={{ __html: css }} />

      {/* selalu mulai dari atas + kunci scroll selama intro */}
      <ScrollReset />

      {/* intro (hanya saat awal masuk) */}
      <Intro />

      {/* ================================ HERO ================================ */}
      <section className="relative isolate overflow-hidden bg-white">
        {/* lingkaran ungu (matahari) + ripple */}
        <div className="pointer-events-none absolute inset-x-0 bottom-0 z-0 flex justify-center" aria-hidden="true">
          <div
            className="sun relative -mb-[100px] size-[440px] rounded-full bg-purple-200 sm:-mb-[200px] sm:size-[640px]"
            style={{ animationDelay: "3.1s" }}
          >
            <span className="ripple absolute inset-0 rounded-full border-2 border-purple-300" style={{ animationDelay: "5s" }} />
            <span className="ripple absolute inset-0 rounded-full border-2 border-purple-300" style={{ animationDelay: "8s" }} />
          </div>
        </div>

        {/* awan: naik dari bawah layar, lalu bergelombang loop */}
        <div className="pointer-events-none absolute inset-0 z-10" aria-hidden="true">
          {clouds.map((c, i) => (
            <div
              key={i}
              className="rise absolute inset-x-0"
              style={{ height: c.h + EXTRA, bottom: -EXTRA, animationDelay: c.rise }}
            >
              <div className="bob h-full" style={{ "--b": c.bob, "--bd": c.bobDelay } as CSSProperties}>
                <svg
                  className="drift block max-w-none"
                  width={CLOUD_W}
                  height={c.h + EXTRA}
                  viewBox={`0 0 ${CLOUD_W} ${c.h + EXTRA}`}
                  style={{ "--d": c.drift, "--dir": c.dir } as CSSProperties}
                >
                  <path d={c.d} fill={c.fill} />
                </svg>
              </div>
            </div>
          ))}
        </div>

        {/* teks hero */}
        <div className="relative z-20 mx-auto flex min-h-[calc(100svh-4rem)] max-w-3xl flex-col items-center justify-center px-4 pb-56 pt-10 text-center sm:pb-64">
          <p className="fade-up text-sm text-slate-700" style={{ animationDelay: "3.8s" }}>
            Untuk mahasiswa tingkat akhir
          </p>

          <h1
            className="fade-up mt-3 max-w-3xl text-4xl font-semibold leading-[1.1] tracking-tight text-slate-900 sm:text-6xl"
            style={{ animationDelay: "4s" }}
          >
            Skripsi rapi dan presentasi siap, tanpa begadang.
          </h1>

          <p
            className="fade-up mt-5 max-w-xl text-base leading-relaxed text-slate-700"
            style={{ animationDelay: "4.2s" }}
          >
            Unggah skripsimu, biarkan AI memahami strukturnya, lalu unduh DOCX yang sesuai standar dan slide sidang yang siap dipresentasikan.
          </p>

          <div
            className="fade-up mt-8 flex flex-col items-center justify-center gap-3 sm:flex-row"
            style={{ animationDelay: "4.4s" }}
          >
            <Button asChild className="h-12 rounded-full bg-slate-900 pl-2 pr-6 text-sm text-white hover:bg-slate-800">
              <Link href="/formatter">
                <span className="grid size-8 place-items-center rounded-full bg-slate-700">
                  <FileCheck2 className="size-4" />
                </span>
                Mulai format skripsi
              </Link>
            </Button>
            <Button
              asChild
              variant="outline"
              className="h-12 rounded-full border-slate-200 bg-white/90 pl-2 pr-6 text-sm text-slate-900 hover:bg-white"
            >
              <Link href="/slides">
                <span className="grid size-8 place-items-center rounded-full bg-slate-100">
                  <Presentation className="size-4" />
                </span>
                Buat presentasi
              </Link>
            </Button>
          </div>
        </div>
      </section>

      {/* ================================ FITUR =============================== */}
      <Bento />

      {/* ============================= TIGA LANGKAH =========================== */}
      <section id="cara-kerja" className="mx-auto max-w-5xl px-4 pb-20 pt-8">
        <div className="text-center">
          <p className="text-sm text-slate-500">Cara kerja</p>
          <h2 className="mt-2 text-3xl font-normal tracking-tight text-slate-900 sm:text-4xl">Tiga langkah saja</h2>
        </div>

        <div className="relative mt-14 grid gap-10 md:grid-cols-3">
          <div className="absolute left-[17%] right-[17%] top-8 hidden border-t border-dashed border-slate-300 md:block" aria-hidden="true" />
          {steps.map(({ icon: Icon, n, title, desc, tile }) => (
            <div key={n} className="relative text-center">
              <div className={`relative mx-auto grid size-16 place-items-center rounded-full ${tile}`}>
                <Icon className="size-6" />
                <span className="absolute -right-1 -top-1 grid size-6 place-items-center rounded-full border border-slate-200 bg-white text-[10px] font-medium text-slate-900">
                  {n}
                </span>
              </div>
              <h3 className="mt-5 font-medium text-slate-900">{title}</h3>
              <p className="mx-auto mt-1 max-w-[16rem] text-sm text-slate-600">{desc}</p>
            </div>
          ))}
        </div>
      </section>

      {/* ============================ TUTORIAL SINGKAT ======================== */}
      <section id="tutorial" className="mx-auto max-w-5xl px-4 pb-24 pt-8">
        <div className="max-w-xl">
          <p className="text-sm text-slate-500">Tutorial</p>
          <h2 className="mt-2 text-3xl font-normal tracking-tight text-slate-900 sm:text-4xl">Tutorial singkat</h2>
          <p className="mt-2 text-sm text-slate-600">Dua alur utama. Ikuti langkahnya satu per satu.</p>
        </div>

        <div className="mt-10 grid gap-12 md:grid-cols-2 md:gap-16">
          {tutorials.map((t) => (
            <div key={t.title}>
              <div className="flex items-center gap-3">
                <span className={`rounded-full px-2.5 py-1 text-xs font-medium ${t.tag}`}>{t.format}</span>
                <h3 className="text-lg font-medium text-slate-900">{t.title}</h3>
              </div>

              <ol className="mt-6">
                {t.steps.map(([title, desc], i) => (
                  <li
                    key={title}
                    className="relative pb-6 pl-11 last:pb-0 before:absolute before:bottom-1 before:left-3.5 before:top-9 before:w-px before:bg-slate-200 last:before:hidden"
                  >
                    <span className="absolute left-0 top-0 grid size-7 place-items-center rounded-full bg-slate-900 text-xs font-medium text-white">
                      {i + 1}
                    </span>
                    <h4 className="font-medium text-slate-900">{title}</h4>
                    <p className="mt-1 text-sm leading-relaxed text-slate-600">{desc}</p>
                  </li>
                ))}
              </ol>

              <Button
                asChild
                variant="outline"
                className="mt-8 rounded-full border-slate-300 bg-white px-6 text-slate-900 hover:bg-slate-50"
              >
                <Link href={t.href}>{t.cta}</Link>
              </Button>
            </div>
          ))}
        </div>

        <div className="mt-14 grid gap-3 rounded-[2rem] bg-green-50 p-6 text-sm text-slate-700 sm:grid-cols-3">
          {tips.map((tip) => (
            <p key={tip} className="flex items-start gap-2">
              <Check className="mt-0.5 size-4 shrink-0 text-slate-900" />
              {tip}
            </p>
          ))}
        </div>
      </section>

      {/* ================================= FAQ ================================ */}
      <section id="faq" className="mx-auto max-w-5xl px-4 pb-24">
        <div className="grid gap-8 md:grid-cols-[1fr_2fr] md:gap-12">
          <div>
            <p className="text-sm text-slate-500">FAQ</p>
            <h2 className="mt-2 text-3xl font-normal tracking-tight text-slate-900 sm:text-4xl">Pertanyaan umum</h2>
          </div>
          <div className="divide-y divide-slate-200 border-y border-slate-200">
            {faqs.map(([q, a]) => (
              <details key={q} className="group py-4">
                <summary className="flex cursor-pointer list-none items-center justify-between gap-4 font-medium text-slate-900 [&::-webkit-details-marker]:hidden">
                  {q}
                  <Plus className="size-4 shrink-0 text-slate-400 transition-transform group-open:rotate-45" />
                </summary>
                <p className="mt-2 pr-8 text-sm leading-relaxed text-slate-600">{a}</p>
              </details>
            ))}
          </div>
        </div>
      </section>

      {/* ================================= CTA ================================ */}
      <section className="mx-auto max-w-6xl px-4 pb-20">
        <div className="relative overflow-hidden rounded-[2rem] bg-slate-900 p-8 text-white sm:p-12">
          <span className="absolute -bottom-24 -right-16 size-64 rounded-full bg-purple-300" aria-hidden="true" />
          <span className="absolute right-44 top-10 hidden size-5 rounded-full bg-green-200 sm:block" aria-hidden="true" />
          <div className="relative max-w-lg">
            <h2 className="text-3xl font-normal tracking-tight">Siap merapikan skripsimu?</h2>
            <p className="mt-3 text-sm text-slate-300">Mulai dari format DOCX, lanjut ke slide sidang dalam satu alur.</p>
            <div className="mt-8 flex flex-col gap-3 sm:flex-row">
              <Button asChild className="rounded-full bg-white px-6 text-slate-900 hover:bg-slate-100">
                <Link href="/formatter">Mulai format skripsi</Link>
              </Button>
              <Button
                asChild
                variant="outline"
                className="rounded-full border-slate-600 bg-transparent px-6 text-white hover:bg-slate-800 hover:text-white"
              >
                <Link href="/slides">Buat presentasi</Link>
              </Button>
            </div>
          </div>
        </div>
      </section>

      <footer className="border-t border-slate-200 py-6 text-center text-xs text-slate-400">
        Periksa kembali hasil akhir sebelum dikumpulkan.
      </footer>
    </main>
  );
}
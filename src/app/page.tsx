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
  { icon: Upload, n: "01", title: "Unggah", desc: "Seret file DOCX atau PDF skripsimu.", tile: "bg-pink-400 text-zinc-950" },
  { icon: Settings2, n: "02", title: "Atur", desc: "Pilih aturan format atau gaya presentasi.", tile: "bg-violet-400 text-zinc-950" },
  { icon: Download, n: "03", title: "Unduh", desc: "Cek preview, lalu unduh DOCX atau PPTX.", tile: "bg-emerald-400 text-zinc-950" },
];

const tutorials = [
  {
    format: "DOCX",
    tag: "bg-violet-400/15 text-violet-300 ring-1 ring-violet-400/30",
    dot: "bg-violet-400",
    line: "before:bg-violet-400/25",
    btn: "border-violet-400/50 text-violet-300 hover:border-violet-400 hover:bg-violet-400 hover:text-zinc-950",
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
    tag: "bg-cyan-400/15 text-cyan-300 ring-1 ring-cyan-400/30",
    dot: "bg-cyan-400",
    line: "before:bg-cyan-400/25",
    btn: "border-cyan-400/50 text-cyan-300 hover:border-cyan-400 hover:bg-cyan-400 hover:text-zinc-950",
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
const tipColors = ["text-emerald-400", "text-cyan-400", "text-pink-400"];

/* ---------------------------------- LOGO ----------------------------------- */
/* Semua yang perlu diganti saat pakai logo asli ada di blok ini.
   - LOGO_VIEWBOX   : viewBox SVG logomu
   - LOGO_TRANSFORM : transform pada <g> (hasil trace/potrace biasanya punya,
                      contoh: "translate(0,1024) scale(0.1,-0.1)")
   - LOGO_STROKE    : tebal outline (dalam satuan koordinat logo)
   - LOGO_GLOW      : besar blur glow (dalam satuan koordinat logo)
   - LOGO_SPAN      : lebar satu putaran gradien (kira-kira lebar logo)
   - LOGO_SHAPES    : daftar path. closed:true = bentuk tertutup (diisi putih
                      di fase awal), tanpa closed = garis terbuka.

   Logo saat ini: topi wisuda kolaboratif (dua figur membentuk papan topi,
   pita topi di bawah, dan rumbai di kanan). Koordinat berbasis kanvas 1024. */

const LOGO_VIEWBOX = "0 170 1024 800";
const LOGO_TRANSFORM: string | undefined = undefined;
const LOGO_STROKE = 10;
const LOGO_GLOW = 24;
const LOGO_SPAN = 1024;
const LOGO_SHAPES: { d: string; closed?: boolean }[] = [
  /* kepala figur kiri */
  { d: "M252 259 a69 69 0 1 0 138 0 a69 69 0 1 0 -138 0 Z", closed: true },
  /* kepala figur kanan */
  { d: "M598 306 a69 69 0 1 0 138 0 a69 69 0 1 0 -138 0 Z", closed: true },
  /* badan kiri + sisi kiri papan topi */
  {
    d: "M38 504 L385 330 C440 303 505 302 540 338 C565 365 552 395 522 410 C440 447 345 495 320 548 C303 588 325 618 362 647 C415 690 495 738 578 742 C535 768 480 768 430 750 C365 727 300 695 230 665 L62 586 C18 567 4 540 14 522 C19 513 28 509 38 504 Z",
    closed: true,
  },
  /* badan kanan + sisi kanan papan topi + tali + rumbai */
  {
    d: "M437 655 C440 635 475 610 505 580 C545 540 555 505 570 465 C585 428 610 400 655 395 C700 391 735 408 760 420 L935 500 C965 513 980 530 982 560 L982 695 C998 700 1003 715 1003 730 C1003 748 993 762 980 770 C1000 800 1018 850 1016 885 C1014 925 990 950 960 950 C925 950 905 925 907 885 C910 850 930 800 945 770 C932 762 922 748 922 730 C922 715 932 702 942 697 L940 568 L700 680 C640 708 580 722 530 718 C480 714 445 685 437 655 Z",
    closed: true,
  },
  /* pita topi */
  {
    d: "M222 690 C330 745 420 790 500 790 C590 790 690 745 778 690 C786 695 785 705 785 720 L785 790 C770 850 640 932 500 932 C360 932 230 850 215 790 L215 705 C215 695 218 690 222 690 Z",
    closed: true,
  },
];

const RAINBOW = ["#ff5c8a", "#ffb84d", "#4ade80", "#22d3ee", "#818cf8", "#e879f9"];

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

/* delay "rise" disetel supaya awan baru naik menjelang intro selesai (~7.3s) */
const clouds = [
  { fill: "#2b2b40", h: 300, base: 150, pattern: [[300, 70], [220, 50], [260, 85], [180, 45], [280, 65], [200, 55]] as Bump[], rise: "6.5s", bob: "9s", bobDelay: "-2s", drift: "80s", dir: "normal" },
  { fill: "#202032", h: 240, base: 110, pattern: [[240, 60], [320, 80], [200, 50], [280, 70], [160, 45], [240, 55]] as Bump[], rise: "6.7s", bob: "7s", bobDelay: "-4s", drift: "62s", dir: "reverse" },
  { fill: "#161624", h: 190, base: 90,  pattern: [[360, 70], [200, 45], [280, 65], [240, 50], [360, 60]] as Bump[],             rise: "6.9s", bob: "8s", bobDelay: "-1s", drift: "50s", dir: "normal" },
  { fill: "#09090b", h: 130, base: 60,  pattern: [[260, 48], [340, 58], [220, 40], [300, 52], [320, 44]] as Bump[],             rise: "7.1s", bob: "6s", bobDelay: "-3s", drift: "42s", dir: "reverse" },
].map((c) => ({ ...c, d: cloudPath(c.pattern, 3, c.base, c.h + EXTRA) }));

/* ------------------------------ ANIMATION CSS ------------------------------ */
/* Timeline intro (total 7.3s):
   0.1s  logo muncul (putih)
   1.3s  logo berubah jadi outline warna-warni, tergambar pelan
   1.9s  nama web muncul
   ~3.1s - 5.4s  logo + nama web tampil penuh (waktu "tahan")
   5.4s  konten intro mulai hilang
   4.7s - 6.5s  kabut menutup layar
   6.1s - 7.3s  overlay intro memudar, hero masuk */

const css = `
:root {
  --c1: #ff5c8a;
  --c2: #ffb84d;
  --c3: #4ade80;
  --c4: #22d3ee;
  --c5: #818cf8;
  --c6: #e879f9;
}
html { scroll-behavior: smooth; background: #09090b; color-scheme: dark; }

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

/* ---- warna-warni: teks, background, dan border yang mengalir ---- */
@keyframes flow { to { background-position: var(--flow, 200%) 0; } }

.rainbow-text {
  background-image: linear-gradient(90deg, var(--c1), var(--c2), var(--c3), var(--c4), var(--c5), var(--c6), var(--c1));
  background-size: 200% 100%;
  -webkit-background-clip: text;
  background-clip: text;
  color: transparent;
  animation: flow 7s linear infinite;
  --flow: 200%;
}
.rainbow-bg {
  background-image: linear-gradient(90deg, var(--c1), var(--c2), var(--c3), var(--c4), var(--c5), var(--c6), var(--c1));
  background-size: 200% 100%;
  animation: flow 7s linear infinite;
  --flow: 200%;
}
.rb-border { position: relative; }
.rb-border::before {
  content: "";
  position: absolute;
  inset: 0;
  border-radius: inherit;
  padding: 1.5px;
  background-image: linear-gradient(110deg, var(--c1), var(--c2), var(--c3), var(--c4), var(--c5), var(--c6), var(--c1));
  background-size: 300% 100%;
  animation: flow 9s linear infinite;
  --flow: 300%;
  -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
  -webkit-mask-composite: xor;
  mask: linear-gradient(#000 0 0) content-box exclude, linear-gradient(#000 0 0);
  pointer-events: none;
}

/* ---- intro: logo putih -> outline warna-warni -> tahan -> kabut -> intro hilang ---- */
@keyframes intro-out {
  0%, 84% { opacity: 1; visibility: visible; }
  100%    { opacity: 0; visibility: hidden; }
}
@keyframes content-out {
  to { opacity: 0; transform: translateY(-10px) scale(.97); }
}
@keyframes fog-in {
  from { opacity: 0;   transform: scale(.35) translate(var(--fx, 0%), var(--fy, 0%)); }
  to   { opacity: .95; transform: scale(1) translate(0, 0); }
}
@keyframes logo-white-out { to { opacity: 0; } }
@keyframes logo-in        { to { opacity: 1; } }
@keyframes logo-glow-in   { to { opacity: .55; } }
@keyframes logo-draw      { from { stroke-dashoffset: 1; } to { stroke-dashoffset: 0; } }
@keyframes halo-in        { to { opacity: .4; } }

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

.intro         { animation: intro-out 7.3s ease forwards; }
.intro-content { animation: content-out .6s ease 5.4s forwards; }
.fog           { animation: fog-in 1.4s cubic-bezier(.3,.7,.2,1) both; }

.logo-white { animation: logo-white-out 1s ease 1.3s forwards; }
.logo-line  { opacity: 0; animation: logo-in .9s ease 1.3s forwards; }
.logo-glow  { opacity: 0; animation: logo-glow-in 1.2s ease 1.6s forwards; }
.logo-draw  { stroke-dasharray: 1; animation: logo-draw 1.8s cubic-bezier(.4,0,.2,1) 1.3s both; }
.logo-halo  {
  opacity: 0;
  border-radius: 9999px;
  filter: blur(64px);
  background: conic-gradient(from 0deg, var(--c1), var(--c2), var(--c3), var(--c4), var(--c5), var(--c6), var(--c1));
  animation: halo-in 1.6s ease 1.4s forwards, spin-slow 18s linear infinite;
}

.rise     { animation: rise 1.6s cubic-bezier(.2,.8,.2,1) both; }
.sun      { animation: sun-rise 1.6s cubic-bezier(.2,.8,.2,1) both; }
.drift    { animation: drift var(--d, 60s) linear infinite var(--dir, normal); will-change: transform; }
.bob      { animation: bob var(--b, 8s) ease-in-out var(--bd, 0s) infinite; transform-origin: 50% 100%; will-change: transform; }
.ripple   { animation: ripple 6s ease-out infinite; }
.spin-slow{ animation: spin-slow 30s linear infinite; }

@media (prefers-reduced-motion: reduce) {
  html { scroll-behavior: auto; }
  .intro { display: none !important; }
  .pop, .float, .fade-up, .rise, .sun, .drift, .bob, .ripple, .spin-slow,
  .rainbow-text, .rainbow-bg, .rb-border::before { animation: none !important; }
  .ripple { opacity: 0; }
}
`;

/* ------------------------------ SCROLL RESET ------------------------------- */
/* Memastikan halaman selalu mulai dari paling atas saat refresh, dan
   mengunci scroll selama intro berjalan (durasi sama dengan .intro). */

const INTRO_MS = 7300;

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

/* Logo web: fase 1 putih solid, fase 2 hanya outline gradien warna-warni
   yang mengalir (SMIL animateTransform) + glow blur di belakangnya. */
function Logo({ className = "" }: { className?: string }) {
  const n = RAINBOW.length;
  return (
    <svg
      viewBox={LOGO_VIEWBOX}
      fill="none"
      className={className}
      style={{ overflow: "visible" }}
      role="img"
      aria-label={`Logo ${siteConfig.name}`}
    >
      <defs>
        <linearGradient
          id="logo-rb"
          gradientUnits="userSpaceOnUse"
          x1="0"
          y1="0"
          x2={LOGO_SPAN}
          y2="0"
          spreadMethod="repeat"
        >
          {RAINBOW.map((c, i) => (
            <stop key={i} offset={i / n} stopColor={c} />
          ))}
          <stop offset="1" stopColor={RAINBOW[0]} />
          <animateTransform
            attributeName="gradientTransform"
            type="translate"
            from="0 0"
            to={`${LOGO_SPAN} 0`}
            dur="4.5s"
            repeatCount="indefinite"
          />
        </linearGradient>
        <filter id="logo-blur" x="-30%" y="-30%" width="160%" height="160%">
          <feGaussianBlur stdDeviation={LOGO_GLOW} />
        </filter>
      </defs>

      <g transform={LOGO_TRANSFORM}>
        {/* fase 1: putih */}
        <g
          className="logo-white"
          stroke="#fff"
          strokeWidth={LOGO_STROKE + 2}
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          {LOGO_SHAPES.map((s, i) => (
            <path key={i} d={s.d} fill={s.closed ? "#fff" : "none"} />
          ))}
        </g>

        {/* fase 2: glow */}
        <g className="logo-glow" filter="url(#logo-blur)">
          <g stroke="url(#logo-rb)" strokeWidth={LOGO_STROKE + 14} strokeLinecap="round" strokeLinejoin="round">
            {LOGO_SHAPES.map((s, i) => (
              <path key={i} d={s.d} />
            ))}
          </g>
        </g>

        {/* fase 2: outline warna-warni, tergambar pelan */}
        <g className="logo-line" stroke="url(#logo-rb)" strokeWidth={LOGO_STROKE} strokeLinecap="round" strokeLinejoin="round">
          {LOGO_SHAPES.map((s, i) => (
            <path key={i} d={s.d} pathLength={1} className="logo-draw" />
          ))}
        </g>
      </g>
    </svg>
  );
}

function DocxIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 150 190" fill="none" className={className} role="img" aria-label="Dokumen DOCX">
      <path
        d="M18 6H100L136 42V172a12 12 0 0 1-12 12H18a12 12 0 0 1-12-12V18A12 12 0 0 1 18 6Z"
        fill="#18181b"
        stroke="#52525b"
        strokeWidth="2.5"
      />
      <path d="M100 6V30a12 12 0 0 0 12 12H136Z" fill="#27272a" />
      <rect x="22" y="24" width="54" height="22" rx="11" fill="#a78bfa" />
      <text x="49" y="39" textAnchor="middle" fontSize="10" fontWeight="700" fill="#09090b" fontFamily="ui-sans-serif, system-ui, sans-serif">
        DOCX
      </text>
      <rect x="22" y="64" width="92" height="8" rx="4" fill="#f472b6" />
      <rect x="22" y="84" width="96" height="5" rx="2.5" fill="#3f3f46" />
      <rect x="22" y="96" width="74" height="12" rx="4" fill="#2e1065" />
      <rect x="22" y="99" width="70" height="5" rx="2.5" fill="#a78bfa" />
      <rect x="22" y="116" width="96" height="5" rx="2.5" fill="#3f3f46" />
      <rect x="22" y="128" width="62" height="5" rx="2.5" fill="#3f3f46" />
      <rect x="22" y="140" width="80" height="5" rx="2.5" fill="#3f3f46" />
      <circle cx="112" cy="158" r="14" fill="#064e3b" />
      <path d="M105.5 158.5l5 5 9-10" stroke="#34d399" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function PptIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 230 160" fill="none" className={className} role="img" aria-label="Slide PPT">
      <rect x="4" y="4" width="222" height="152" rx="16" fill="#18181b" stroke="#52525b" strokeWidth="2.5" />
      <circle cx="24" cy="24" r="4.5" fill="#f472b6" />
      <circle cx="38" cy="24" r="4.5" fill="#fbbf24" />
      <circle cx="52" cy="24" r="4.5" fill="#34d399" />
      <rect x="164" y="14" width="46" height="20" rx="10" fill="#22d3ee" />
      <text x="187" y="28" textAnchor="middle" fontSize="10" fontWeight="700" fill="#09090b" fontFamily="ui-sans-serif, system-ui, sans-serif">
        PPT
      </text>
      <rect x="22" y="48" width="104" height="9" rx="4.5" fill="#52525b" />
      <rect x="22" y="64" width="72" height="5" rx="2.5" fill="#3f3f46" />
      <rect x="26" y="112" width="16" height="28" rx="4" fill="#f472b6" />
      <rect x="50" y="96" width="16" height="44" rx="4" fill="#fbbf24" />
      <rect x="74" y="84" width="16" height="56" rx="4" fill="#34d399" />
      <rect x="98" y="104" width="16" height="36" rx="4" fill="#818cf8" />
      <path d="M20 141H122" stroke="#52525b" strokeWidth="2" strokeLinecap="round" />
      <g transform="rotate(-90 172 96)">
        <circle cx="172" cy="96" r="26" stroke="#3f3f46" strokeWidth="13" />
        <circle cx="172" cy="96" r="26" stroke="#22d3ee" strokeWidth="13" strokeDasharray="70 163.4" />
        <circle cx="172" cy="96" r="26" stroke="#e879f9" strokeWidth="13" strokeDasharray="42 163.4" strokeDashoffset="-70" />
      </g>
    </svg>
  );
}

function BooksIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 140 120" fill="none" className={className} role="img" aria-label="Tumpukan buku">
      <rect x="12" y="80" width="116" height="28" rx="7" fill="#818cf8" />
      <rect x="30" y="80" width="8" height="28" fill="#4f46e5" />
      <rect x="98" y="80" width="4" height="28" fill="#4f46e5" />
      <rect x="22" y="52" width="100" height="28" rx="7" fill="#22d3ee" />
      <rect x="38" y="52" width="8" height="28" fill="#0891b2" />
      <rect x="86" y="52" width="4" height="28" fill="#0891b2" />
      <rect x="14" y="24" width="92" height="28" rx="7" fill="#f472b6" />
      <rect x="30" y="24" width="6" height="28" fill="#be185d" />
      <rect x="74" y="24" width="4" height="28" fill="#db2777" />
    </svg>
  );
}

function OrbitIcon({ className = "" }: { className?: string }) {
  return (
    <svg viewBox="0 0 160 160" fill="none" className={className} aria-hidden="true">
      <circle cx="80" cy="80" r="30" stroke="#71717a" strokeWidth="2" />
      <circle cx="80" cy="80" r="52" stroke="#22d3ee" strokeOpacity=".6" strokeWidth="2" strokeDasharray="4 6" />
      <circle cx="80" cy="80" r="74" stroke="#3f3f46" strokeWidth="2" />
      <g className="spin-slow" style={{ transformOrigin: "80px 80px" }}>
        <circle cx="132" cy="80" r="7" fill="#fbbf24" />
        <circle cx="28" cy="80" r="6" fill="#34d399" />
        <circle cx="80" cy="6" r="6" fill="#f472b6" />
      </g>
      <circle cx="80" cy="80" r="18" fill="#818cf8" />
      <path d="M72 80.5l6 6 11-12" stroke="#09090b" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
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

/* label kecil di atas judul section, dengan titik warna-warni */
function Eyebrow({ children }: { children: ReactNode }) {
  return (
    <p className="flex items-center gap-2 text-sm font-medium">
      <span className="rainbow-bg size-1.5 rounded-full" aria-hidden="true" />
      <span className="rainbow-text">{children}</span>
    </p>
  );
}

/* ---------------------------------- INTRO ---------------------------------- */
/* Hanya tampil di awal: logo (putih -> outline warna-warni) + nama web,
   ditahan beberapa detik, lalu ditutup kabut berwarna dan overlay memudar,
   baru awan dan teks hero masuk. */

function Intro() {
  return (
    <div className="intro fixed inset-0 z-50 overflow-hidden bg-zinc-950" aria-hidden="true">
      <div className="intro-content absolute inset-0 z-10 flex flex-col items-center justify-center px-4">
        <div className="relative">
          {/* halo warna-warni di belakang logo */}
          <div className="pointer-events-none absolute inset-0 grid place-items-center">
            <div className="logo-halo size-[18rem] sm:size-[26rem]" />
          </div>
          <div className="pop relative w-40 sm:w-56" style={{ animationDelay: "0.1s" }}>
            <div className="float">
              <Logo className="w-full" />
            </div>
          </div>
        </div>
        <p
          className="fade-up mt-10 text-2xl font-normal tracking-tight text-zinc-100 sm:text-3xl"
          style={{ animationDelay: "1.9s" }}
        >
          {siteConfig.name}
        </p>
      </div>

      {/* kabut berwarna: tiga gumpalan menyapu masuk dan menutup layar */}
      <div
        className="fog absolute -bottom-[25vmax] -left-[20vmax] z-20 size-[75vmax] rounded-full bg-fuchsia-500/60 blur-3xl"
        style={{ "--fx": "-25%", "--fy": "25%", animationDelay: "4.7s" } as CSSProperties}
      />
      <div
        className="fog absolute -bottom-[20vmax] -right-[20vmax] z-20 size-[70vmax] rounded-full bg-cyan-400/50 blur-3xl"
        style={{ "--fx": "25%", "--fy": "25%", animationDelay: "4.9s" } as CSSProperties}
      />
      <div className="absolute inset-0 z-30 grid place-items-center">
        <div className="fog size-[90vmax] rounded-full bg-zinc-950 blur-3xl" style={{ animationDelay: "5.1s" }} />
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
      <div className={`relative min-h-[260px] flex-1 overflow-hidden rounded-[2rem] border border-white/10 p-6 ${cardClass}`}>
        <span className={`inline-block rounded-full px-2.5 py-1 text-xs font-medium ${chipClass}`}>{idx}</span>
        {children}
        <h3 className={`absolute bottom-6 left-6 text-xl font-medium leading-snug text-white ${titleClass}`}>{title}</h3>
      </div>
      {/* penjelasan: tanpa card */}
      <p className="mt-4 max-w-md px-1 text-sm leading-relaxed text-neutral-400">{desc}</p>
    </div>
  );
}

function Bento() {
  return (
    <section id="fitur" className="mx-auto max-w-6xl px-4 py-24">
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        {/* judul section: teks biasa, bukan card */}
        <div className="col-span-2 flex flex-col justify-end pb-4 lg:col-start-3 lg:row-start-1">
          <Eyebrow>Fitur</Eyebrow>
          <h2 className="mt-2 text-4xl font-normal leading-[1.1] tracking-tight text-white sm:text-5xl">
            Satu alur dari skripsi sampai <span className="rainbow-text">sidang</span>
          </h2>
          <p className="mt-3 max-w-md text-sm leading-relaxed text-neutral-400">
            Format skripsi, buat slide, dan tetap pegang kendali atas isinya.
          </p>
        </div>

        {/* 01 Formatter */}
        <FeatureCell
          idx="01"
          title="AI Thesis Formatter"
          desc="Margin, font, spasi, heading, daftar isi, dan daftar pustaka diterapkan konsisten."
          className="col-span-2 lg:col-span-2 lg:col-start-1 lg:row-span-2 lg:row-start-1"
          cardClass="bg-[#17132b] text-white"
          chipClass="bg-violet-400 text-zinc-950"
        >
          <div className="absolute left-6 top-16 flex gap-2">
            {[Ruler, Type, ListOrdered, BookMarked].map((Icon, i) => (
              <IconDot
                key={i}
                icon={Icon}
                className={i === 0 ? "bg-violet-400 text-zinc-950" : "bg-white/10 text-white"}
              />
            ))}
          </div>
          <div className="absolute -bottom-8 -right-3 w-[46%] max-w-[200px] rotate-6">
            <div className="float">
              <DocxIcon />
            </div>
          </div>
        </FeatureCell>

        {/* ikon saja: buku */}
        <div className="col-span-1 grid min-h-[190px] place-items-center rounded-[2rem] border border-white/10 bg-zinc-900 lg:col-start-3 lg:row-start-2">
          <div className="float">
            <BooksIcon className="w-24 sm:w-28" />
          </div>
        </div>

        {/* ikon saja: orbit (border warna-warni bergerak) */}
        <div className="rb-border col-span-1 grid min-h-[190px] place-items-center rounded-[2rem] bg-zinc-950 lg:col-start-4 lg:row-start-2">
          <OrbitIcon className="w-28 sm:w-32" />
        </div>

        {/* 02 Slides */}
        <FeatureCell
          idx="02"
          title="Skripsi to Slides"
          desc="Alur presentasi 5 sampai 40 slide dengan poin kunci dan catatan presenter."
          className="col-span-2 lg:col-span-3"
          cardClass="bg-[#0b1d29] text-white"
          chipClass="bg-cyan-400 text-zinc-950"
        >
          <span className="rainbow-bg absolute right-10 top-8 size-24 rounded-full opacity-90 sm:right-24 sm:size-32" />
          <div className="absolute left-6 top-16 flex gap-2">
            {[Presentation, StickyNote, LayoutGrid].map((Icon, i) => (
              <IconDot
                key={i}
                icon={Icon}
                className={i === 0 ? "bg-cyan-400 text-zinc-950" : "bg-white/10 text-white"}
              />
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
          cardClass="bg-[#0f2219] text-white"
          chipClass="bg-emerald-400 text-zinc-950"
          titleClass="max-w-[80%]"
        >
          <span className="absolute right-5 top-5 grid size-14 place-items-center rounded-full bg-zinc-950 text-emerald-300">
            <ShieldCheck className="size-6" />
          </span>
          <div className="absolute left-6 top-16 flex gap-2">
            {[Lock, EyeOff].map((Icon, i) => (
              <IconDot key={i} icon={Icon} className="bg-white/10 text-emerald-200" />
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
    <main className="bg-zinc-950 text-neutral-300">
      <style dangerouslySetInnerHTML={{ __html: css }} />

      {/* selalu mulai dari atas + kunci scroll selama intro */}
      <ScrollReset />

      {/* intro (hanya saat awal masuk) */}
      <Intro />

      {/* ================================ HERO ================================ */}
      <section className="relative isolate overflow-hidden bg-zinc-950">
        {/* pendar warna-warni di latar */}
        <div
          className="pointer-events-none absolute inset-0 z-0"
          aria-hidden="true"
          style={{
            background:
              "radial-gradient(60% 50% at 18% 18%, rgba(129,140,248,.20), transparent 70%), radial-gradient(50% 45% at 85% 12%, rgba(244,114,182,.18), transparent 70%), radial-gradient(45% 40% at 62% 72%, rgba(34,211,238,.14), transparent 70%)",
          }}
        />

        {/* lingkaran (matahari warna-warni) + ripple */}
        <div className="pointer-events-none absolute inset-x-0 bottom-0 z-0 flex justify-center" aria-hidden="true">
          <div
            className="sun relative -mb-[100px] size-[440px] sm:-mb-[200px] sm:size-[640px]"
            style={{ animationDelay: "6.4s" }}
          >
            <div
              className="spin-slow absolute inset-0 rounded-full"
              style={{
                background:
                  "conic-gradient(from 0deg, #ff5c8a, #ffb84d, #4ade80, #22d3ee, #818cf8, #e879f9, #ff5c8a)",
                boxShadow: "0 0 140px rgba(129,140,248,.45)",
                animationDuration: "24s",
              }}
            />
            <span className="ripple absolute inset-0 rounded-full border-2 border-fuchsia-400" style={{ animationDelay: "8.3s" }} />
            <span className="ripple absolute inset-0 rounded-full border-2 border-cyan-400" style={{ animationDelay: "11.3s" }} />
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
          <p className="fade-up text-sm text-neutral-400" style={{ animationDelay: "7.2s" }}>
            Untuk mahasiswa tingkat akhir
          </p>

          <h1
            className="fade-up mt-3 max-w-3xl text-4xl font-semibold leading-[1.1] tracking-tight text-white sm:text-6xl"
            style={{ animationDelay: "7.4s" }}
          >
            Skripsi rapi dan presentasi siap, <span className="rainbow-text">tanpa begadang.</span>
          </h1>

          <p
            className="fade-up mt-5 max-w-xl text-base leading-relaxed text-neutral-300"
            style={{ animationDelay: "7.6s" }}
          >
            Unggah skripsimu, biarkan AI memahami strukturnya, lalu unduh DOCX yang sesuai standar dan slide sidang yang siap dipresentasikan.
          </p>

          <div
            className="fade-up mt-8 flex flex-col items-center justify-center gap-3 sm:flex-row"
            style={{ animationDelay: "7.8s" }}
          >
            <Button
              asChild
              className="rainbow-bg h-12 rounded-full pl-2 pr-6 text-sm font-medium text-zinc-950 hover:brightness-110"
            >
              <Link href="/formatter">
                <span className="grid size-8 place-items-center rounded-full bg-zinc-950/85 text-white">
                  <FileCheck2 className="size-4" />
                </span>
                Mulai format skripsi
              </Link>
            </Button>
            <Button
              asChild
              variant="outline"
              className="rb-border h-12 rounded-full border-0 bg-zinc-950/70 pl-2 pr-6 text-sm text-white backdrop-blur hover:bg-zinc-900 hover:text-white"
            >
              <Link href="/slides">
                <span className="grid size-8 place-items-center rounded-full bg-white/10">
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
        <div className="flex flex-col items-center text-center">
          <Eyebrow>Cara kerja</Eyebrow>
          <h2 className="mt-2 text-3xl font-normal tracking-tight text-white sm:text-4xl">
            Tiga langkah <span className="rainbow-text">saja</span>
          </h2>
        </div>

        <div className="relative mt-14 grid gap-10 md:grid-cols-3">
          <div className="absolute left-[17%] right-[17%] top-8 hidden border-t border-dashed border-white/20 md:block" aria-hidden="true" />
          {steps.map(({ icon: Icon, n, title, desc, tile }) => (
            <div key={n} className="relative text-center">
              <div className={`relative mx-auto grid size-16 place-items-center rounded-full ${tile}`}>
                <Icon className="size-6" />
                <span className="absolute -right-1 -top-1 grid size-6 place-items-center rounded-full border border-white/25 bg-zinc-950 text-[10px] font-medium text-white">
                  {n}
                </span>
              </div>
              <h3 className="mt-5 font-medium text-white">{title}</h3>
              <p className="mx-auto mt-1 max-w-[16rem] text-sm text-neutral-400">{desc}</p>
            </div>
          ))}
        </div>
      </section>

      {/* ============================ TUTORIAL SINGKAT ======================== */}
      <section id="tutorial" className="mx-auto max-w-5xl px-4 pb-24 pt-8">
        <div className="max-w-xl">
          <Eyebrow>Tutorial</Eyebrow>
          <h2 className="mt-2 text-3xl font-normal tracking-tight text-white sm:text-4xl">Tutorial singkat</h2>
          <p className="mt-2 text-sm text-neutral-400">Dua alur utama. Ikuti langkahnya satu per satu.</p>
        </div>

        <div className="mt-10 grid gap-12 md:grid-cols-2 md:gap-16">
          {tutorials.map((t) => (
            <div key={t.title}>
              <div className="flex items-center gap-3">
                <span className={`rounded-full px-2.5 py-1 text-xs font-medium ${t.tag}`}>{t.format}</span>
                <h3 className="text-lg font-medium text-neutral-50">{t.title}</h3>
              </div>

              <ol className="mt-6">
                {t.steps.map(([title, desc], i) => (
                  <li
                    key={title}
                    className={`relative pb-6 pl-11 last:pb-0 before:absolute before:bottom-1 before:left-3.5 before:top-9 before:w-px last:before:hidden ${t.line}`}
                  >
                    <span className={`absolute left-0 top-0 grid size-7 place-items-center rounded-full text-xs font-semibold text-zinc-950 ${t.dot}`}>
                      {i + 1}
                    </span>
                    <h4 className="font-medium text-white">{title}</h4>
                    <p className="mt-1 text-sm leading-relaxed text-neutral-400">{desc}</p>
                  </li>
                ))}
              </ol>

              <Button
                asChild
                variant="outline"
                className={`mt-8 rounded-full bg-transparent px-6 ${t.btn}`}
              >
                <Link href={t.href}>{t.cta}</Link>
              </Button>
            </div>
          ))}
        </div>

        <div className="mt-14 grid gap-3 rounded-[2rem] border border-white/10 bg-zinc-900 p-6 text-sm text-neutral-300 sm:grid-cols-3">
          {tips.map((tip, i) => (
            <p key={tip} className="flex items-start gap-2">
              <Check className={`mt-0.5 size-4 shrink-0 ${tipColors[i % tipColors.length]}`} />
              {tip}
            </p>
          ))}
        </div>
      </section>

      {/* ================================= FAQ ================================ */}
      <section id="faq" className="mx-auto max-w-5xl px-4 pb-24">
        <div className="grid gap-8 md:grid-cols-[1fr_2fr] md:gap-12">
          <div>
            <Eyebrow>FAQ</Eyebrow>
            <h2 className="mt-2 text-3xl font-normal tracking-tight text-white sm:text-4xl">Pertanyaan umum</h2>
          </div>
          <div className="divide-y divide-zinc-800 border-y border-zinc-800">
            {faqs.map(([q, a]) => (
              <details key={q} className="group py-4">
                <summary className="flex cursor-pointer list-none items-center justify-between gap-4 font-medium text-white transition-colors hover:text-pink-300 group-open:text-pink-300 [&::-webkit-details-marker]:hidden">
                  {q}
                  <Plus className="size-4 shrink-0 text-violet-400 transition-transform group-open:rotate-45" />
                </summary>
                <p className="mt-2 pr-8 text-sm leading-relaxed text-neutral-400">{a}</p>
              </details>
            ))}
          </div>
        </div>
      </section>

      {/* ================================= CTA ================================ */}
      <section className="mx-auto max-w-6xl px-4 pb-20">
        <div className="rb-border relative overflow-hidden rounded-[2rem] bg-zinc-900 p-8 text-white sm:p-12">
          <span className="rainbow-bg absolute -bottom-24 -right-16 size-64 rounded-full" aria-hidden="true" />
          <span className="ripple absolute -bottom-24 -right-16 size-64 rounded-full border-2 border-fuchsia-300" aria-hidden="true" />
          <span className="absolute right-44 top-10 hidden size-5 rounded-full bg-pink-400 sm:block" aria-hidden="true" />
          <span className="absolute bottom-10 right-72 hidden size-3 rounded-full bg-cyan-400 sm:block" aria-hidden="true" />
          <div className="relative max-w-lg">
            <h2 className="text-3xl font-normal tracking-tight">
              Siap merapikan <span className="rainbow-text">skripsimu?</span>
            </h2>
            <p className="mt-3 text-sm text-neutral-300">Mulai dari format DOCX, lanjut ke slide sidang dalam satu alur.</p>
            <div className="mt-8 flex flex-col gap-3 sm:flex-row">
              <Button asChild className="rainbow-bg rounded-full px-6 font-medium text-zinc-950 hover:brightness-110">
                <Link href="/formatter">Mulai format skripsi</Link>
              </Button>
              <Button
                asChild
                variant="outline"
                className="rounded-full border-white/20 bg-transparent px-6 text-white hover:border-white/40 hover:bg-white/10 hover:text-white"
              >
                <Link href="/slides">Buat presentasi</Link>
              </Button>
            </div>
          </div>
        </div>
      </section>

      <footer className="flex items-center justify-center gap-2 border-t border-zinc-800 py-6 text-center text-xs text-neutral-500">
        <span className="rainbow-bg size-1.5 rounded-full" aria-hidden="true" />
        Periksa kembali hasil akhir sebelum dikumpulkan.
      </footer>
    </main>
  );
}
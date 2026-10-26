"use client";
import { useCallback, useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { Controller, useForm } from "react-hook-form";
import type { Control, FieldErrors, Path } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import {
  BookMarked, Check, ChevronDown, Hash, Heading, Library, List, Minus, Plus, Ruler, Table, Wand2,
} from "lucide-react";
import type { LucideIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { FileDropzone } from "@/features/file-upload/components/file-dropzone";
import { formatConfigSchema, defaultFormatConfig, configForPreset, type FormatConfig } from "../schemas/format-config";
import { PRESETS, presetById } from "../presets";
import { useFormatThesis } from "../hooks/use-format-thesis";
import { ResultView } from "./result-view";

/* ───────────────────────── Tipe & helper ───────────────────────── */

type View = "input" | "loading" | "result";
type Ctl = Control<FormatConfig>;
type Name = Path<FormatConfig>;
type Opt = { value: string; label: string; hint?: string };

type Item =
  | { kind: "choice"; name: Name; label: string; options: Opt[]; hint?: string }
  | { kind: "number"; name: Name; label: string; unit: string; step: number; hint?: string }
  | { kind: "toggle"; name: Name; label: string; hint?: string };

type Group = {
  id: string;
  title: string;
  desc: string;
  Icon: LucideIcon;
  blocks: { title?: string; items: Item[] }[];
};

const cx = (...c: (string | false | null | undefined)[]) => c.filter(Boolean).join(" ");
const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
const o = (value: string, label: string, hint?: string): Opt => ({ value, label, hint });
const choice = (name: Name, label: string, options: Opt[], hint?: string): Item => ({ kind: "choice", name, label, options, hint });
const num = (name: Name, label: string, unit: string, step: number, hint?: string): Item => ({ kind: "number", name, label, unit, step, hint });
const sw = (name: Name, label: string, hint?: string): Item => ({ kind: "toggle", name, label, hint });

/* ───────────────────────── Pengaturan waktu transisi ───────────────────────── */

const SLIDE_MS = 1100; // lama satu lembar bergerak
const LAYER_STAGGER_MS = 110; // jeda antar lembar
const MAX_DELAY_MS = 2 * LAYER_STAGGER_MS; // 3 lembar → 2 jeda
const CLOSE_MS = SLIDE_MS + MAX_DELAY_MS + 60; // sampai layar benar-benar tertutup
const HOLD_MS = 150; // layar tertutup, UI diganti
const OPEN_MS = SLIDE_MS + MAX_DELAY_MS + 60; // sampai tirai benar-benar terbuka
const MIN_LOADING_MS = 2500; // loading minimal terlihat

const STAGES = [
  "Membaca struktur dokumen…",
  "AI menganalisis bagian yang ambigu…",
  "Menerapkan format…",
  "Memeriksa keutuhan isi…",
];

/* ───────────────────────── Daftar pengaturan (per kelompok) ───────────────────────── */

const GROUPS: Group[] = [
  {
    id: "dasar",
    title: "Kertas & teks isi",
    desc: "Ukuran kertas, font, margin, dan spasi",
    Icon: Ruler,
    blocks: [
      {
        title: "Kertas dan font",
        items: [
          choice("paper", "Ukuran kertas", [o("A4", "A4"), o("Letter", "Letter")]),
          choice("fontFamily", "Jenis font", [o("Times New Roman", "Times New Roman"), o("Arial", "Arial"), o("Calibri", "Calibri")]),
        ],
      },
      {
        title: "Margin",
        items: [
          num("marginTop", "Margin atas", "cm", 0.1),
          num("marginBottom", "Margin bawah", "cm", 0.1),
          num("marginLeft", "Margin kiri", "cm", 0.1),
          num("marginRight", "Margin kanan", "cm", 0.1),
        ],
      },
      {
        title: "Isi dan paragraf",
        items: [
          num("fontSize", "Ukuran font", "pt", 1),
          num("lineSpacing", "Spasi baris", "×", 0.1, "Contoh: 1,5 atau 2"),
          choice("lineRule", "Jenis spasi", [
            o("auto", "Kelipatan", "Spasi berupa kelipatan baris, misalnya 1,5 atau 2."),
            o("exact", "Pasti", "Tinggi tiap baris dipatok dalam satuan pt."),
          ]),
          num("indent", "Indentasi paragraf", "cm", 0.01, "Jarak awal paragraf dari tepi kiri"),
        ],
      },
    ],
  },
  {
    id: "sampul",
    title: "Sampul",
    desc: "Ukuran dan gaya tulisan di halaman judul",
    Icon: BookMarked,
    blocks: [
      {
        title: "Ukuran huruf",
        items: [num("coverTitleSize", "Font judul sampul", "pt", 1), num("coverInstitutionSize", "Font identitas kampus", "pt", 1)],
      },
      {
        title: "Huruf kapital dan tebal",
        items: [
          sw("coverTitleUppercase", "Judul sampul huruf kapital"),
          sw("coverNameUppercase", "Nama penulis huruf kapital"),
          sw("coverInstitutionUppercase", "Identitas kampus huruf kapital"),
          sw("coverInstitutionBold", "Identitas kampus tebal"),
          sw("coverIdentityBold", "Nama & NPM tebal"),
        ],
      },
      {
        title: "Teks tambahan",
        items: [
          sw("coverAuthorLabel", "Label “Disusun oleh”", "Tulisan ini muncul di atas nama penulis"),
          sw("coverTranslationSameSize", "Judul terjemahan sama besar"),
        ],
      },
    ],
  },
  {
    id: "bab",
    title: "Judul bab",
    desc: "Penulisan nomor bab dan jarak judul",
    Icon: Heading,
    blocks: [
      {
        items: [
          choice("chapterLabel", "Penulisan nomor bab", [
            o("bab", "BAB I", "Tampil: BAB I, lalu judul di baris bawahnya."),
            o("roman", "I saja", "Tampil: I, lalu judul di baris bawahnya."),
          ]),
          choice("headingLineSpacing", "Spasi judul", [o("single", "Tunggal"), o("body", "Sama dengan isi")]),
          num("chapterSpaceAfter", "Jarak setelah judul bab", "baris", 0.5),
          sw("chapterUppercase", "Judul bab huruf kapital"),
        ],
      },
    ],
  },
  {
    id: "nomor",
    title: "Nomor halaman",
    desc: "Posisi dan awal penomoran",
    Icon: Hash,
    blocks: [
      {
        items: [
          choice("pageNumbers", "Posisi nomor halaman", [
            o("bottom-center", "Bawah tengah"),
            o("bottom-right", "Bawah kanan"),
            o("keep", "Ikuti dokumen asli"),
            o("none", "Tanpa nomor"),
          ]),
          choice("frontNumberingFrom", "Penomoran halaman awal mulai dari", [
            o("first", "Setelah sampul"),
            o("preface", "Kata Pengantar"),
          ]),
          sw("romanFrontMatter", "Romawi untuk bagian awal", "Kata pengantar, daftar isi, dan sejenisnya memakai i, ii, iii"),
        ],
      },
    ],
  },
  {
    id: "tabel",
    title: "Tabel & gambar",
    desc: "Tampilan tabel, judul tabel, dan judul gambar",
    Icon: Table,
    blocks: [
      {
        title: "Tabel",
        items: [
          num("tableFontSize", "Font isi tabel", "pt", 1),
          choice("tableBorders", "Garis tabel", [o("all", "Garis penuh"), o("minimal", "Atas & bawah saja"), o("keep", "Ikuti dokumen asli")]),
          choice("tableNumericAlign", "Posisi angka di tabel", [o("center", "Rata tengah"), o("right", "Rata kanan")]),
          sw("normalizeTables", "Rapikan tata letak tabel"),
        ],
      },
      {
        title: "Judul tabel dan gambar",
        items: [
          num("captionFontSize", "Font judul tabel/gambar", "pt", 1),
          choice("tableCaptionAlign", "Posisi judul tabel", [o("left", "Rata kiri"), o("center", "Rata tengah")]),
          choice("figureCaptionAlign", "Posisi judul gambar", [o("center", "Rata tengah"), o("left", "Rata kiri")]),
          choice("captionNumbering", "Cara penomoran", [
            o("chapter", "Per bab", "Contoh: Tabel 2.1"),
            o("global", "Berurutan", "Contoh: Tabel 1, Tabel 2, dst."),
          ]),
          choice("captionSeparator", "Tanda setelah nomor", [
            o("space", "Spasi", "Contoh: Tabel 1 Judul"),
            o("period", "Titik", "Contoh: Tabel 1. Judul"),
          ]),
          choice("figureLabel", "Sebutan untuk gambar", [o("auto", "Ikuti dokumen"), o("Gambar", "Gambar"), o("Ilustrasi", "Ilustrasi")]),
        ],
      },
    ],
  },
  {
    id: "daftar",
    title: "Daftar isi & daftar lain",
    desc: "Daftar isi, daftar butir, dan penandanya",
    Icon: List,
    blocks: [
      {
        items: [
          choice("tocMode", "Daftar isi yang sudah ada", [
            o("keep", "Pertahankan", "Daftar isi lama di dokumen dibiarkan."),
            o("rebuild", "Ganti otomatis", "Daftar isi lama diganti daftar isi baru."),
          ]),
          choice("listStyle", "Bentuk daftar isi", [
            o("word", "Otomatis Word", "Memakai fitur daftar isi bawaan Word."),
            o("columns", "Kolom", "Tersusun dalam kolom: nomor | judul | halaman."),
          ]),
          choice("listMarkers", "Penanda butir daftar", [
            o("uniform", "Diseragamkan"),
            o("keep", "Ikuti dokumen", "Penanda asli seperti 1), a., dan seterusnya dipertahankan."),
          ]),
          choice("listLayout", "Daftar bertingkat", [
            o("nested", "Menjorok per tingkat"),
            o("flat", "Rata di margin", "Semua penanda sejajar di tepi margin."),
          ]),
        ],
      },
    ],
  },
  {
    id: "pustaka",
    title: "Daftar pustaka & lainnya",
    desc: "Gaya sitasi, abstrak, dan perapian akhir",
    Icon: Library,
    blocks: [
      {
        items: [
          choice("citationStyle", "Format sitasi", [o("APA", "APA"), o("IEEE", "IEEE"), o("Harvard", "Harvard"), o("Vancouver", "Vancouver")]),
          num("referenceHanging", "Gantungan daftar pustaka", "cm", 0.01, "Baris kedua tiap pustaka menjorok ke dalam"),
          sw("referenceGap", "Baris kosong antar pustaka"),
          sw("abstractIndent", "Abstrak memakai indentasi"),
          sw("cleanupBlankPages", "Bersihkan halaman kosong"),
        ],
      },
    ],
  },
];

/* ───────────────────────── Transisi tirai kain (ringan) ─────────────────────────
   Tiga lembar kain naik dari bawah menutup layar, lalu terus naik ke atas saat
   terbuka. Semua gerak murni transform CSS (dikerjakan GPU), tanpa hitung per frame
   dan tanpa filter/bayangan, jadi tetap mulus di HP.
   Tepi atas dan bawah tiap lembar melengkung (SVG statis yang ikut bergerak).
   Oranye memimpin sebagai aksen, hitam menyusul. Saat terbuka urutannya dibalik. */

const EASE = "cubic-bezier(0.65, 0, 0.35, 1)";
const EDGE = 120; // tinggi lengkung di tepi atas/bawah lembar (px)

const SHEETS = [
  { color: "#fb923c" }, // oranye · aksen, memimpin saat menutup
  { color: "#292524" }, // abu gelap hangat
  { color: "#0a0a0a" }, // hitam pekat · lapis depan, menutup terakhir
] as const;

type Phase = "idle" | "in" | "out";

// lengkung tepi: tengah sedikit lebih maju dari sisi kiri-kanan
function Edge({ position, fill }: { position: "top" | "bottom"; fill: string }) {
  return (
    <svg
      viewBox="0 0 1600 120"
      preserveAspectRatio="none"
      className="absolute left-0 w-full"
      style={{
        height: EDGE,
        ...(position === "top"
          ? { bottom: "calc(100% - 1px)" }
          : { top: "calc(100% - 1px)", transform: "scaleY(-1)" }),
      }}
    >
      <path d="M0 120 V80 Q800 -20 1600 80 V120 Z" fill={fill} />
    </svg>
  );
}

function CurtainTransition({ covered }: { covered: boolean }) {
  const [phase, setPhase] = useState<Phase>("idle");
  const was = useRef(false);

  useEffect(() => {
    if (covered === was.current) return;
    was.current = covered;
    if (covered) {
      setPhase("in");
      return;
    }
    setPhase("out");
    // setelah benar-benar keluar layar, kembalikan ke bawah tanpa animasi
    const t = setTimeout(() => setPhase("idle"), SLIDE_MS + MAX_DELAY_MS + 30);
    return () => clearTimeout(t);
  }, [covered]);

  const pos =
    phase === "in"
      ? "translate3d(0, 0, 0)"
      : phase === "out"
        ? `translate3d(0, calc(-100% - ${EDGE}px), 0)`
        : `translate3d(0, calc(100% + ${EDGE}px), 0)`;

  return (
    <div
      aria-hidden="true"
      className={cx("fixed inset-0 z-[100] overflow-hidden motion-reduce:hidden", covered ? "pointer-events-auto" : "pointer-events-none")}
      style={{ visibility: phase === "idle" ? "hidden" : "visible" }}
    >
      {SHEETS.map((sheet, i) => {
        // menutup: oranye duluan · membuka: hitam duluan
        const order = phase === "out" ? SHEETS.length - 1 - i : i;
        const delay = order * LAYER_STAGGER_MS;
        return (
          <div
            key={sheet.color}
            className="absolute inset-0"
            style={{
              backgroundColor: sheet.color,
              transform: pos,
              transition: phase === "idle" ? "none" : `transform ${SLIDE_MS}ms ${EASE} ${delay}ms`,
              willChange: "transform",
            }}
          >
            <Edge position="top" fill={sheet.color} />
            <Edge position="bottom" fill={sheet.color} />
          </div>
        );
      })}
    </div>
  );
}

/* ───────────────────────── Layar loading ───────────────────────── */

function LoadingScreen({ stages }: { stages: string[] }) {
  const [i, setI] = useState(0);

  useEffect(() => {
    const t = setInterval(() => setI((v) => Math.min(v + 1, stages.length - 1)), 3500);
    return () => clearInterval(t);
  }, [stages.length]);

  const last = i === stages.length - 1;

  return (
    <div
      role="status"
      aria-live="polite"
      className="mx-auto flex min-h-screen max-w-md flex-col items-center justify-center px-4 py-16 text-center"
    >
      <svg viewBox="0 0 200 220" className="w-44" aria-hidden="true">
        <defs>
          <clipPath id="ftw-doc"><rect x="52" y="40" width="96" height="124" rx="12" /></clipPath>
        </defs>
        <circle cx="100" cy="104" r="84" fill="#fed7aa" />
        <circle cx="168" cy="40" r="12" fill="#fb923c" />
        <circle cx="30" cy="172" r="7" fill="#fdba74" />
        <rect x="52" y="40" width="96" height="124" rx="12" fill="#ffffff" stroke="#171717" strokeWidth="2.5" />
        <g clipPath="url(#ftw-doc)">
          <rect x="68" y="58" width="48" height="7" rx="3.5" fill="#171717" />
          {[78, 92, 106, 120, 134].map((y, k) => (
            <rect key={y} x="68" y={y} width={k % 2 ? 52 : 64} height="5" rx="2.5" fill="#d4d4d4" />
          ))}
          <g>
            <animateTransform
              attributeName="transform"
              type="translate"
              values="0 0; 0 108; 0 0"
              keyTimes="0; 0.5; 1"
              calcMode="spline"
              keySplines="0.45 0 0.55 1; 0.45 0 0.55 1"
              dur="2.8s"
              repeatCount="indefinite"
            />
            <rect x="52" y="44" width="96" height="16" fill="#f97316" fillOpacity="0.18" />
            <rect x="52" y="44" width="96" height="2.5" fill="#f97316" />
          </g>
        </g>
      </svg>

      <h2 className="mt-6 text-2xl font-normal tracking-tight text-neutral-900">Sedang memformat dokumenmu</h2>
      <p className="mt-2 min-h-5 text-sm text-neutral-500">
        {last ? "Hampir selesai. Dokumen yang panjang butuh waktu lebih lama." : "Mohon tunggu, jangan tutup halaman ini."}
      </p>

      <svg width="192" height="4" className="mt-6" aria-hidden="true">
        <rect width="192" height="4" rx="2" fill="#f5f5f5" />
        <rect width="64" height="4" rx="2" fill="#f97316">
          <animate attributeName="x" values="-64;192" dur="1.6s" repeatCount="indefinite" />
        </rect>
      </svg>

      <ol className="mt-8 w-full max-w-xs space-y-3 text-left">
        {stages.map((s, k) => {
          const done = k < i;
          const active = k === i;
          return (
            <li
              key={s}
              className={cx("flex items-center gap-3 text-sm", done ? "text-neutral-500" : active ? "text-neutral-900" : "text-neutral-300")}
            >
              <span
                className={cx(
                  "grid size-5 shrink-0 place-items-center rounded-full border bg-white",
                  done ? "border-orange-400 bg-orange-400 text-white" : active ? "border-orange-400" : "border-neutral-200",
                )}
              >
                {done ? <Check className="size-3" /> : active ? <span className="size-2 animate-pulse rounded-full bg-orange-400" /> : null}
              </span>
              {s}
            </li>
          );
        })}
      </ol>
    </div>
  );
}

/* ───────────────────────── Komponen input custom ───────────────────────── */

function Setting({ label, hint, error, children }: { label: string; hint?: string; error?: string; children: ReactNode }) {
  return (
    <div className="space-y-2">
      <div>
        <p className="text-sm font-medium text-neutral-900">{label}</p>
        {hint && <p className="mt-0.5 text-xs leading-relaxed text-neutral-500">{hint}</p>}
      </div>
      {children}
      {error && <p className="text-xs text-red-600">{error}</p>}
    </div>
  );
}

/** Pilihan berbentuk tombol pil (pengganti <select> kecil) */
function Choice({ control, item }: { control: Ctl; item: Extract<Item, { kind: "choice" }> }) {
  return (
    <Controller
      control={control}
      name={item.name}
      render={({ field, fieldState }) => {
        const value = String(field.value ?? "");
        const current = item.options.find((x) => x.value === value);
        return (
          <Setting label={item.label} hint={item.hint} error={fieldState.error?.message}>
            <div role="radiogroup" aria-label={item.label} className="flex flex-wrap gap-2">
              {item.options.map((opt) => {
                const on = opt.value === value;
                return (
                  <button
                    key={opt.value}
                    type="button"
                    role="radio"
                    aria-checked={on}
                    onClick={() => field.onChange(opt.value)}
                    className={cx(
                      "inline-flex items-center gap-1.5 rounded-full px-4 py-2 text-sm transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400",
                      on ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-700 hover:bg-neutral-200",
                    )}
                  >
                    {on && <Check className="size-3.5 text-orange-300" />}
                    {opt.label}
                  </button>
                );
              })}
            </div>
            {current?.hint && <p className="border-l-2 border-orange-300 pl-3 text-xs leading-relaxed text-neutral-600">{current.hint}</p>}
          </Setting>
        );
      }}
    />
  );
}

/** Angka dengan tombol − / + (pengganti <input type="number">) */
function Stepper({ value, onChange, step, unit }: { value: number; onChange: (v: number) => void; step: number; unit: string }) {
  const [draft, setDraft] = useState<string | null>(null);
  const decimals = (String(step).split(".")[1] ?? "").length;
  const base = Number.isFinite(value) ? value : 0;
  const shown = draft ?? (Number.isFinite(value) ? String(value).replace(".", ",") : "");

  const bump = (dir: 1 | -1) => {
    setDraft(null);
    onChange(Math.max(0, Number((base + dir * step).toFixed(decimals + 2))));
  };

  const btn =
    "grid size-8 shrink-0 place-items-center rounded-full bg-white text-neutral-700 transition hover:bg-neutral-900 hover:text-white focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400";

  return (
    <div className="flex items-center gap-1 rounded-full bg-neutral-100 p-1 focus-within:ring-2 focus-within:ring-orange-400">
      <button type="button" aria-label="Kurangi" onClick={() => bump(-1)} className={btn}>
        <Minus className="size-3.5" />
      </button>
      <div className="flex min-w-0 flex-1 items-baseline justify-center gap-1">
        <input
          type="text"
          inputMode="decimal"
          value={shown}
          onChange={(e) => {
            const t = e.target.value;
            if (!/^[0-9]*[.,]?[0-9]*$/.test(t)) return;
            setDraft(t);
            const n = parseFloat(t.replace(",", "."));
            onChange(Number.isNaN(n) ? NaN : n);
          }}
          onBlur={() => setDraft(null)}
          className="w-14 bg-transparent text-center text-sm text-neutral-900 outline-none"
        />
        <span className="text-xs text-neutral-500">{unit}</span>
      </div>
      <button type="button" aria-label="Tambah" onClick={() => bump(1)} className={btn}>
        <Plus className="size-3.5" />
      </button>
    </div>
  );
}

function NumberField({ control, item }: { control: Ctl; item: Extract<Item, { kind: "number" }> }) {
  return (
    <Controller
      control={control}
      name={item.name}
      render={({ field, fieldState }) => (
        <Setting label={item.label} hint={item.hint} error={fieldState.error?.message}>
          <Stepper value={Number(field.value)} onChange={field.onChange} step={item.step} unit={item.unit} />
        </Setting>
      )}
    />
  );
}

/** Saklar on/off (pengganti checkbox) */
function ToggleField({ control, item }: { control: Ctl; item: Extract<Item, { kind: "toggle" }> }) {
  return (
    <Controller
      control={control}
      name={item.name}
      render={({ field }) => {
        const on = Boolean(field.value);
        return (
          <button
            type="button"
            role="switch"
            aria-checked={on}
            onClick={() => field.onChange(!on)}
            className="flex h-full w-full items-center justify-between gap-4 border-b border-neutral-100 py-3 text-left transition hover:border-neutral-300 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400"
          >
            <span className="min-w-0">
              <span className="block text-sm text-neutral-900">{item.label}</span>
              {item.hint && <span className="mt-0.5 block text-xs leading-relaxed text-neutral-500">{item.hint}</span>}
            </span>
            <span className={cx("relative h-6 w-11 shrink-0 rounded-full transition-colors", on ? "bg-orange-500" : "bg-neutral-300")}>
              <span className={cx("absolute top-0.5 size-5 rounded-full bg-white transition-all", on ? "left-[22px]" : "left-0.5")} />
            </span>
          </button>
        );
      }}
    />
  );
}

/** Dropdown custom (pengganti <select> untuk daftar panjang) */
function Dropdown({ value, onChange, options }: { value: string; onChange: (v: string) => void; options: Opt[] }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const current = options.find((x) => x.value === value);

  return (
    <div ref={ref} className="relative">
      <button
        type="button"
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
        className="flex w-full items-center justify-between gap-3 rounded-full bg-neutral-100 px-5 py-3 text-left text-sm text-neutral-900 transition hover:bg-neutral-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400"
      >
        <span className="truncate">{current?.label ?? "Pilih pedoman"}</span>
        <ChevronDown className={cx("size-4 shrink-0 text-neutral-500 transition", open && "rotate-180")} />
      </button>
      {open && (
        <ul role="listbox" className="absolute z-20 mt-2 max-h-72 w-full overflow-auto rounded-2xl border border-neutral-200 bg-white p-1.5 shadow-lg">
          {options.map((opt) => {
            const on = opt.value === value;
            return (
              <li key={opt.value} role="option" aria-selected={on}>
                <button
                  type="button"
                  onClick={() => {
                    onChange(opt.value);
                    setOpen(false);
                  }}
                  className={cx(
                    "flex w-full items-center justify-between gap-3 rounded-xl px-4 py-2.5 text-left text-sm transition",
                    on ? "bg-orange-100 text-neutral-900" : "text-neutral-700 hover:bg-neutral-100",
                  )}
                >
                  <span>{opt.label}</span>
                  {on && <Check className="size-4 text-orange-600" />}
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}

/** Baris yang bisa dibuka/ditutup */
function Section({
  group, open, onToggle, errorCount, children,
}: { group: Group; open: boolean; onToggle: () => void; errorCount: number; children: ReactNode }) {
  const { Icon } = group;
  return (
    <section className="border-b border-neutral-200 last:border-b-0">
      <button
        type="button"
        onClick={onToggle}
        aria-expanded={open}
        className="flex w-full items-center gap-4 py-5 text-left focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400"
      >
        <span
          className={cx(
            "grid size-9 shrink-0 place-items-center rounded-full border bg-white transition-colors",
            open ? "border-orange-400 text-orange-500" : "border-neutral-300 text-neutral-500",
          )}
        >
          <Icon className="size-4" />
        </span>
        <span className="min-w-0 flex-1">
          <span className="block text-base font-medium text-neutral-900">{group.title}</span>
          <span className="block text-xs text-neutral-500">{group.desc}</span>
        </span>
        {errorCount > 0 && <span className="text-xs text-red-600">{errorCount} perlu dicek</span>}
        <ChevronDown className={cx("size-4 shrink-0 text-neutral-400 transition", open && "rotate-180")} />
      </button>
      {open && <div className="space-y-8 pb-8 sm:pl-[3.25rem]">{children}</div>}
    </section>
  );
}

/* ───────────────────────── Halaman utama ───────────────────────── */

export function FormatterWorkspace() {
  const [file, setFile] = useState<File | null>(null);
  const [openId, setOpenId] = useState<string | null>("dasar");
  const [view, setView] = useState<View>("input");
  const [covered, setCovered] = useState(false);

  const mutation = useFormatThesis();
  const { control, handleSubmit, reset, getValues, watch, formState: { errors } } = useForm<FormatConfig>({
    resolver: zodResolver(formatConfigSchema), defaultValues: defaultFormatConfig,
  });

  /* — Orkestrasi transisi: antrean supaya tahap tidak saling tabrak — */
  const chain = useRef<Promise<void>>(Promise.resolve());
  const alive = useRef(true);
  const viewRef = useRef<View>("input");
  const readyAt = useRef(0);

  useEffect(() => {
    alive.current = true;
    return () => { alive.current = false; };
  }, []);

  const go = useCallback((next: View, onSwap?: () => void) => {
    chain.current = chain.current.then(async () => {
      if (!alive.current) return;
      // pastikan loading sempat terlihat sebentar sebelum pindah
      if (viewRef.current === "loading") {
        await sleep(Math.max(0, MIN_LOADING_MS - (Date.now() - readyAt.current)));
        if (!alive.current) return;
      }
      setCovered(true);
      await sleep(CLOSE_MS);
      if (!alive.current) return;
      viewRef.current = next;
      setView(next);
      onSwap?.();
      window.scrollTo(0, 0);
      await sleep(HOLD_MS);
      if (!alive.current) return;
      setCovered(false);
      await sleep(OPEN_MS);
      readyAt.current = Date.now();
    });
  }, []);

  /* — Form — */
  const preset = presetById(watch("presetId"));
  const [paper, fontFamily, fontSize, lineSpacing] = watch(["paper", "fontFamily", "fontSize", "lineSpacing"]);
  const applyPreset = (id: string) => reset({ ...configForPreset(id), docType: getValues("docType") });

  const onSubmit = (config: FormatConfig) => {
    if (!file || mutation.isPending) return;
    go("loading");
    mutation.mutate(
      { file, config },
      { onSuccess: () => go("result"), onError: () => go("input") },
    );
  };

  const onInvalid = (errs: FieldErrors<FormatConfig>) => {
    const g = GROUPS.find((grp) => grp.blocks.some((b) => b.items.some((i) => i.name in errs)));
    if (g) setOpenId(g.id);
  };

  const errorCountOf = (g: Group) =>
    g.blocks.reduce((sum, b) => sum + b.items.filter((i) => (errors as Record<string, unknown>)[i.name]).length, 0);

  const restart = () => {
    setFile(null);
    reset(defaultFormatConfig);
    mutation.reset();
  };

  return (
    <div className="relative min-h-screen bg-white text-neutral-900">
      {/* ── Tahap 1: input ── */}
      {view === "input" && (
        <form onSubmit={handleSubmit(onSubmit, onInvalid)}>
          <div className="mx-auto max-w-6xl px-4 pb-36 pt-12 sm:px-8 sm:pt-16">
            {mutation.isError && (
              <p className="mb-10 border-l-2 border-red-400 pl-4 text-sm leading-relaxed text-red-700">
                {mutation.error.message}
              </p>
            )}

            <div className="grid gap-14 lg:grid-cols-[21rem_1fr]">
              {/* Dropzone */}
              <aside className="space-y-8 lg:sticky lg:top-10 lg:self-start">
                <div>
                  <h1 className="text-3xl font-normal leading-[1.1] tracking-tight sm:text-4xl">AI Thesis Formatter</h1>
                  <p className="mt-3 text-sm leading-relaxed text-neutral-500">
                    Unggah skripsimu, pilih aturan formatnya, lalu biarkan sisanya kami kerjakan.
                  </p>
                </div>

                <FileDropzone file={file} onChange={setFile} />

                <Controller
                  control={control}
                  name="docType"
                  render={({ field }) => (
                    <Setting label="Jenis dokumen">
                      <div role="radiogroup" aria-label="Jenis dokumen" className="flex flex-wrap gap-2">
                        {[o("skripsi", "Skripsi"), o("tesis", "Tesis"), o("proposal", "Proposal")].map((opt) => {
                          const on = field.value === opt.value;
                          return (
                            <button
                              key={opt.value}
                              type="button"
                              role="radio"
                              aria-checked={on}
                              onClick={() => field.onChange(opt.value)}
                              className={cx(
                                "rounded-full px-4 py-2 text-sm transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400",
                                on ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-700 hover:bg-neutral-200",
                              )}
                            >
                              {opt.label}
                            </button>
                          );
                        })}
                      </div>
                    </Setting>
                  )}
                />

                <p className="border-l-2 border-orange-300 pl-4 text-xs leading-relaxed text-neutral-600">
                  Unggah DOCX agar tabel dan gambar ikut terjaga. PDF tidak menyimpan tabel dan gambar
                  sebagai struktur, sehingga dokumennya terpaksa disusun ulang dari teks.
                </p>

                <div className="space-y-2">
                  <p className="text-sm font-medium text-neutral-900">Ringkasan aturan</p>
                  <div className="flex flex-wrap gap-1.5">
                    {[paper, fontFamily, `${fontSize} pt`, `Spasi ${String(lineSpacing).replace(".", ",")}`].map((chip) => (
                      <span key={chip} className="rounded-full bg-neutral-100 px-3 py-1 text-xs text-neutral-700">{chip}</span>
                    ))}
                  </div>
                </div>
              </aside>

              {/* Workspace */}
              <div className="min-w-0 rounded-[2rem] border border-neutral-300 p-6 sm:p-8">
                <div className="space-y-4">
                  <div>
                    <h2 className="text-lg font-medium tracking-tight">Mulai dari pedoman kampus</h2>
                    <p className="mt-1 text-sm leading-relaxed text-neutral-500">
                      Semua aturan di bawah terisi otomatis sesuai pedoman. Kalau kampusmu tidak ada di daftar, pilih Umum.
                    </p>
                  </div>
                  <Controller
                    control={control}
                    name="presetId"
                    render={({ field }) => (
                      <Dropdown
                        value={String(field.value ?? "")}
                        onChange={(id) => applyPreset(id)}
                        options={PRESETS.map((p) => o(p.id, p.name))}
                      />
                    )}
                  />
                  {preset.id !== "umum" && (
                    <div className="space-y-2 border-l-2 border-orange-300 pl-4 text-xs leading-relaxed text-neutral-600">
                      <p className="font-medium text-neutral-900">{preset.source}</p>
                      <ul className="list-disc space-y-1 pl-4">
                        {preset.notes.map((n) => <li key={n}>{n}</li>)}
                      </ul>
                    </div>
                  )}
                </div>

                <div className="mt-8 border-t border-neutral-200 pt-8">
                  <div className="mb-3">
                    <h2 className="text-lg font-medium tracking-tight">Sesuaikan aturan</h2>
                    <p className="mt-1 text-sm text-neutral-500">Buka hanya bagian yang ingin kamu ubah.</p>
                  </div>
                  <div className="border-t border-neutral-200">
                    {GROUPS.map((g) => (
                      <Section
                        key={g.id}
                        group={g}
                        open={openId === g.id}
                        onToggle={() => setOpenId((cur) => (cur === g.id ? null : g.id))}
                        errorCount={errorCountOf(g)}
                      >
                        {g.blocks.map((b, i) => (
                          <div key={i} className={cx("space-y-4", i > 0 && "border-t border-neutral-100 pt-6")}>
                            {b.title && <p className="text-xs font-medium text-neutral-500">{b.title}</p>}
                            <div className="grid gap-x-6 gap-y-5 sm:grid-cols-2">
                              {b.items.map((item) => (
                                <div key={item.name} className={item.kind === "choice" ? "sm:col-span-2" : undefined}>
                                  {item.kind === "choice" && <Choice control={control} item={item} />}
                                  {item.kind === "number" && <NumberField control={control} item={item} />}
                                  {item.kind === "toggle" && <ToggleField control={control} item={item} />}
                                </div>
                              ))}
                            </div>
                          </div>
                        ))}
                      </Section>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Bar aksi tetap di bawah */}
          <div className="fixed inset-x-0 bottom-0 z-30 border-t border-neutral-200 bg-white/90 backdrop-blur">
            <div className="mx-auto flex max-w-6xl items-center justify-between gap-4 px-4 py-3 sm:px-8">
              <p className="hidden min-w-0 truncate text-sm text-neutral-500 sm:block">
                {file ? file.name : "Pilih berkas dulu untuk mulai."}
              </p>
              <Button
                type="submit"
                disabled={!file || mutation.isPending}
                className="h-12 w-full rounded-full bg-neutral-900 pl-2 pr-6 text-sm text-white hover:bg-neutral-800 sm:w-64"
              >
                <span className="mr-2 grid size-8 place-items-center rounded-full bg-orange-400 text-neutral-950">
                  <Wand2 className="size-4" />
                </span>
                Format sekarang
              </Button>
            </div>
          </div>
        </form>
      )}

      {/* ── Tahap 2: loading ── */}
      {view === "loading" && <LoadingScreen stages={STAGES} />}

      {/* ── Tahap 3: hasil ── */}
      {view === "result" && mutation.data && (
        <ResultView
          result={mutation.data}
          onEdit={() => go("input")}
          onRestart={() => go("input", restart)}
        />
      )}

      {/* Transisi tirai (selalu di atas) */}
      <CurtainTransition covered={covered} />
    </div>
  );
}
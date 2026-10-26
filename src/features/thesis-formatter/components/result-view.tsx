"use client";
import { useMemo, useState } from "react";
import type { ReactNode } from "react";
import { AlertTriangle, Check, Download, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import type { FormatResult, Role } from "../types";

const ROLE_LABEL: Record<Role, string> = {
  cover: "Sampul",
  front_heading: "Bagian awal",
  list_heading: "Judul daftar",
  list_entry: "Entri daftar",
  chapter_label: "Label bab",
  chapter: "Bab",
  section: "Subbab",
  subsection: "Anak subbab",
  figure: "Gambar",
  figure_caption: "Keterangan gambar",
  table_caption: "Keterangan tabel",
  body: "Isi",
  reference_heading: "Judul pustaka",
  reference: "Entri pustaka",
  appendix_heading: "Lampiran",
  empty: "Kosong",
};

const ROLE_STYLE: Partial<Record<Role, string>> = {
  chapter: "bg-neutral-900 text-white",
  chapter_label: "bg-neutral-900 text-white",
  front_heading: "bg-orange-300 text-neutral-900",
  list_heading: "bg-orange-300 text-neutral-900",
  reference_heading: "bg-orange-300 text-neutral-900",
  appendix_heading: "bg-orange-300 text-neutral-900",
  section: "bg-orange-200 text-neutral-900",
  subsection: "bg-orange-100 text-neutral-800",
  figure: "bg-neutral-200 text-neutral-800",
  figure_caption: "bg-neutral-100 text-neutral-700",
  table_caption: "bg-neutral-100 text-neutral-700",
  cover: "bg-neutral-200 text-neutral-700",
};

/** Sedikit menjorok supaya hierarki bab → subbab terbaca */
const ROLE_INDENT: Partial<Record<Role, string>> = {
  section: "ml-4",
  subsection: "ml-8",
};

const LEVEL = {
  ok: { border: "border-green-400", text: "text-green-800", Icon: Check, icon: "text-green-600" },
  warn: { border: "border-amber-400", text: "text-amber-800", Icon: AlertTriangle, icon: "text-amber-600" },
  error: { border: "border-red-400", text: "text-red-700", Icon: X, icon: "text-red-600" },
} as const;

type Tab = "ringkasan" | "pemeriksaan" | "urutan";

const TABS: { id: Tab; label: string }[] = [
  { id: "ringkasan", label: "Ringkasan" },
  { id: "pemeriksaan", label: "Pemeriksaan" },
  { id: "urutan", label: "Urutan dokumen" },
];

const cx = (...c: (string | false | null | undefined)[]) => c.filter(Boolean).join(" ");

function download(result: FormatResult) {
  const bin = atob(result.docxBase64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const url = URL.createObjectURL(
    new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = result.fileName;
  a.click();
  URL.revokeObjectURL(url);
}

function SectionTitle({ title, desc }: { title: string; desc?: string }) {
  return (
    <div>
      <h3 className="text-base font-medium text-neutral-900">{title}</h3>
      {desc && <p className="mt-0.5 text-xs leading-relaxed text-neutral-500">{desc}</p>}
    </div>
  );
}

export function ResultView({
  result, onEdit, onRestart,
}: { result: FormatResult; onEdit?: () => void; onRestart?: () => void }) {
  const [tab, setTab] = useState<Tab>("ringkasan");
  const [showAll, setShowAll] = useState(false);
  const preview = useMemo(
    () => (showAll ? result.preview : result.preview.filter((p) => p.role !== "empty").slice(0, 40)),
    [result.preview, showAll],
  );

  const failed = result.integrity.filter((i) => !i.ok).length;
  const notes = result.validation.filter((v) => v.level !== "ok").length;
  const issues = failed + notes;

  const stats = [
    ["Bab", result.stats.chapters],
    ["Subbab", result.stats.sections],
    ["Anak subbab", result.stats.subsections],
    ["Bagian awal", result.stats.frontMatter],
    ["Tabel", result.stats.tables],
    ["Gambar", result.stats.figures],
    ["Pustaka", result.stats.references],
    ["Lampiran", result.stats.appendices],
  ] as const;

  const textBtn =
    "rounded-full px-3 py-1.5 text-sm text-neutral-500 transition hover:text-neutral-900 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400";

  return (
    <div className="mx-auto max-w-5xl space-y-12 px-4 py-14 sm:px-8 sm:py-20">
      {/* Header + unduh */}
      <header className="flex flex-col gap-8 sm:flex-row sm:items-end sm:justify-between">
        <div className="min-w-0 space-y-3">
          <p className="text-sm text-neutral-500">Hasil format</p>
          <h1 className="break-words text-3xl font-normal leading-[1.1] tracking-tight text-neutral-900 sm:text-4xl">
            {result.fileName}
          </h1>
          <p className="max-w-xl text-sm leading-relaxed text-neutral-600">
            {result.mode === "preserve"
              ? "Dokumen asli disunting langsung — struktur, tabel, dan gambar tetap utuh."
              : "Dokumen disusun ulang dari teks PDF."}
          </p>

          {issues === 0 ? (
            <p className="flex items-center gap-2 text-sm text-green-700">
              <Check className="size-4" /> Semua pemeriksaan lolos
            </p>
          ) : (
            <button
              type="button"
              onClick={() => setTab("pemeriksaan")}
              className="flex items-center gap-2 text-sm text-amber-700 underline-offset-4 hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400"
            >
              <AlertTriangle className="size-4" /> {issues} hal perlu dicek — lihat detail
            </button>
          )}
        </div>

        <div className="flex shrink-0 flex-col gap-3 sm:items-end">
          <Button
            onClick={() => download(result)}
            className="h-12 w-full rounded-full bg-neutral-900 pl-2 pr-6 text-sm text-white hover:bg-neutral-800 sm:w-auto"
          >
            <span className="mr-2 grid size-8 place-items-center rounded-full bg-orange-400 text-neutral-950">
              <Download className="size-4" />
            </span>
            Unduh DOCX
          </Button>
          <div className="flex flex-wrap gap-1 sm:justify-end">
            {onEdit && <button type="button" onClick={onEdit} className={textBtn}>Ubah aturan</button>}
            {onRestart && <button type="button" onClick={onRestart} className={textBtn}>Mulai dari awal</button>}
          </div>
        </div>
      </header>

      {/* Tab */}
      <div>
        <div role="tablist" aria-label="Detail hasil" className="flex flex-wrap gap-x-7 border-b border-neutral-200">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={tab === t.id}
              onClick={() => setTab(t.id)}
              className={cx(
                "-mb-px flex items-center gap-2 border-b-2 pb-3 text-sm transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400",
                tab === t.id
                  ? "border-neutral-900 text-neutral-900"
                  : "border-transparent text-neutral-500 hover:text-neutral-900",
              )}
            >
              {t.label}
              {t.id === "pemeriksaan" && issues > 0 && (
                <span className="rounded-full bg-orange-400 px-1.5 text-[11px] text-neutral-950">{issues}</span>
              )}
            </button>
          ))}
        </div>

        <div className="pt-10">
          {/* Tab 1: Ringkasan */}
          {tab === "ringkasan" && (
            <div className="space-y-12">
              <section className="space-y-5">
                <SectionTitle title="Struktur yang terdeteksi" desc="Jumlah bagian yang dikenali dari dokumenmu." />
                <dl className="grid grid-cols-2 gap-x-6 gap-y-8 sm:grid-cols-4">
                  {stats.map(([label, value]) => (
                    <div key={label} className="border-t border-neutral-200 pt-4">
                      <dd className="text-3xl font-normal tabular-nums text-neutral-900">{value}</dd>
                      <dt className="mt-1 text-xs text-neutral-500">{label}</dt>
                    </div>
                  ))}
                </dl>
              </section>

              <div className="grid gap-12 lg:grid-cols-2">
                <section className="space-y-4">
                  <SectionTitle title="Yang sudah diterapkan" />
                  {result.changes.length === 0 ? (
                    <p className="text-sm text-neutral-500">Tidak ada perubahan yang tercatat.</p>
                  ) : (
                    <ul className="space-y-3 text-sm text-neutral-700">
                      {result.changes.map((c, i) => (
                        <li key={i} className="flex gap-3">
                          <Check className="mt-0.5 size-4 shrink-0 text-orange-500" />
                          <span className="leading-relaxed">{c}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </section>

                <section className="space-y-4">
                  <SectionTitle title="Saran" desc="Hal yang sebaiknya kamu cek sendiri di Word." />
                  {result.recommendations.length === 0 ? (
                    <p className="text-sm text-neutral-500">Tidak ada saran tambahan.</p>
                  ) : (
                    <ul className="space-y-3 text-sm leading-relaxed text-neutral-700">
                      {result.recommendations.map((r, i) => (
                        <li key={i} className="border-l-2 border-orange-300 pl-4">{r}</li>
                      ))}
                    </ul>
                  )}
                </section>
              </div>
            </div>
          )}

          {/* Tab 2: Pemeriksaan */}
          {tab === "pemeriksaan" && (
            <div className="space-y-12">
              <section className="space-y-4">
                <SectionTitle
                  title={failed === 0 ? "Keutuhan dokumen — terverifikasi" : `Keutuhan dokumen — ${failed} perlu diperiksa`}
                  desc="Perbandingan sebelum → sesudah diformat."
                />
                <div className="border-t border-neutral-200">
                  {result.integrity.map((i) => (
                    <div key={i.label} className="flex items-center justify-between gap-3 border-b border-neutral-100 py-3 text-sm">
                      <span className="flex items-center gap-3 text-neutral-800">
                        {i.ok
                          ? <Check className="size-4 shrink-0 text-green-600" />
                          : <X className="size-4 shrink-0 text-red-600" />}
                        {i.label}
                      </span>
                      <span className="shrink-0 tabular-nums text-neutral-500">
                        {i.before} <span className="mx-1 text-neutral-300">→</span> {i.after}
                      </span>
                    </div>
                  ))}
                </div>
              </section>

              <section className="space-y-4">
                <SectionTitle title="Catatan pemeriksaan" />
                <div className="space-y-3">
                  {result.validation.map((v, i) => {
                    const { border, text, Icon, icon } = LEVEL[v.level];
                    return (
                      <p key={i} className={cx("flex gap-3 border-l-2 pl-4 text-sm leading-relaxed", border, text)}>
                        <Icon className={cx("mt-0.5 size-4 shrink-0", icon)} />
                        <span>{v.message}</span>
                      </p>
                    );
                  })}
                </div>
              </section>
            </div>
          )}

          {/* Tab 3: Urutan dokumen */}
          {tab === "urutan" && (
            <section className="space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <SectionTitle
                  title="Urutan dokumen"
                  desc={`Menampilkan ${preview.length} dari ${result.preview.length} bagian.`}
                />
                <button
                  type="button"
                  onClick={() => setShowAll((v) => !v)}
                  className="rounded-full bg-neutral-100 px-4 py-2 text-xs text-neutral-700 transition hover:bg-neutral-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-orange-400"
                >
                  {showAll ? "Ringkas" : "Tampilkan semua"}
                </button>
              </div>
              <div className="max-h-[32rem] overflow-y-auto border-t border-neutral-200">
                {preview.map((p, i) => (
                  <div
                    key={i}
                    className={cx("flex items-center gap-3 border-b border-neutral-100 py-2.5", ROLE_INDENT[p.role])}
                  >
                    <span
                      className={cx(
                        "w-32 shrink-0 truncate rounded-full px-2.5 py-0.5 text-center text-[11px]",
                        ROLE_STYLE[p.role] ?? "bg-neutral-100 text-neutral-600",
                      )}
                    >
                      {ROLE_LABEL[p.role]}
                    </span>
                    <span className="min-w-0 flex-1 truncate text-sm text-neutral-700">{p.text || "—"}</span>
                  </div>
                ))}
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
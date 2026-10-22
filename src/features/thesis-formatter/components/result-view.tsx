"use client";
import { useMemo, useState } from "react";
import { Check, Download, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
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
  chapter: "bg-slate-900 text-white",
  chapter_label: "bg-slate-900 text-white",
  front_heading: "bg-purple-200 text-purple-900",
  list_heading: "bg-purple-200 text-purple-900",
  reference_heading: "bg-purple-200 text-purple-900",
  appendix_heading: "bg-purple-200 text-purple-900",
  section: "bg-green-200 text-green-900",
  subsection: "bg-green-100 text-green-900",
  figure: "bg-amber-200 text-amber-900",
  figure_caption: "bg-amber-100 text-amber-900",
  table_caption: "bg-amber-100 text-amber-900",
  cover: "bg-slate-200 text-slate-700",
};

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

const LEVEL_STYLE = {
  ok: "bg-green-50 text-green-900",
  warn: "bg-amber-50 text-amber-900",
  error: "bg-red-50 text-red-800",
} as const;

export function ResultView({ result }: { result: FormatResult }) {
  const [showAll, setShowAll] = useState(false);
  const preview = useMemo(
    () => (showAll ? result.preview : result.preview.filter((p) => p.role !== "empty").slice(0, 40)),
    [result.preview, showAll],
  );

  const failed = result.integrity.filter((i) => !i.ok).length;

  return (
    <div className="space-y-6">
      <Card className="rounded-[2rem] border-0 bg-slate-900 shadow-none">
        <CardContent className="flex flex-col gap-4 py-6 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <p className="text-sm text-slate-300">
              {result.mode === "preserve"
                ? "Dokumen asli disunting di tempat — struktur, tabel, dan gambar dipertahankan"
                : "Dokumen disusun ulang dari teks PDF"}
            </p>
            <p className="mt-1 text-lg font-medium tracking-tight text-white">{result.fileName}</p>
          </div>
          <Button onClick={() => download(result)} className="h-12 rounded-full bg-white px-6 text-sm text-slate-900 hover:bg-slate-100">
            <Download className="mr-2 size-4" /> Unduh DOCX
          </Button>
        </CardContent>
      </Card>

      <div className="grid gap-6 lg:grid-cols-2">
        <Card className="rounded-[2rem] border-0 bg-purple-100 shadow-none">
          <CardHeader>
            <CardTitle className="text-lg font-medium tracking-tight text-slate-900">
              Keutuhan dokumen {failed === 0 ? "— terverifikasi" : `— ${failed} perlu diperiksa`}
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-2">
            {result.integrity.map((i) => (
              <div key={i.label} className="flex items-center justify-between gap-3 rounded-2xl bg-white px-4 py-3 text-sm">
                <span className="flex items-center gap-2 text-slate-700">
                  {i.ok
                    ? <Check className="size-4 shrink-0 text-green-600" />
                    : <X className="size-4 shrink-0 text-red-600" />}
                  {i.label}
                </span>
                <span className="shrink-0 tabular-nums text-slate-500">
                  {i.before} <span className="mx-1 text-slate-300">→</span> {i.after}
                </span>
              </div>
            ))}
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-green-100 shadow-none">
          <CardHeader>
            <CardTitle className="text-lg font-medium tracking-tight text-slate-900">Struktur terdeteksi</CardTitle>
          </CardHeader>
          <CardContent>
            <dl className="grid grid-cols-2 gap-2 sm:grid-cols-3">
              {([
                ["Bab", result.stats.chapters],
                ["Subbab", result.stats.sections],
                ["Anak subbab", result.stats.subsections],
                ["Bagian awal", result.stats.frontMatter],
                ["Tabel", result.stats.tables],
                ["Gambar", result.stats.figures],
                ["Pustaka", result.stats.references],
                ["Lampiran", result.stats.appendices],
              ] as const).map(([label, value]) => (
                <div key={label} className="rounded-2xl bg-white px-4 py-3">
                  <dt className="text-xs text-slate-500">{label}</dt>
                  <dd className="mt-0.5 text-xl font-medium tabular-nums text-slate-900">{value}</dd>
                </div>
              ))}
            </dl>
          </CardContent>
        </Card>
      </div>

      <Card className="rounded-[2rem] border-0 bg-slate-50 shadow-none">
        <CardHeader><CardTitle className="text-lg font-medium tracking-tight text-slate-900">Catatan pemeriksaan</CardTitle></CardHeader>
        <CardContent className="space-y-2">
          {result.validation.map((v, i) => (
            <p key={i} className={`rounded-2xl px-4 py-3 text-sm leading-relaxed ${LEVEL_STYLE[v.level]}`}>{v.message}</p>
          ))}
        </CardContent>
      </Card>

      <div className="grid gap-6 lg:grid-cols-2">
        <Card className="rounded-[2rem] border-0 bg-slate-50 shadow-none">
          <CardHeader><CardTitle className="text-lg font-medium tracking-tight text-slate-900">Yang diterapkan</CardTitle></CardHeader>
          <CardContent>
            <ul className="space-y-2 text-sm text-slate-700">
              {result.changes.map((c, i) => (
                <li key={i} className="flex gap-2.5">
                  <Check className="mt-0.5 size-4 shrink-0 text-green-600" />
                  <span className="leading-relaxed">{c}</span>
                </li>
              ))}
            </ul>
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-slate-50 shadow-none">
          <CardHeader><CardTitle className="text-lg font-medium tracking-tight text-slate-900">Saran</CardTitle></CardHeader>
          <CardContent>
            <ul className="space-y-2 text-sm leading-relaxed text-slate-700">
              {result.recommendations.map((r, i) => <li key={i} className="rounded-2xl bg-white px-4 py-3">{r}</li>)}
            </ul>
          </CardContent>
        </Card>
      </div>

      <Card className="rounded-[2rem] border-0 bg-slate-50 shadow-none">
        <CardHeader className="flex-row items-center justify-between space-y-0">
          <CardTitle className="text-lg font-medium tracking-tight text-slate-900">Urutan dokumen</CardTitle>
          <button
            type="button"
            onClick={() => setShowAll((v) => !v)}
            className="rounded-full bg-white px-4 py-2 text-xs text-slate-600 transition hover:bg-slate-100"
          >
            {showAll ? "Ringkas" : "Tampilkan semua"}
          </button>
        </CardHeader>
        <CardContent className="space-y-1.5">
          {preview.map((p, i) => (
            <div key={i} className="flex items-start gap-3 rounded-2xl bg-white px-4 py-2.5">
              <span className={`shrink-0 rounded-full px-2.5 py-0.5 text-[11px] ${ROLE_STYLE[p.role] ?? "bg-slate-100 text-slate-600"}`}>
                {ROLE_LABEL[p.role]}
              </span>
              <span className="min-w-0 flex-1 truncate text-sm text-slate-700">{p.text || "—"}</span>
            </div>
          ))}
        </CardContent>
      </Card>
    </div>
  );
}
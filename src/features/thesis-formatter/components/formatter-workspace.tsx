"use client";
import { useState } from "react";
import type { ReactNode } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { Wand2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Field, selectCls, inputCls } from "@/components/shared/field";
import { ProcessingProgress } from "@/components/shared/processing-progress";
import { FileDropzone } from "@/features/file-upload/components/file-dropzone";
import { formatConfigSchema, defaultFormatConfig, configForPreset, type FormatConfig } from "../schemas/format-config";
import { PRESETS, presetById } from "../presets";
import { useFormatThesis } from "../hooks/use-format-thesis";
import { ResultView } from "./result-view";

const NUMS = [
  { name: "marginTop", label: "Margin atas (cm)", step: 0.1 },
  { name: "marginBottom", label: "Margin bawah (cm)", step: 0.1 },
  { name: "marginLeft", label: "Margin kiri (cm)", step: 0.1 },
  { name: "marginRight", label: "Margin kanan (cm)", step: 0.1 },
  { name: "fontSize", label: "Ukuran font (pt)", step: 1 },
  { name: "lineSpacing", label: "Spasi baris", step: 0.1 },
  { name: "indent", label: "Indentasi paragraf (cm)", step: 0.01 },
  { name: "chapterSpaceAfter", label: "Jarak setelah judul bab (baris)", step: 0.5 },
  { name: "tableFontSize", label: "Font isi tabel (pt)", step: 1 },
  { name: "captionFontSize", label: "Font judul tabel/gambar (pt)", step: 1 },
  { name: "referenceHanging", label: "Gantungan daftar pustaka (cm)", step: 0.01 },
  { name: "coverTitleSize", label: "Font judul sampul (pt)", step: 1 },
  { name: "coverInstitutionSize", label: "Font identitas kampus (pt)", step: 1 },
] as const;

const SELECTS = [
  { name: "lineRule", label: "Jenis spasi", options: [["auto", "Kelipatan (1,5 / 2)"], ["exact", "Pasti (exactly pt)"]] },
  { name: "chapterLabel", label: "Nomor bab", options: [["bab", "BAB I ↵ JUDUL"], ["roman", "I ↵ JUDUL"]] },
  { name: "headingLineSpacing", label: "Spasi judul", options: [["single", "Tunggal"], ["body", "Sama dengan isi"]] },
  { name: "frontNumberingFrom", label: "Nomor halaman awal", options: [["first", "Sejak setelah sampul"], ["preface", "Sejak Kata Pengantar"]] },
  { name: "tableCaptionAlign", label: "Judul tabel", options: [["left", "Rata kiri"], ["center", "Rata tengah"]] },
  { name: "figureCaptionAlign", label: "Judul gambar", options: [["center", "Rata tengah"], ["left", "Rata kiri"]] },
  { name: "tableNumericAlign", label: "Angka di tabel", options: [["center", "Rata tengah"], ["right", "Rata kanan"]] },
  { name: "captionNumbering", label: "Nomor tabel/gambar", options: [["chapter", "Per bab (Tabel 2.1)"], ["global", "Berurutan (Tabel 1)"]] },
  { name: "captionSeparator", label: "Setelah nomor", options: [["space", "Spasi (Tabel 1 Judul)"], ["period", "Titik (Tabel 1. Judul)"]] },
  { name: "figureLabel", label: "Sebutan gambar", options: [["auto", "Ikuti dokumen"], ["Gambar", "Gambar"], ["Ilustrasi", "Ilustrasi"]] },
  { name: "listStyle", label: "Format daftar isi", options: [["word", "Daftar isi otomatis Word"], ["columns", "Kolom (nomor | judul | halaman)"]] },
  { name: "listMarkers", label: "Penanda butir daftar", options: [["uniform", "Diseragamkan"], ["keep", "Ikuti dokumen (1), a., …"]] },
  { name: "listLayout", label: "Daftar bertingkat", options: [["nested", "Menjorok per tingkat"], ["flat", "Rata: semua penanda di margin"]] },
] as const;

const CHECKS = [
  { name: "chapterUppercase", label: "Judul bab huruf kapital" },
  { name: "romanFrontMatter", label: "Romawi untuk bagian awal" },
  { name: "abstractIndent", label: "Abstrak memakai indentasi" },
  { name: "referenceGap", label: "Baris kosong antar pustaka" },
  { name: "normalizeTables", label: "Rapikan tata letak tabel" },
  { name: "cleanupBlankPages", label: "Bersihkan halaman kosong" },
  { name: "coverInstitutionUppercase", label: "Identitas kampus huruf kapital" },
  { name: "coverInstitutionBold", label: "Identitas kampus tebal" },
  { name: "coverTitleUppercase", label: "Judul sampul huruf kapital" },
  { name: "coverNameUppercase", label: "Nama penulis huruf kapital" },
  { name: "coverIdentityBold", label: "Nama & NPM tebal" },
  { name: "coverAuthorLabel", label: "Label \"Disusun oleh\" di sampul" },
  { name: "coverTranslationSameSize", label: "Judul terjemahan sama besar" },
] as const;

function StepTitle({ n, children }: { n: number; children: ReactNode }) {
  return (
    <CardTitle className="flex items-center gap-3 text-lg font-medium tracking-tight text-neutral-900">
      <span className="grid size-7 place-items-center rounded-full bg-neutral-900 text-xs font-medium text-orange-300">{n}</span>
      {children}
    </CardTitle>
  );
}

export function FormatterWorkspace() {
  const [file, setFile] = useState<File | null>(null);
  const mutation = useFormatThesis();
  const { register, handleSubmit, reset, getValues, watch, formState: { errors } } = useForm<FormatConfig>({
    resolver: zodResolver(formatConfigSchema), defaultValues: defaultFormatConfig,
  });

  const preset = presetById(watch("presetId"));
  const applyPreset = (id: string) => reset({ ...configForPreset(id), docType: getValues("docType") });

  const onSubmit = (config: FormatConfig) => file && mutation.mutate({ file, config });

  return (
    <div className="mx-auto max-w-6xl space-y-8 bg-neutral-950 px-4 py-12">
      <div className="relative">
        <span className="pointer-events-none absolute -top-4 right-0 hidden size-28 rounded-full bg-orange-400 sm:block" aria-hidden="true" />
        <span className="pointer-events-none absolute right-24 top-12 hidden size-6 rounded-full bg-orange-300 sm:block" aria-hidden="true" />
        <div className="relative max-w-2xl">
          <p className="text-sm text-neutral-400">Format skripsi</p>
          <h1 className="mt-2 text-3xl font-normal leading-[1.1] tracking-tight text-white sm:text-4xl">AI Thesis Formatter</h1>
          <p className="mt-3 text-sm leading-relaxed text-neutral-400">
            Untuk berkas DOCX, dokumen asli disunting langsung: tabel, gambar, dan urutan bagian dipertahankan apa adanya.
            Yang diubah hanya aturan format — font, margin, spasi, dan tata letak tabel.
          </p>
        </div>
      </div>

      <form onSubmit={handleSubmit(onSubmit)} className="grid gap-6 lg:grid-cols-[1fr_1.2fr]">
        <Card className="rounded-[2rem] border-0 bg-orange-200 shadow-none">
          <CardHeader><StepTitle n={1}>Unggah skripsi</StepTitle></CardHeader>
          <CardContent className="space-y-4">
            <FileDropzone file={file} onChange={setFile} />
            <Field label="Jenis dokumen">
              <select className={selectCls} {...register("docType")}>
                <option value="skripsi">Skripsi</option><option value="tesis">Tesis</option><option value="proposal">Proposal</option>
              </select>
            </Field>
            <p className="rounded-2xl bg-white/70 p-4 text-xs leading-relaxed text-neutral-700">
              Unggah DOCX agar tabel dan gambar ikut terjaga. PDF tidak menyimpan tabel dan gambar
              sebagai struktur, sehingga dokumennya terpaksa disusun ulang dari teks.
            </p>
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-white shadow-none">
          <CardHeader><StepTitle n={2}>Aturan format</StepTitle></CardHeader>
          <CardContent className="space-y-5">
            <Field label="Pedoman kampus">
              <select className={selectCls} {...register("presetId", { onChange: (e) => applyPreset(e.target.value) })}>
                {PRESETS.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </Field>
            {preset.id !== "umum" && (
              <div className="space-y-2 rounded-2xl bg-orange-50 p-4 text-xs leading-relaxed text-neutral-700">
                <p className="font-medium text-neutral-900">{preset.source}</p>
                <ul className="list-disc space-y-1 pl-4">
                  {preset.notes.map((n) => <li key={n}>{n}</li>)}
                </ul>
                <p className="text-neutral-500">Semua nilai di bawah sudah diisi dari pedoman dan tetap bisa diubah.</p>
              </div>
            )}

            <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
              <Field label="Ukuran kertas">
                <select className={selectCls} {...register("paper")}><option>A4</option><option>Letter</option></select>
              </Field>
              <Field label="Font">
                <select className={selectCls} {...register("fontFamily")}><option>Times New Roman</option><option>Arial</option><option>Calibri</option></select>
              </Field>
              <Field label="Nomor halaman">
                <select className={selectCls} {...register("pageNumbers")}>
                  <option value="bottom-center">Bawah tengah</option>
                  <option value="bottom-right">Bawah kanan</option>
                  <option value="keep">Ikuti dokumen asli</option>
                  <option value="none">Tanpa</option>
                </select>
              </Field>
              {NUMS.map((n) => (
                <Field key={n.name} label={n.label} error={errors[n.name]?.message}>
                  <Input type="number" step={n.step} className={inputCls} {...register(n.name, { valueAsNumber: true })} />
                </Field>
              ))}
              {SELECTS.map((s) => (
                <Field key={s.name} label={s.label}>
                  <select className={selectCls} {...register(s.name)}>
                    {s.options.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                  </select>
                </Field>
              ))}
              <Field label="Format sitasi">
                <select className={selectCls} {...register("citationStyle")}><option>APA</option><option>IEEE</option><option>Harvard</option><option>Vancouver</option></select>
              </Field>
              <Field label="Daftar isi lama">
                <select className={selectCls} {...register("tocMode")}>
                  <option value="keep">Pertahankan</option>
                  <option value="rebuild">Ganti otomatis</option>
                </select>
              </Field>
              <Field label="Border tabel">
                <select className={selectCls} {...register("tableBorders")}>
                  <option value="all">Garis penuh</option>
                  <option value="minimal">Garis atas &amp; bawah</option>
                  <option value="keep">Ikuti dokumen asli</option>
                </select>
              </Field>
            </div>

            <div className="grid gap-2 sm:grid-cols-2">
              {CHECKS.map((c) => (
                <label
                  key={c.name}
                  className="flex cursor-pointer items-center gap-2.5 rounded-full bg-orange-100 px-4 py-2.5 text-sm text-neutral-800 transition hover:bg-orange-200"
                >
                  <input type="checkbox" className="size-4 accent-orange-500" {...register(c.name)} /> {c.label}
                </label>
              ))}
            </div>

            <Button
              type="submit"
              disabled={!file || mutation.isPending}
              className="h-12 w-full rounded-full bg-neutral-900 pl-2 pr-6 text-sm text-white hover:bg-neutral-800"
            >
              <span className="mr-2 grid size-8 place-items-center rounded-full bg-orange-400 text-neutral-950">
                <Wand2 className="size-4" />
              </span>
              Format sekarang
            </Button>
          </CardContent>
        </Card>
      </form>

      <ProcessingProgress
        active={mutation.isPending}
        stages={["Membaca struktur dokumen…", "AI menganalisis bagian yang ambigu…", "Menerapkan format…", "Memeriksa keutuhan isi…"]}
      />
      {mutation.isError && <p className="rounded-3xl bg-red-500/10 p-5 text-sm text-red-300">{mutation.error.message}</p>}
      {mutation.data && <ResultView result={mutation.data} />}
    </div>
  );
}
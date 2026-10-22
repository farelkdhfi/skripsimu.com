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
import { formatConfigSchema, defaultFormatConfig, type FormatConfig } from "../schemas/format-config";
import { useFormatThesis } from "../hooks/use-format-thesis";
import { ResultView } from "./result-view";

const NUMS = [
  { name: "marginTop", label: "Margin atas (cm)", step: 0.1 },
  { name: "marginBottom", label: "Margin bawah (cm)", step: 0.1 },
  { name: "marginLeft", label: "Margin kiri (cm)", step: 0.1 },
  { name: "marginRight", label: "Margin kanan (cm)", step: 0.1 },
  { name: "fontSize", label: "Ukuran font (pt)", step: 1 },
  { name: "lineSpacing", label: "Spasi baris", step: 0.1 },
  { name: "indent", label: "Indentasi paragraf (cm)", step: 0.05 },
  { name: "coverTitleSize", label: "Font judul sampul (pt)", step: 1 },
  { name: "coverInstitutionSize", label: "Font identitas kampus (pt)", step: 1 },
] as const;

const CHECKS = [
  { name: "chapterUppercase", label: "Judul bab huruf kapital" },
  { name: "romanFrontMatter", label: "Romawi untuk bagian awal" },
  { name: "normalizeTables", label: "Rapikan tata letak tabel" },
  { name: "cleanupBlankPages", label: "Bersihkan halaman kosong" },
  { name: "coverInstitutionUppercase", label: "Identitas kampus huruf kapital" },
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
  const { register, handleSubmit, formState: { errors } } = useForm<FormatConfig>({
    resolver: zodResolver(formatConfigSchema), defaultValues: defaultFormatConfig,
  });

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
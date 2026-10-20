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
] as const;

const CHECKS = [
  { name: "chapterUppercase", label: "Judul bab huruf kapital" },
  { name: "toc", label: "Daftar isi otomatis" },
  { name: "listFigures", label: "Daftar gambar" },
  { name: "listTables", label: "Daftar tabel" },
] as const;

function StepTitle({ n, children }: { n: number; children: ReactNode }) {
  return (
    <CardTitle className="flex items-center gap-3 text-lg font-medium tracking-tight text-slate-900">
      <span className="grid size-7 place-items-center rounded-full bg-slate-900 text-xs font-medium text-white">{n}</span>
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
    <div className="mx-auto max-w-6xl space-y-8 bg-white px-4 py-12">
      <div className="relative">
        <span className="pointer-events-none absolute -top-4 right-0 hidden size-28 rounded-full bg-purple-100 sm:block" aria-hidden="true" />
        <span className="pointer-events-none absolute right-24 top-12 hidden size-6 rounded-full bg-green-200 sm:block" aria-hidden="true" />
        <div className="relative max-w-2xl">
          <p className="text-sm text-slate-500">Format skripsi</p>
          <h1 className="mt-2 text-3xl font-normal leading-[1.1] tracking-tight text-slate-900 sm:text-4xl">AI Thesis Formatter</h1>
          <p className="mt-3 text-sm leading-relaxed text-slate-600">AI menganalisis struktur, library dokumen menerapkan format secara konsisten. Isi tidak diubah.</p>
        </div>
      </div>

      <form onSubmit={handleSubmit(onSubmit)} className="grid gap-6 lg:grid-cols-[1fr_1.2fr]">
        <Card className="rounded-[2rem] border-0 bg-purple-100 shadow-none">
          <CardHeader><StepTitle n={1}>Unggah skripsi</StepTitle></CardHeader>
          <CardContent className="space-y-4">
            <FileDropzone file={file} onChange={setFile} />
            <Field label="Jenis dokumen">
              <select className={selectCls} {...register("docType")}>
                <option value="skripsi">Skripsi</option><option value="tesis">Tesis</option><option value="proposal">Proposal</option>
              </select>
            </Field>
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-green-100 shadow-none">
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
                  <option value="bottom-center">Bawah tengah</option><option value="bottom-right">Bawah kanan</option><option value="none">Tanpa</option>
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
            </div>
            <div className="grid gap-2 sm:grid-cols-2">
              {CHECKS.map((c) => (
                <label
                  key={c.name}
                  className="flex cursor-pointer items-center gap-2.5 rounded-full bg-white px-4 py-2.5 text-sm text-slate-700 transition hover:bg-slate-50"
                >
                  <input type="checkbox" className="size-4 accent-purple-600" {...register(c.name)} /> {c.label}
                </label>
              ))}
            </div>
            <Button
              type="submit"
              disabled={!file || mutation.isPending}
              className="h-12 w-full rounded-full bg-slate-900 pl-2 pr-6 text-sm text-white hover:bg-slate-800"
            >
              <span className="mr-2 grid size-8 place-items-center rounded-full bg-slate-700">
                <Wand2 className="size-4" />
              </span>
              Format sekarang
            </Button>
          </CardContent>
        </Card>
      </form>

      <ProcessingProgress active={mutation.isPending} stages={["Membaca dokumen…", "AI menganalisis struktur bab…", "Menerapkan format…", "Menyusun file DOCX…"]} />
      {mutation.isError && <p className="rounded-3xl bg-red-50 p-5 text-sm text-red-700">{mutation.error.message}</p>}
      {mutation.data && <ResultView result={mutation.data} />}
    </div>
  );
}
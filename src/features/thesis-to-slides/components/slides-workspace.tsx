"use client";
import { useState } from "react";
import type { ReactNode } from "react";
import { useForm, useWatch } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { Download, Presentation } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Field, selectCls, inputCls } from "@/components/shared/field";
import { ProcessingProgress } from "@/components/shared/processing-progress";
import { FileDropzone } from "@/features/file-upload/components/file-dropzone";
import { downloadBlob } from "@/features/file-upload/utils/download";
import { slidesConfigSchema, type SlidesConfig } from "../schemas/slides";
import { useExportPptx, useGenerateSlides } from "../hooks/use-slides";
import { SlidePreview } from "./slide-preview";

function StepTitle({ n, children }: { n: number; children: ReactNode }) {
  return (
    <CardTitle className="flex items-center gap-3 text-lg font-medium tracking-tight text-slate-900">
      <span className="grid size-7 place-items-center rounded-full bg-slate-900 text-xs font-medium text-white">{n}</span>
      {children}
    </CardTitle>
  );
}

export function SlidesWorkspace() {
  const [file, setFile] = useState<File | null>(null);
  const gen = useGenerateSlides();
  const exp = useExportPptx();
  const { register, handleSubmit, control, getValues, formState: { errors } } = useForm<SlidesConfig>({
    resolver: zodResolver(slidesConfigSchema),
    defaultValues: { preset: "10", customCount: 12, language: "Indonesia", style: "academic-blue", detail: "sedang", speakerNotes: true },
  });
  const preset = useWatch({ control, name: "preset" });
  const style = gen.variables?.config.style ?? "academic-blue";

  return (
    <div className="mx-auto max-w-6xl space-y-8 bg-white px-4 py-12">
      <div className="relative">
        <span className="pointer-events-none absolute -top-4 right-0 hidden size-28 rounded-full bg-green-100 sm:block" aria-hidden="true" />
        <span className="pointer-events-none absolute right-24 top-12 hidden size-6 rounded-full bg-purple-200 sm:block" aria-hidden="true" />
        <div className="relative max-w-2xl">
          <p className="text-sm text-slate-500">Presentasi sidang</p>
          <h1 className="mt-2 text-3xl font-normal leading-[1.1] tracking-tight text-slate-900 sm:text-4xl">Skripsi to Slides</h1>
          <p className="mt-3 text-sm leading-relaxed text-slate-600">Alur presentasi disusun AI dari isi skripsi. Grafik hanya dibuat dari angka yang benar-benar ada di dokumen.</p>
        </div>
      </div>

      <form onSubmit={handleSubmit((config) => file && gen.mutate({ file, config }))} className="grid gap-6 lg:grid-cols-2">
        <Card className="rounded-[2rem] border-0 bg-purple-100 shadow-none">
          <CardHeader><StepTitle n={1}>Unggah skripsi</StepTitle></CardHeader>
          <CardContent><FileDropzone file={file} onChange={setFile} /></CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-green-100 shadow-none">
          <CardHeader><StepTitle n={2}>Pengaturan presentasi</StepTitle></CardHeader>
          <CardContent className="space-y-5">
            <div className="grid grid-cols-2 gap-3">
              <Field label="Jumlah slide">
                <select className={selectCls} {...register("preset")}>
                  {["5", "10", "15", "20"].map((n) => <option key={n} value={n}>{n} slide</option>)}
                  <option value="custom">Jumlah khusus</option>
                </select>
              </Field>
              {preset === "custom" && (
                <Field label="Jumlah (3-40)" error={errors.customCount?.message}>
                  <Input type="number" className={inputCls} {...register("customCount", { valueAsNumber: true })} />
                </Field>
              )}
              <Field label="Bahasa"><select className={selectCls} {...register("language")}><option>Indonesia</option><option>English</option></select></Field>
              <Field label="Gaya visual">
                <select className={selectCls} {...register("style")}><option value="academic-blue">Academic Blue</option><option value="soft-violet">Soft Violet</option><option value="mono-dark">Mono Dark</option></select>
              </Field>
              <Field label="Tingkat detail">
                <select className={selectCls} {...register("detail")}><option value="ringkas">Ringkas</option><option value="sedang">Sedang</option><option value="detail">Detail</option></select>
              </Field>
            </div>
            <label className="flex cursor-pointer items-center gap-2.5 rounded-full bg-white px-4 py-2.5 text-sm text-slate-700 transition hover:bg-slate-50">
              <input type="checkbox" className="size-4 accent-purple-600" {...register("speakerNotes")} /> Sertakan catatan presenter
            </label>
            <Button
              type="submit"
              disabled={!file || gen.isPending}
              className="h-12 w-full rounded-full bg-slate-900 pl-2 pr-6 text-sm text-white hover:bg-slate-800"
            >
              <span className="mr-2 grid size-8 place-items-center rounded-full bg-slate-700">
                <Presentation className="size-4" />
              </span>
              Buat presentasi
            </Button>
          </CardContent>
        </Card>
      </form>

      <ProcessingProgress active={gen.isPending} stages={["Membaca dokumen…", "AI meringkas bagian skripsi…", "Menyusun alur presentasi…", "Memvalidasi slide…"]} />
      {gen.isError && <p className="rounded-3xl bg-red-50 p-5 text-sm text-red-700">{gen.error.message}</p>}

      {gen.data && (
        <div className="space-y-5">
          {gen.data.warnings.map((w, i) => <p key={i} className="rounded-3xl bg-amber-50 p-4 text-sm text-amber-800">{w}</p>)}
          <SlidePreview deck={gen.data.deck} style={style} />
          <Button
            className="h-12 rounded-full bg-slate-900 pl-2 pr-6 text-sm text-white hover:bg-slate-800"
            disabled={exp.isPending}
            onClick={() => exp.mutate({ deck: gen.data!.deck, style: getValues("style") }, { onSuccess: (b) => downloadBlob(b, "presentasi-skripsi.pptx") })}
          >
            <span className="mr-2 grid size-8 place-items-center rounded-full bg-slate-700">
              <Download className="size-4" />
            </span>
            {exp.isPending ? "Menyiapkan PPTX…" : "Unduh PPTX"}
          </Button>
          {exp.isError && <p className="text-sm text-red-600">{exp.error.message}</p>}
        </div>
      )}
    </div>
  );
}
"use client";
import { AlertTriangle, CheckCircle2, Download, Lightbulb } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { base64ToBlob, downloadBlob } from "@/features/file-upload/utils/download";
import type { FormatResult, Role } from "../types";

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const style: Record<Role, string> = {
  cover: "text-center font-semibold", chapter: "text-center font-bold uppercase mt-6", section: "font-semibold mt-3",
  subsection: "font-semibold italic mt-2", figure_caption: "text-center text-xs text-slate-500", table_caption: "text-center text-xs text-slate-500",
  body: "text-justify indent-6", reference: "pl-6 -indent-6", skip: "",
};

const titleCls = "text-lg font-medium tracking-tight";

export function ResultView({ result }: { result: FormatResult }) {
  const s = result.stats;
  const chips = [`${s.chapters} bab`, `${s.sections} subbab`, `${s.figures} gambar`, `${s.tables} tabel`];
  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_1.4fr]">
      <div className="space-y-4">
        <Button
          className="h-12 w-full rounded-full bg-slate-900 pl-2 pr-6 text-sm text-white hover:bg-slate-800"
          onClick={() => downloadBlob(base64ToBlob(result.docxBase64, DOCX), result.fileName)}
        >
          <span className="mr-2 grid size-8 place-items-center rounded-full bg-slate-700">
            <Download className="size-4" />
          </span>
          Unduh DOCX
        </Button>

        <Card className="rounded-[2rem] border-0 bg-purple-100 shadow-none">
          <CardHeader><CardTitle className={`${titleCls} text-slate-900`}>Ringkasan perubahan</CardTitle></CardHeader>
          <CardContent className="space-y-3 text-sm text-slate-700">
            <div className="flex flex-wrap gap-2">
              {chips.map((c) => (
                <span key={c} className="rounded-full bg-white/70 px-2.5 py-1 text-xs font-medium text-slate-700">{c}</span>
              ))}
            </div>
            <ul className="space-y-1.5">
              {result.changes.map((c, i) => (
                <li key={i} className="flex gap-2.5">
                  <span className="mt-2 size-1.5 shrink-0 rounded-full bg-slate-900" />
                  {c}
                </li>
              ))}
            </ul>
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-green-100 shadow-none">
          <CardHeader><CardTitle className={`${titleCls} text-slate-900`}>Validasi</CardTitle></CardHeader>
          <CardContent className="space-y-2 text-sm text-slate-700">
            {result.validation.map((v, i) => (
              <p key={i} className="flex gap-2">
                {v.level === "ok" ? <CheckCircle2 className="mt-0.5 size-4 shrink-0 text-green-600" /> : <AlertTriangle className="mt-0.5 size-4 shrink-0 text-amber-600" />}
                {v.message}
              </p>
            ))}
          </CardContent>
        </Card>

        <Card className="rounded-[2rem] border-0 bg-slate-900 text-white shadow-none">
          <CardHeader>
            <CardTitle className={`${titleCls} flex items-center gap-3 text-white`}>
              <span className="grid size-8 place-items-center rounded-full bg-green-200 text-slate-900">
                <Lightbulb className="size-4" />
              </span>
              Rekomendasi AI
            </CardTitle>
          </CardHeader>
          <CardContent className="text-sm text-slate-300">
            <ul className="space-y-1.5">
              {result.recommendations.map((r, i) => (
                <li key={i} className="flex gap-2.5">
                  <span className="mt-2 size-1.5 shrink-0 rounded-full bg-purple-300" />
                  {r}
                </li>
              ))}
            </ul>
          </CardContent>
        </Card>
      </div>

      <Card className="rounded-[2rem] border border-slate-200 bg-white shadow-none">
        <CardHeader><CardTitle className={`${titleCls} text-slate-900`}>Preview struktur (150 blok pertama)</CardTitle></CardHeader>
        <CardContent>
          <div className="max-h-[680px] space-y-1.5 overflow-auto rounded-[1.5rem] bg-slate-50 p-6 font-serif text-sm leading-relaxed text-slate-800">
            {result.preview.map((b, i) => <p key={i} className={style[b.role]}>{b.text}</p>)}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
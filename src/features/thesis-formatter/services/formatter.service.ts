import { extractParagraphs } from "@/features/file-upload/services/extract-text";
import { formatDocxBytes, describeChanges, outputName } from "./format-core";
import { useNodeIO } from "../front/node";
import { buildDocxFromText } from "./build-docx";
import { classifyNodes } from "./analyze.service";
import { buildOutline, outlineStats } from "../structure/outline";
import type { Signals } from "../structure/roles";
import type { FormatConfig } from "../schemas/format-config";
import type { DocStats, FormatResult, IntegrityCheck } from "../types";
import { extractCoverAI } from "./cover-ai.service";
import { extractFrontAI } from "./front-ai.service";
import { presetById } from "../presets";
import type { FrontFieldInput } from "../front/types";

const EMPTY_STATS: DocStats = {
  chapters: 0, sections: 0, subsections: 0, figures: 0, figureCaptions: 0,
  tableCaptions: 0, tables: 0, references: 0, frontMatter: 0, appendices: 0,
  referenceHeadings: 0,
};

async function formatFromDocx(file: File, c: FormatConfig, frontFields?: FrontFieldInput): Promise<FormatResult> {
  useNodeIO();
  return formatDocxBytes(new Uint8Array(await file.arrayBuffer()), file.name, c, frontFields, {
    classify: classifyNodes, cover: extractCoverAI, front: extractFrontAI,
  });
}

/* ------------------------------------------------------------------ */
/* Jalur PDF — terpaksa menyusun ulang                                  */
/* ------------------------------------------------------------------ */

async function formatFromPdf(file: File, paragraphs: string[], c: FormatConfig): Promise<FormatResult> {
  const signals: Signals[] = paragraphs.map((t) => ({ text: t, hasDrawing: false, isTable: false }));
  const ai = await classifyNodes(signals);
  const outline = buildOutline(
    paragraphs.map((text) => ({
      kind: "paragraph" as const, id: 0, xml: "", text,
      hasDrawing: false, hasPageBreak: false, hadPageBreakBefore: false,
      hasSectPr: false, hasNumbering: false, hasField: false, styleId: null,
      align: null, tabCount: (text.match(/\t/g) || []).length, allBold: false,
    })),
    ai.roles,
  );

  const buf = await buildDocxFromText(paragraphs, outline, c);
  const stats = { ...EMPTY_STATS, ...outlineStats([], outline) };

  const validation: FormatResult["validation"] = [
    { level: "warn", message: "Sumber PDF: tabel, gambar, dan batas section tidak tersimpan sebagai struktur di dalam PDF, sehingga dokumen disusun ulang dari teks. Unggah DOCX agar tabel dan gambar dipertahankan utuh." },
  ];
  for (const n of ai.notices) validation.push({ level: "warn", message: n });
  if (stats.chapters === 0)
    validation.push({ level: "warn", message: "Judul bab tidak terdeteksi. Pastikan penulisannya seperti 'BAB I PENDAHULUAN'." });

  return {
    fileName: outputName(file.name, c),
    docxBase64: buf.toString("base64"),
    source: "pdf",
    mode: "rebuild",
    preview: paragraphs.slice(0, 200).map((text, i) => ({
      role: outline.items[i]?.role ?? "body",
      zone: outline.items[i]?.zone ?? "main",
      text: text.slice(0, 280),
    })),
    stats,
    integrity: [
      { label: "Paragraf", before: paragraphs.length, after: paragraphs.length, ok: true },
      { label: "Tabel & gambar", before: "tidak tersedia di PDF", after: "tidak dipindahkan", ok: false },
    ],
    changes: describeChanges(c, stats, ["Dokumen disusun ulang dari teks PDF"]),
    validation,
    recommendations: [
      ...ai.recommendations,
      `Format sitasi ${c.citationStyle}: periksa kesesuaian sitasi dan daftar pustaka secara manual.`,
      "Sisipkan kembali tabel dan gambar secara manual, atau unggah ulang versi DOCX untuk hasil yang utuh.",
    ],
  };
}

/* ------------------------------------------------------------------ */

/** `frontFields`: isian data bagian depan dari form konfirmasi (opsional). */
export async function formatThesis(file: File, c: FormatConfig, frontFields?: FrontFieldInput): Promise<FormatResult> {
  const isDocx = /\.docx$/i.test(file.name) ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  // DOCX tidak pernah melewati ekstraksi teks: begitu dokumen menjadi array
  // string, tabel dan gambar sudah hilang sebelum format diterapkan.
  if (isDocx) return formatFromDocx(file, c, frontFields);

  const { paragraphs } = await extractParagraphs(file);
  if (paragraphs.length < 3)
    throw Object.assign(
      new Error("Teks dokumen tidak terbaca. Jika PDF hasil scan, gunakan DOCX atau PDF berbasis teks."),
      { status: 422 },
    );
  return formatFromPdf(file, paragraphs, c);
}
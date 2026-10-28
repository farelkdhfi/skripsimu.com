/**
 * Inti pemformatan DOCX — TANPA kode khusus Node, jadi bisa jalan di browser (Web Worker) maupun
 * di server/tes. Panggilan AI disuntikkan lewat `AiPorts`: di browser lewat fetch ke
 * /api/formatter/ai (hanya teks kecil yang dikirim), di server langsung ke layanan AI.
 */
import JSZip from "jszip";
import { transformDocx } from "../docx/transform";
import { parseDocument, flatten } from "../docx/document-model";
import { presetById } from "../presets";
import type { Signals } from "../structure/roles";
import type { FormatConfig } from "../schemas/format-config";
import type { DocStats, FormatResult, IntegrityCheck } from "../types";
import type { FrontFieldInput } from "../front/types";
import type { FrontPointers } from "../front/ai";
import type { CoverFields } from "../docx/cover-template";
import type { ClassifyResult } from "./analyze.service";

export type AiPorts = {
  classify: (signals: Signals[]) => Promise<ClassifyResult>;
  cover: (lines: string[]) => Promise<CoverFields | null>;
  front: (listing: string) => Promise<FrontPointers | null>;
};

/** Base64 tanpa Buffer (browser) — dipotong per 32 KB supaya tidak melampaui batas argumen. */
export function toBase64(bytes: Uint8Array): string {
  const B = (globalThis as { Buffer?: { from(b: Uint8Array): { toString(enc: string): string } } }).Buffer;
  if (B) return B.from(bytes).toString("base64");
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

/** "skripsi.docx" + preset UNPAD-FAPET → "skripsi-FORMATTED-UNPAD-FAPET.docx" */
export function outputName(name: string, c: FormatConfig): string {
  const base = name.replace(/\.(docx|pdf)$/i, "");
  const code = presetById(c.presetId)?.code ?? "";
  return `${base}-FORMATTED${code ? `-${code}` : ""}.docx`;
}

export function describeChanges(c: FormatConfig, s: DocStats, extra: string[]): string[] {
  const preset = presetById(c.presetId);
  return [
    preset.id !== "umum" ? `Preset pedoman: ${preset.name} — ${preset.source}` : "",
    `Kertas ${c.paper}, margin ${c.marginTop}/${c.marginBottom}/${c.marginLeft}/${c.marginRight} cm`,
    `Font ${c.fontFamily} ${c.fontSize} pt, spasi ${c.lineRule === "exact" ? `exactly ${c.lineSpacing * c.fontSize} pt` : c.lineSpacing}, indentasi paragraf ${c.indent} cm`,
    s.chapters ? `${s.chapters} bab, ${s.sections} subbab, ${s.subsections} anak subbab diformat sebagai heading` : "",
    s.frontMatter ? `${s.frontMatter} halaman bagian awal (motto, persembahan, pengesahan, kata pengantar, daftar) dikenali dan dipertahankan urutannya` : "",
    s.references ? `${s.references} entri daftar pustaka diberi hanging indent` : "",
    s.appendices ? `${s.appendices} bagian lampiran dipertahankan` : "",
    c.pageNumbers === "keep" ? "Penomoran halaman dokumen asli dipertahankan"
      : c.pageNumbers === "none" ? "Nomor halaman dihapus"
        : `Nomor halaman: ${c.pageNumbers}${c.romanFrontMatter ? ", romawi untuk bagian awal dan angka mulai BAB I" : ""}`,
    ...extra,
  ].filter(Boolean);
}

/* ------------------------------------------------------------------ */
/* Jalur DOCX — menyunting dokumen asli di tempat                       */
/* ------------------------------------------------------------------ */

export async function formatDocxBytes(
  input: Uint8Array,
  fileName: string,
  c: FormatConfig,
  frontFields: FrontFieldInput | undefined,
  ai: AiPorts,
): Promise<FormatResult> {

  // AI dijalankan lebih dulu atas teks node, tetapi hasilnya hanya mengisi
  // node yang tidak bisa diputuskan oleh aturan.
  const { parseDocument, flatten } = await import("../docx/document-model");
  const { default: JSZip } = await import("jszip");
  const probe = await JSZip.loadAsync(input);
  const docXml = await probe.file("word/document.xml")?.async("string");
  if (!docXml) throw Object.assign(new Error("File DOCX tidak dapat dibaca."), { status: 422 });

  const nodes = flatten(parseDocument(docXml).body);
  const signals: Signals[] = nodes.map((n) =>
    n.kind === "table"
      ? { text: "", hasDrawing: false, isTable: true }
      : { text: n.text, hasDrawing: n.hasDrawing, isTable: false },
  );

  const cls = await ai.classify(signals);
  const result = await transformDocx(input, c, { aiRoles: cls.roles, coverAI: ai.cover, frontFields, frontAI: ai.front });

  const failed = (prefix: string) => result.issues.some((i) => i.message.startsWith(prefix));
  // Jumlah yang diharapkan = asli + perubahan yang SENGAJA dibuat (bagian depan dari template,
  // blok tanda tangan, bingkai pernyataan). Selisih di luar itu = ada yang hilang.
  const want = result.intended;
  const integrity: IntegrityCheck[] = [
    { label: "Gambar & objek", before: result.before.drawings, after: result.after.drawings, ok: result.after.drawings >= result.before.drawings + want.drawings },
    { label: "Tabel", before: result.before.tables, after: result.after.tables, ok: result.after.tables === result.before.tables + want.tables },
    { label: "Baris tabel", before: result.before.rows, after: result.after.rows, ok: result.after.rows === result.before.rows + want.rows },
    { label: "Sel tabel", before: result.before.cells, after: result.after.cells, ok: result.after.cells === result.before.cells + want.cells },
    { label: "Teks dokumen", before: `${result.before.text.length} karakter`, after: `${result.after.text.length} karakter`, ok: !failed("Teks dokumen") },
  ];

  const validation: FormatResult["validation"] = [];
  for (const n of cls.notices) validation.push({ level: "warn", message: n });
  for (const i of result.issues) validation.push({ level: i.level, message: i.message });
  for (const n of result.notes) validation.push({ level: "warn", message: n });

  if (!result.issues.length) {
    validation.push({
      level: "ok",
      message: result.front.built
        ? "Isi naskah di luar bagian depan terverifikasi utuh (teks, tabel, dan gambar). Bagian depan disusun ulang dari template; datanya diperiksa ada di teks asli."
        : "Isi, tabel, dan gambar dokumen asli terverifikasi utuh: teks identik, jumlah tabel, baris, sel, dan gambar tidak berubah.",
    });
  }
  if (result.stats.chapters === 0)
    validation.push({ level: "warn", message: "Judul bab tidak terdeteksi. Pastikan penulisannya seperti 'BAB I PENDAHULUAN'." });
  if (result.stats.references === 0)
    validation.push({ level: "warn", message: "Daftar pustaka tidak terdeteksi." });
  if (c.tocMode === "keep" && result.stats.frontMatter > 0)
    validation.push({ level: "ok", message: "Daftar isi/tabel/gambar bawaan dokumen dipertahankan apa adanya, tidak diganti versi otomatis." });

  const extra = [
    result.after.tables ? `${result.after.tables} tabel dirapikan tata letaknya (lebar mengikuti margin, border, baris header diulang per halaman) tanpa mengubah isi sel` : "",
    result.stats.figures ? `${result.stats.figures} gambar dipertahankan di posisi aslinya` : "",
    result.report.pageBreaksAdded ? `${result.report.pageBreaksAdded} pindah halaman ditambahkan pada judul yang belum memulai halaman baru` : "",
    result.report.duplicatePageBreaksRemoved ? `${result.report.duplicatePageBreaksRemoved} pindah halaman ganda dihapus (penyebab halaman kosong)` : "",
    result.report.blankParagraphsRemoved ? `${result.report.blankParagraphsRemoved} baris kosong beruntun dirapikan di luar halaman sampul` : "",
    result.report.orphanHeadingsFixed ? `${result.report.orphanHeadingsFixed} judul yang tertinggal sendirian di bawah halaman disatukan kembali dengan isinya` : "",
    result.images.scaled ? `${result.images.scaled} gambar yang melebihi margin diperkecil secara proporsional` : "",
    result.landscapeSections ? `${result.landscapeSections} section landscape (halaman tabel lebar) dipertahankan orientasinya` : "",
    `${result.sections} section dokumen dipertahankan beserta header/footer-nya`,
    result.toc.tocs ? `${result.toc.tocs} daftar (isi/tabel/gambar/lampiran) dijadikan daftar otomatis yang bisa diklik` : "",
    result.lists ? `${result.lists} butir daftar ketikan diubah menjadi numbering/bullet otomatis Word` : "",
    result.portraitConverted ? `${result.portraitConverted} section tabel landscape dikembalikan ke portrait karena tabelnya muat` : "",
    result.signatures.blocks ? `${result.signatures.blocks} blok tanda tangan dirapikan menjadi tabel tanpa border dengan ruang tanda tangan` : "",
    result.captions.table.numbered ? `${result.captions.table.numbered} keterangan tabel dinomori ulang ${c.captionNumbering === "global" ? "berurutan" : "per bab"} (caption Word: style Caption + nomor otomatis)` : "",
    result.captions.figure.numbered ? `${result.captions.figure.numbered} keterangan ${(result.captions.figure.label ?? "Gambar").toLowerCase()} dinomori ulang ${c.captionNumbering === "global" ? "berurutan" : "per bab"}, diletakkan di bawah gambar` : "",
    result.captions.table.refsUpdated + result.captions.figure.refsUpdated
      ? `${result.captions.table.refsUpdated + result.captions.figure.refsUpdated} rujukan tabel/gambar di dalam teks disesuaikan dengan nomor baru` : "",
    result.captions.notes ? `${result.captions.notes} catatan di bawah tabel diformat 10 pt` : "",
    result.headings.renumbered ? `${result.headings.renumbered} nomor bab/subbab yang tidak konsisten diperbaiki (mis. "C." → "1.3", "BAB 2" → "BAB II")` : "",
    result.headings.relevelled ? `${result.headings.relevelled} judul dipindah ke tingkat yang benar (subbab/anak subbab)` : "",
    result.headings.retitled ? `${result.headings.retitled} judul subbab diseragamkan huruf awal kapital` : "",
  ];

  return {
    fileName: outputName(fileName, c),
    docxBase64: toBase64(result.buffer),
    source: "docx",
    mode: "preserve",
    preview: result.preview as FormatResult["preview"],
    stats: result.stats,
    integrity,
    changes: describeChanges(c, result.stats, extra),
    validation,
    recommendations: [
      ...cls.recommendations,
      c.citationStyle === "APA" || c.citationStyle === "Harvard"
        ? `Daftar pustaka diurutkan sesuai abjad (${c.citationStyle}). Isi tiap entri tidak diubah; periksa entri yang dilaporkan tidak sesuai gaya.`
        : `Gaya ${c.citationStyle}: urutan daftar pustaka mengikuti nomor sitasi di teks sehingga tidak diurutkan sesuai abjad.`,
      "Saat membuka file di Word, pilih \"Yes\" untuk memperbarui field agar nomor halaman daftar isi/tabel/gambar benar dan bisa diklik. Bisa juga Ctrl+A lalu F9.", "Rujukan tabel di teks diperbarui sebagai teks biasa. Bila nanti menambah tabel baru di Word, gunakan References → Insert Caption agar penomorannya tetap berurutan.",
    ],
  };
}
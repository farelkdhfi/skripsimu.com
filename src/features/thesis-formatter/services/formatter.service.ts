import { extractParagraphs } from "@/features/file-upload/services/extract-text";
import { transformDocx } from "../docx/transform";
import { buildDocxFromText } from "./build-docx";
import { classifyNodes } from "./analyze.service";
import { buildOutline, outlineStats } from "../structure/outline";
import type { Signals } from "../structure/roles";
import type { FormatConfig } from "../schemas/format-config";
import type { DocStats, FormatResult, IntegrityCheck } from "../types";
import { extractCoverAI } from "./cover-ai.service";

const EMPTY_STATS: DocStats = {
  chapters: 0, sections: 0, subsections: 0, figures: 0, figureCaptions: 0,
  tableCaptions: 0, tables: 0, references: 0, frontMatter: 0, appendices: 0,
  referenceHeadings: 0,
};

function describeChanges(c: FormatConfig, s: DocStats, extra: string[]): string[] {
  return [
    `Kertas ${c.paper}, margin ${c.marginTop}/${c.marginBottom}/${c.marginLeft}/${c.marginRight} cm`,
    `Font ${c.fontFamily} ${c.fontSize} pt, spasi ${c.lineSpacing}, indentasi paragraf ${c.indent} cm`,
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

async function formatFromDocx(file: File, c: FormatConfig): Promise<FormatResult> {
  const input = Buffer.from(await file.arrayBuffer());

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

  const ai = await classifyNodes(signals);
  const result = await transformDocx(input, c, { aiRoles: ai.roles, coverAI: extractCoverAI });

  const failed = (prefix: string) => result.issues.some((i) => i.message.startsWith(prefix));
  const sig = result.signatures.added;
  const integrity: IntegrityCheck[] = [
    { label: "Gambar & objek", before: result.before.drawings, after: result.after.drawings, ok: !failed("") || result.after.drawings >= result.before.drawings },
    { label: "Tabel", before: result.before.tables, after: result.after.tables, ok: result.after.tables === result.before.tables - result.coverRemoved.tables },
    { label: "Baris tabel", before: result.before.rows, after: result.after.rows, ok: result.after.rows === result.before.rows - result.coverRemoved.rows },
    { label: "Sel tabel", before: result.before.cells, after: result.after.cells, ok: result.after.cells === result.before.cells - result.coverRemoved.cells },
    { label: "Teks dokumen", before: `${result.before.text.length} karakter`, after: `${result.after.text.length} karakter`, ok: !failed("Teks dokumen") },
  ];

  const validation: FormatResult["validation"] = [];
  for (const n of ai.notices) validation.push({ level: "warn", message: n });
  for (const i of result.issues) validation.push({ level: i.level, message: i.message });
  for (const n of result.notes) validation.push({ level: "warn", message: n });

  if (!result.issues.length) {
    validation.push({
      level: "ok",
      message: "Isi, tabel, dan gambar dokumen asli terverifikasi utuh: teks identik, jumlah tabel, baris, sel, dan gambar tidak berubah.",
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
    result.captions.table.numbered ? `${result.captions.table.numbered} keterangan tabel dinomori ulang per bab (caption Word: style Caption + nomor otomatis)` : "",
    result.captions.figure.numbered ? `${result.captions.figure.numbered} keterangan gambar dinomori ulang per bab, diletakkan di bawah gambar` : "",
    result.captions.table.refsUpdated + result.captions.figure.refsUpdated
      ? `${result.captions.table.refsUpdated + result.captions.figure.refsUpdated} rujukan tabel/gambar di dalam teks disesuaikan dengan nomor baru` : "",
    result.captions.notes ? `${result.captions.notes} catatan di bawah tabel diformat 10 pt` : "",
    result.headings.renumbered ? `${result.headings.renumbered} nomor bab/subbab yang tidak konsisten diperbaiki (mis. "C." → "1.3", "BAB 2" → "BAB II")` : "",
    result.headings.relevelled ? `${result.headings.relevelled} judul dipindah ke tingkat yang benar (subbab/anak subbab)` : "",
    result.headings.retitled ? `${result.headings.retitled} judul subbab diseragamkan huruf awal kapital` : "",
  ];

  return {
    fileName: file.name.replace(/\.(docx|pdf)$/i, "") + "-formatted.docx",
    docxBase64: result.buffer.toString("base64"),
    source: "docx",
    mode: "preserve",
    preview: result.preview as FormatResult["preview"],
    stats: result.stats,
    integrity,
    changes: describeChanges(c, result.stats, extra),
    validation,
    recommendations: [
      ...ai.recommendations,
      c.citationStyle === "APA" || c.citationStyle === "Harvard"
        ? `Daftar pustaka diurutkan sesuai abjad (${c.citationStyle}). Isi tiap entri tidak diubah; periksa entri yang dilaporkan tidak sesuai gaya.`
        : `Gaya ${c.citationStyle}: urutan daftar pustaka mengikuti nomor sitasi di teks sehingga tidak diurutkan sesuai abjad.`,
      "Saat membuka file di Word, pilih \"Yes\" untuk memperbarui field agar nomor halaman daftar isi/tabel/gambar benar dan bisa diklik. Bisa juga Ctrl+A lalu F9.", "Rujukan tabel di teks diperbarui sebagai teks biasa. Bila nanti menambah tabel baru di Word, gunakan References → Insert Caption agar penomorannya tetap berurutan.",
    ],
  };
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
    fileName: file.name.replace(/\.(docx|pdf)$/i, "") + "-formatted.docx",
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

export async function formatThesis(file: File, c: FormatConfig): Promise<FormatResult> {
  const isDocx = /\.docx$/i.test(file.name) ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  // DOCX tidak pernah melewati ekstraksi teks: begitu dokumen menjadi array
  // string, tabel dan gambar sudah hilang sebelum format diterapkan.
  if (isDocx) return formatFromDocx(file, c);

  const { paragraphs } = await extractParagraphs(file);
  if (paragraphs.length < 3)
    throw Object.assign(
      new Error("Teks dokumen tidak terbaca. Jika PDF hasil scan, gunakan DOCX atau PDF berbasis teks."),
      { status: 422 },
    );
  return formatFromPdf(file, paragraphs, c);
}
import { extractParagraphs } from "@/features/file-upload/services/extract-text";
import { classifyParagraphs } from "./analyze.service";
import { assembleBlocks } from "./assemble-blocks";
import { buildDocx } from "./build-docx";
import type { FormatConfig } from "../schemas/format-config";
import type { FormatResult } from "../types";

export async function formatThesis(file: File, c: FormatConfig): Promise<FormatResult> {
  const { paragraphs, kind } = await extractParagraphs(file);
  if (paragraphs.length < 3)
    throw Object.assign(new Error("Teks dokumen tidak terbaca. Jika PDF hasil scan, gunakan DOCX atau PDF berbasis teks."), { status: 422 });

  const { roles, recommendations } = await classifyParagraphs(paragraphs);
  const { blocks, stats } = assembleBlocks(paragraphs, roles, c);
  const buf = await buildDocx(blocks, c);

  const validation: FormatResult["validation"] = [];
  if (kind === "pdf") validation.push({ level: "warn", message: "Sumber PDF: struktur dibaca dari teks sehingga akurasi lebih rendah. Unggah DOCX untuk hasil terbaik." });
  else validation.push({ level: "warn", message: "Tabel dan gambar pada file asli belum ikut dipindahkan di versi MVP ini. Periksa dan sisipkan kembali jika ada." });
  if (stats.chapters === 0) validation.push({ level: "warn", message: "Judul bab tidak terdeteksi. Pastikan penulisan seperti 'BAB I PENDAHULUAN'." });
  if (stats.references === 0) validation.push({ level: "warn", message: "Daftar pustaka tidak terdeteksi." });
  const delta = Math.abs(stats.wordsAfter - stats.wordsBefore + 0) / Math.max(1, stats.wordsBefore);
  validation.push(delta < 0.03 || stats.removed > 0
    ? { level: "ok", message: `Isi terjaga: ${stats.wordsBefore} kata asli, ${stats.wordsAfter} kata hasil (selisih dari daftar isi lama yang diganti).` }
    : { level: "warn", message: "Jumlah kata berbeda dari aslinya, periksa hasil." });

  const changes = [
    `Kertas ${c.paper}, margin ${c.marginTop}/${c.marginBottom}/${c.marginLeft}/${c.marginRight} cm`,
    `Font ${c.fontFamily} ${c.fontSize} pt, spasi ${c.lineSpacing}, indentasi paragraf ${c.indent} cm`,
    `${stats.chapters} bab, ${stats.sections} subbab, ${stats.subsections} anak subbab diformat sebagai heading`,
    stats.references ? `${stats.references} entri daftar pustaka diberi hanging indent` : "",
    c.toc ? "Daftar isi otomatis ditambahkan. Di Word, klik 'Yes' saat diminta memperbarui field." : "",
    stats.removed ? `${stats.removed} baris daftar isi/gambar/tabel lama diganti versi baru` : "",
    c.pageNumbers !== "none" ? `Nomor halaman: ${c.pageNumbers}` : "",
  ].filter(Boolean);

  return {
    fileName: file.name.replace(/\.(docx|pdf)$/i, "") + "-formatted.docx",
    docxBase64: buf.toString("base64"),
    source: kind,
    preview: blocks.slice(0, 150).map((b) => ({ role: b.role, text: b.text.slice(0, 280) })),
    stats, changes, validation,
    recommendations: [...recommendations, `Format sitasi ${c.citationStyle}: periksa kesesuaian sitasi dan daftar pustaka secara manual, isi tidak diubah otomatis.`],
  };
}
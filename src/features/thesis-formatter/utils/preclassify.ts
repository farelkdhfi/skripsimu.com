import type { Role } from "../types";

const FRONT =
  /^(abstrak|abstract|kata pengantar|halaman pengesahan|lembar pengesahan|halaman persetujuan|pernyataan keaslian|ucapan terima kasih|persembahan|motto|daftar (isi|gambar|tabel|lampiran|singkatan|istilah|simbol)|daftar pustaka|daftar referensi|referensi|references|bibliography|lampiran(\s+\d+)?|riwayat hidup)$/i;

const headingish = (p: string) => !/[.!?;,:]$/.test(p) && p.split(/\s+/).length <= 14;

export function preclassify(p: string): Role | null {
  // sisa daftar isi: dot leader atau tab + nomor halaman
  if (/\.{4,}\s*[\divxlc]+\s*$/i.test(p) || /\t\s*[\divxlc]+\s*$/i.test(p)) return "skip";

  if (p.length <= 200) {
    if (/^(gambar|figure)\s+\d/i.test(p)) return "figure_caption";
    if (/^(tabel|table)\s+\d/i.test(p)) return "table_caption";
  }
  if (p.length <= 100) {
    if (/^bab\s+([ivxlc]+|\d+)\b/i.test(p)) return "chapter";
    if (FRONT.test(p.replace(/[:\s]+$/, ""))) return "chapter";
    if (headingish(p)) {
      if (/^\d+\.\d+\.\d+\.?\s+\S/.test(p)) return "subsection";
      if (/^\d+\.\d+\.?\s+\S/.test(p)) return "section";
      if (/^[A-Z]\.\s+\S/.test(p)) return "section";
    }
  }
  return null;
}
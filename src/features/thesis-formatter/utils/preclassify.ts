import type { Role } from "../types";

export function preclassify(p: string): Role | null {
  if (/\.{4,}\s*[\divxlc]+\s*$/i.test(p)) return "skip";
  if (p.length <= 200) {
    if (/^(gambar|figure)\s+\d/i.test(p)) return "figure_caption";
    if (/^(tabel|table)\s+\d/i.test(p)) return "table_caption";
  }
  if (p.length <= 100) {
    if (/^bab\s+([ivxlc]+|\d+)\b/i.test(p)) return "chapter";
    if (/^\d+\.\d+\.\d+\.?\s+\S/.test(p)) return "subsection";
    if (/^\d+\.\d+\.?\s+\S/.test(p)) return "section";
  }
  return null;
}
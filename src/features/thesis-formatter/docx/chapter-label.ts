/**
 * Judul bab tanpa kata "BAB": "BAB I ↵ PENDAHULUAN" → "I ↵ PENDAHULUAN"
 * (mis. pedoman UNPAD Fakultas Peternakan). Dijalankan setelah judul bab
 * diseragamkan dan sebelum daftar isi dibuat, sehingga daftar isi ikut
 * menampilkan "I PENDAHULUAN".
 */
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { leadingRaw, removeLeading } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";

/** Kata BAB/CHAPTER di awal judul, diikuti nomor Romawi/angka. */
const WORD_RE = /^[ \t\u00a0]*(?:bab|chapter)[ \t\u00a0]+(?=[IVXLCDMivxlcdm\d])/i;

export function romanOnlyChapters(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
): { changed: number; regions: VerifyRegion[] } {
  const regions: VerifyRegion[] = [];
  let changed = 0;
  nodes.forEach((n, i) => {
    const role = outline.items[i]?.role;
    if (n.kind !== "paragraph" || (role !== "chapter" && role !== "chapter_label")) return;
    const m = WORD_RE.exec(leadingRaw(n.xml, 20));
    if (!m) return;
    const before = paragraphLine(n.xml);
    n.xml = removeLeading(n.xml, m[0].length);
    n.text = n.text.replace(/^\s*(?:bab|chapter)\s+/i, "");
    const after = paragraphLine(n.xml);
    regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: "", afterText: "" });
    changed++;
  });
  return { changed, regions };
}
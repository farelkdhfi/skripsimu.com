/**
 * Paragraf lanjutan bukan alinea baru → tanpa indentasi baris pertama.
 *
 * Paragraf yang diawali HURUF KECIL dan menyambung sesuatu yang belum selesai:
 *  - setelah rumus ("dengan Y adalah …", "st(i, j) pada Persamaan 2.1 adalah …")
 *  - setelah daftar, tabel, atau gambar
 *  - setelah baris "Nama : …" / kalimat berakhiran ":" atau "," ("menyatakan bahwa …")
 * Alinea baru (diawali huruf kapital) tetap menjorok walau hanya satu baris.
 * Paragraf berhuruf kecil setelah kalimat yang sudah selesai (berakhiran titik) tidak diubah.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

const ENDED_RE = /[.!?)"\u201d']$/;

export function unindentContinuations(nodes: (ParaNode | TableNode)[], outline: Outline, headings: Set<number> = new Set()): number {
  let count = 0;
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph") return;
    const it = outline.items[i];
    if (!it || it.role !== "body" || it.zone === "cover") return;
    if (n.hasNumbering || n.hasDrawing) return;
    if (!/^[a-z]/.test(n.text.replace(/^[\s\u00a0]+/, ""))) return;

    let k = i - 1;
    while (k >= 0) {
      const b = nodes[k];
      if (b.kind === "paragraph" && !b.text.trim() && !b.hasDrawing && !/<m:oMath\b/.test(b.xml)) { k--; continue; }
      break;
    }
    if (k < 0) return;
    const prev = nodes[k];
    // Paragraf di bawah judul (termasuk judul tingkat 4) adalah alinea baru.
    if (prev.kind === "paragraph" && headings.has(prev.id)) return;
    let continues = prev.kind === "table";
    if (prev.kind === "paragraph") {
      const pt = prev.text.trim();
      const pEnd = prev.xml.indexOf("</w:pPr>");
      const centered = pEnd !== -1 && /<w:jc w:val="center"\/>/.test(prev.xml.slice(0, pEnd));
      // Rumus ketikan ("g = (Spost − Spre) / (Smaks − Spre)") bisa berakhiran ")" → tetap rumus.
      const typedFormula = pt.length <= 150 && (/\s[=≈≤≥<>]\s/.test(pt) || centered);
      continues = typedFormula || /<m:oMath\b/.test(prev.xml) || prev.hasNumbering || prev.hasDrawing ||
        /[:;,]$/.test(pt) || !ENDED_RE.test(pt);
    }
    if (!continues) return;

    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
    const ind = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    if (!/w:firstLine="[1-9]/.test(ind)) return;
    const left = /w:(?:left|start)="(-?\d+)"/.exec(ind)?.[1] ?? "0";
    const right = /w:(?:right|end)="(-?\d+)"/.exec(ind)?.[1] ?? "0";
    n.xml = withProps(n.xml, "w:pPr", {
      "w:ind": `<w:ind w:left="${left}" w:right="${right}" w:firstLine="0"/>`,
    }, PPR_ORDER);
    count++;
  });
  return count;
}
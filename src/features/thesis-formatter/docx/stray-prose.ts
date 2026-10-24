/**
 * Kalimat isi yang "nyasar" rata tengah/kanan di isi bab
 * ("Angket diisi secara daring … pada jam kuliah.") dikembalikan menjadi alinea biasa.
 * Baris yang memang sengaja di tengah tidak disentuh: rumus, sumber/keterangan gambar & tabel,
 * kutipan/epigraf, baris bertab, butir daftar, dan baris yang menempel ke gambar/tabel.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

const NOTE_RE = /^(sumber|keterangan|ket\.|catatan|note|source)\b/i;
const MATHY_RE = /[=≈≤≥<>±×÷∑√∫]|\(\d+(?:\.\d+)*\)\s*$/;
const NEAR_OBJECT = new Set(["figure", "figure_caption", "table_caption"]);

export function realignStrayProse(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  o: { line: number; firstLine: number },
): number {
  const blank = (n: ParaNode | TableNode) => n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing;
  const neighbor = (i: number, step: 1 | -1) => {
    for (let k = i + step; k >= 0 && k < nodes.length; k += step) if (!blank(nodes[k])) return k;
    return -1;
  };
  const nearObject = (k: number) =>
    k !== -1 && (nodes[k].kind === "table" || NEAR_OBJECT.has(outline.items[k]?.role ?? ""));

  let count = 0;
  nodes.forEach((p, i) => {
    if (p.kind !== "paragraph") return;
    const it = outline.items[i];
    if (!it || it.zone !== "main" || it.role !== "body") return;
    if (p.hasDrawing || p.hasNumbering || p.hasField || p.tabCount > 0 || /<m:oMath\b/.test(p.xml)) return;

    const pEnd = p.xml.indexOf("</w:pPr>");
    const pPr = pEnd === -1 ? "" : p.xml.slice(0, pEnd);
    if (!/<w:jc w:val="(?:center|right|end)"\/>/.test(pPr)) return;

    const t = p.text.replace(/\s+/g, " ").trim();
    if (t.split(" ").length < 8 || !/^[A-Z]/.test(t) || !/[.!?]$/.test(t)) return;
    if (NOTE_RE.test(t) || MATHY_RE.test(t)) return;

    // Kutipan/epigraf: seluruh teks miring.
    const runs = [...p.xml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)].filter((m) => /<w:t[\s>]/.test(m[1]));
    if (runs.length && runs.every((m) => /<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(m[1]))) return;

    // Menempel ke gambar/tabel → keterangan objek.
    if (nearObject(neighbor(i, -1)) || nearObject(neighbor(i, 1))) return;

    p.xml = withProps(p.xml, "w:pPr", {
      "w:spacing": `<w:spacing w:before="0" w:after="0" w:line="${o.line}" w:lineRule="auto"/>`,
      "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="${o.firstLine}"/>`,
      "w:jc": `<w:jc w:val="both"/>`,
    }, PPR_ORDER);
    count++;
  });
  return count;
}
/**
 * Spasi baris "Exactly" (pedoman yang menulis "tepat 24 pt", mis. UNPAD).
 * Spasi kelipatan (auto) paragraf diubah menjadi pasti: line/240 × ukuran huruf isi.
 * Dilewati: sampul, gambar inline / rumus / objek (tinggi baris harus mengikuti isinya),
 * paragraf yang memuat huruf lebih besar dari teks isi, dan spasi yang sudah exact/atLeast.
 * Gambar MELAYANG tidak memengaruhi tinggi baris, jadi paragrafnya tetap diubah.
 */
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

const INLINE_OBJECT = /<wp:inline\b|<w:object[\s>]|<w:pict[\s>]|<m:oMath/;

export function applyExactSpacing(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    opts: { halfPts: number },
): number {
    let count = 0;
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || !it || it.zone === "cover" || it.role === "figure") return;
        if (INLINE_OBJECT.test(n.xml)) return;

        const s = n.xml.indexOf("<w:pPr>");
        const e = s === -1 ? -1 : n.xml.indexOf("</w:pPr>", s);
        if (s === -1 || e === -1) return;

        // Huruf lebih besar dari teks isi (judul besar) → tinggi baris exact akan memotong huruf.
        const runs = n.xml.slice(e);
        if ([...runs.matchAll(/<w:sz\s+w:val="(\d+)"/g)].some((m) => Number(m[1]) > opts.halfPts)) return;

        // w:spacing milik paragraf, bukan spasi karakter di dalam rPr / pPrChange.
        const pPr = n.xml.slice(s, e);
        const masked = pPr.replace(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, (m) => " ".repeat(m.length));
        const m = /<w:spacing\b[^>]*\/>/.exec(masked);
        if (!m) return;
        const tag = pPr.slice(m.index, m.index + m[0].length);
        const line = /w:line="(\d+)"/.exec(tag);
        if (!line) return;
        if ((/w:lineRule="(\w+)"/.exec(tag)?.[1] ?? "auto") !== "auto") return;

        const exact = Math.round((Number(line[1]) / 240) * opts.halfPts * 10);
        let next = tag.replace(/w:line="\d+"/, `w:line="${exact}"`);
        next = /w:lineRule=/.test(next)
            ? next.replace(/w:lineRule="\w+"/, `w:lineRule="exact"`)
            : next.replace(/\s*\/>$/, ` w:lineRule="exact"/>`);
        const at = s + m.index;
        n.xml = n.xml.slice(0, at) + next + n.xml.slice(at + m[0].length);
        count++;
    });
    return count;
}
/**
 * Judul lampiran sebaris ("Lampiran 1. Hasil Analisis …") di bagian akhir dirapikan:
 * rata kiri dari margin, baris lanjutan menjorok sejajar teks judul (setelah "Lampiran 1. "),
 * spasi tunggal, menempel ke isi lampiran di bawahnya. Teks tidak diubah.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

const TITLE_RE = /^((?:lampiran|appendix)\s+(?:\d{1,3}|[A-Z])\s*[.:]?\s+)(\S[\s\S]*)$/i;

function textWidth(s: string, halfPts: number): number {
    const em = halfPts * 10;
    let w = 0;
    for (const ch of s) {
        if (/[ .,:;!'|iljtfrI1()]/.test(ch)) w += 0.3;
        else if (/[mwMW]/.test(ch)) w += 0.8;
        else if (/[A-Z]/.test(ch)) w += 0.68;
        else w += 0.5;
    }
    return Math.round(w * em);
}

export function formatAppendixTitles(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    opts: { halfPts: number; exact: boolean },
): number {
    let count = 0;
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || !it || it.zone !== "back" || it.role !== "body") return;
        if (n.hasNumbering || n.hasDrawing || n.tabCount) return;
        const t = n.text.replace(/\s+/g, " ").trim();
        const m = TITLE_RE.exec(t);
        if (!m || t.length > 250) return;
        // "Lampiran 1 menunjukkan …" = kalimat biasa, bukan judul.
        if (!/^[A-Z0-9("“]/.test(m[2])) return;
        if (/[.!?]$/.test(m[2]) && t.length > 150) return;
        const hang = Math.min(2268, textWidth(m[1], opts.halfPts) + 60);
        n.xml = withProps(n.xml, "w:pPr", {
            "w:keepNext": "<w:keepNext/>",
            "w:spacing": `<w:spacing w:before="0" w:beforeAutospacing="0" w:after="240" w:afterAutospacing="0" w:line="240" w:lineRule="${opts.exact ? "exact" : "auto"}"/>`,
            "w:ind": `<w:ind w:left="${hang}" w:right="0" w:hanging="${hang}"/>`,
            "w:jc": `<w:jc w:val="left"/>`,
        }, PPR_ORDER);
        count++;
    });
    return count;
}
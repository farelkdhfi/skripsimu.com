/**
 * Jarak nomor subbab ke judulnya diseragamkan. "1.1⇥Latar Belakang" memakai tab, dan posisi
 * tab aslinya sering beda-beda (sisa penomoran otomatis Word, tab bawaan): satu judul mulai
 * di 1,25 cm, yang lain di 2,27 cm. Semua judul bernomor+tab diberi SATU tab stop di
 * indentasi alinea (kelipatan berikutnya bila nomornya panjang), jadi teks judul sejajar
 * awal alinea.
 */
import { withProps, PPR_ORDER } from "./xml";
import { leadingRaw } from "./lists";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

const NUM_TAB_RE = /^[ \u00a0]*(\d+(?:\.\d+)+\.?)[ \u00a0]*\t/;

/** Lebar kira-kira nomor tebal ("2.1.10") dalam twip. */
function numberWidth(num: string, halfPts: number): number {
    const em = halfPts * 10;
    let w = 0;
    for (const ch of num) w += /\d/.test(ch) ? 0.5 : 0.25;
    return Math.round(w * em * 1.05);
}

export function alignHeadingTabs(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    opts: { indent: number; halfPts: number },
): number {
    let count = 0;
    const step = Math.max(284, Math.round(opts.indent / 2));
    nodes.forEach((n, i) => {
        const role = outline.items[i]?.role;
        if (n.kind !== "paragraph" || (role !== "section" && role !== "subsection")) return;
        const m = NUM_TAB_RE.exec(leadingRaw(n.xml, 40));
        if (!m) return;
        const need = numberWidth(m[1], opts.halfPts) + 90;
        let pos = Math.max(opts.indent, step);
        while (pos < need) pos += step;
        const tabs = `<w:tabs><w:tab w:val="left" w:pos="${pos}"/></w:tabs>`;
        const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
        if ((/<w:tabs>[\s\S]*?<\/w:tabs>/.exec(pPr)?.[0] ?? "") === tabs) return;
        n.xml = withProps(n.xml, "w:pPr", { "w:tabs": tabs }, PPR_ORDER);
        count++;
    });
    return count;
}
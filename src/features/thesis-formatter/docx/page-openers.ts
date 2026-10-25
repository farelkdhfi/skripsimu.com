/**
 * Pengaman pindah halaman. Judul pembuka halaman (bab, daftar, pustaka, lampiran) yang SUDAH
 * membuka halaman baru dicatat sebelum modul penyusun ulang berjalan (daftar isi kolom, tanda
 * tangan, sampul, halaman kosong). Bila sesudahnya pindah halamannya ikut terhapus (mis.
 * paragraf kosong berisi page break dimakan penyusun Daftar Isi), judul itu diberi
 * pageBreakBefore. Judul yang dari awal tidak membuka halaman tidak disentuh.
 */
import { flatten, type BlockNode, type ParaNode, type TableNode } from "./document-model";
import { withProps, PPR_ORDER } from "./xml";
import { startsNewPage, type Role, type Zone } from "../structure/roles";

export type PageInfo = Map<ParaNode | TableNode, { role: Role; zone: Zone }>;

const PAGE_BR = /<w:br\b[^>]*w:type="page"/;
const DRAWING = /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/;

function pPrOf(xml: string): string {
    const e = xml.indexOf("</w:pPr>");
    return e === -1 ? "" : xml.slice(0, e);
}

function opensItself(p: ParaNode): boolean {
    if (/<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPrOf(p.xml))) return true;
    const br = p.xml.search(PAGE_BR);
    if (br === -1) return false;
    const t = p.xml.search(/<w:t[\s>]/);
    return t === -1 || br < t;
}

function breaksAfter(p: ParaNode): boolean {
    if (PAGE_BR.test(p.xml)) return true;
    const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(p.xml)?.[0];
    return Boolean(sect) && !/<w:type\s+w:val="continuous"/.test(sect!);
}

export function openedHeadings(body: BlockNode[], info: PageInfo): Set<ParaNode> {
    const out = new Set<ParaNode>();
    const list = flatten(body);
    list.forEach((n, i) => {
        if (n.kind !== "paragraph") return;
        const it = info.get(n);
        if (!it || it.zone === "cover" || !startsNewPage(it.role)) return;
        if (opensItself(n)) { out.add(n); return; }
        for (let k = i - 1; k >= 0; k--) {
            const p = list[k];
            if (p.kind === "table") return;
            if (breaksAfter(p)) { out.add(n); return; }
            if (p.text.trim() || DRAWING.test(p.xml)) return;
        }
        out.add(n);
    });
    return out;
}

export function restoreOpeners(body: BlockNode[], before: Set<ParaNode>, info: PageInfo): number {
    const now = openedHeadings(body, info);
    const alive = new Set(flatten(body));
    let count = 0;
    for (const p of before) {
        if (now.has(p) || !alive.has(p)) continue;
        p.xml = withProps(p.xml, "w:pPr", { "w:pageBreakBefore": "<w:pageBreakBefore/>" }, PPR_ORDER);
        count++;
    }
    return count;
}
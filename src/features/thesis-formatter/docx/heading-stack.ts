/**
 * Judul yang langsung diikuti judul lain (3.1 → 3.1.1 → A. → (1), atau BAB → 1.1) tidak
 * menumpuk jarak: jarak tambahan di atas judul hanya dipakai bila di atasnya paragraf teks.
 * Jarak di bawah judul sebelumnya (mis. setelah judul bab) tetap berlaku.
 */
import { withProps, PPR_ORDER } from "./xml";
import { flatten, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { PageInfo } from "./page-openers";
import type { Role } from "../structure/roles";

const STACK = new Set<Role>(["chapter", "chapter_label", "section", "subsection"]);
const TARGET = new Set<Role>(["section", "subsection"]);

function zeroBefore(p: ParaNode): boolean {
    const pPr = (/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "").replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "");
    const tag = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "<w:spacing/>";
    if (/w:before="0"/.test(tag) && !/w:beforeLines=|w:beforeAutospacing="(?:1|true|on)"/.test(tag)) return false;
    const next = tag
        .replace(/\s*w:before(?:Autospacing|Lines)?="[^"]*"/g, "")
        .replace(/^<w:spacing/, `<w:spacing w:before="0" w:beforeAutospacing="0"`);
    p.xml = withProps(p.xml, "w:pPr", { "w:spacing": next }, PPR_ORDER);
    return true;
}

export function collapseStackedHeadings(model: DocModel, info: PageInfo, extra: Set<ParaNode>): number {
    const list = flatten(model.body);
    const roleOf = (n: ParaNode | TableNode) => info.get(n)?.role;
    const isHead = (n: ParaNode | TableNode) =>
        n.kind === "paragraph" && (extra.has(n) || STACK.has(roleOf(n)!));
    let count = 0;

    list.forEach((n, i) => {
        if (n.kind !== "paragraph" || info.get(n)?.zone !== "main") return;
        if (!extra.has(n) && !TARGET.has(roleOf(n)!)) return;
        for (let k = i - 1; k >= 0; k--) {
            const p = list[k];
            if (p.kind === "table") return;
            // Pindah halaman / section di antaranya: judul memang membuka halaman baru, biarkan.
            if (/<w:sectPr\b|<w:br\b[^>]*w:type="page"/.test(p.xml)) return;
            if (!p.text.trim() && !/<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/.test(p.xml)) continue;
            if (isHead(p) && zeroBefore(n)) count++;
            return;
        }
    });
    return count;
}
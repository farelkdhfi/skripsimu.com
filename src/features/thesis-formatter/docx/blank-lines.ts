/**
 * Baris kosong hasil Enter dibuang; jaraknya diatur formatter lewat spasi paragraf.
 *  1. Tepat setelah judul: bab, subbab, Kata Pengantar, Abstrak, Daftar …, Pustaka, Lampiran.
 *  2. Di antara dua paragraf teks di bagian awal & isi (Kata Pengantar, Abstrak, bab),
 *     dan sebelum judul subbab di dalam bab.
 *  3. Sebelum "Kata Kunci"/"Keywords": diganti jarak N baris (aturan preset `keywordGapLines`).
 * Tidak disentuh: sampul, ruang tanda tangan ("Sumedang, 8 Juni 2026" ↵↵ "Penulis"), blok judul
 * di atas Abstrak, dan paragraf kosong yang membawa section/pindah halaman/bookmark/field/gambar.
 */
import { withProps, PPR_ORDER } from "./xml";
import { flatten, isEmptyParagraph, type BlockNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { PageInfo } from "./page-openers";
import type { Role } from "../structure/roles";

const HEADINGS = new Set<Role>([
    "front_heading", "list_heading", "chapter", "chapter_label",
    "section", "subsection", "reference_heading", "appendix_heading",
]);
const KEYWORDS_RE = /^\s*(?:kata[\s-]*kunci|key\s*words?)\s*:/i;
const OBJECT_RE = /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]|<m:oMath/;
/** Panjang minimal paragraf teks (baris tanggal, "Penulis", nama penanda tangan lebih pendek). */
const PROSE_MIN = 40;

/** Paragraf kosong biasa (hasil Enter) yang aman dibuang. */
function isGap(n: ParaNode | TableNode): n is ParaNode {
    return n.kind === "paragraph" && isEmptyParagraph(n) && !OBJECT_RE.test(n.xml) &&
        !/<w:sectPr\b/.test(n.xml) && !/<w:br\b[^>]*w:type="page"/.test(n.xml) &&
        !/<w:bookmarkStart\b|<w:fldChar\b|<w:instrText\b|<w:fldSimple\b/.test(n.xml);
}

const textOf = (n: ParaNode | TableNode | null) => (n && n.kind === "paragraph" ? n.text.replace(/\s+/g, " ").trim() : "");

/** Jarak sebelum paragraf = N × tinggi barisnya sendiri. */
function setGapBefore(p: ParaNode, lines: number, halfPts: number, exact: boolean) {
    const pPr = (/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "").replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "");
    const tag = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "<w:spacing/>";
    const line = Number(/w:line="(\d+)"/.exec(tag)?.[1] ?? 240);
    const rule = /w:lineRule="(\w+)"/.exec(tag)?.[1] ?? "auto";
    const lineTw = rule === "exact" || rule === "atLeast"
        ? line
        : Math.round((line / 240) * halfPts * 10 * (exact ? 1 : 1.15));
    const before = Math.round(lines * lineTw);
    const next = tag
        .replace(/\s*w:before(?:Autospacing|Lines)?="[^"]*"/g, "")
        .replace(/^<w:spacing/, `<w:spacing w:before="${before}" w:beforeAutospacing="0"`);
    p.xml = withProps(p.xml, "w:pPr", { "w:spacing": next }, PPR_ORDER);
}

function dropFrom(list: BlockNode[], drop: Set<ParaNode>): BlockNode[] {
    return list
        .filter((n) => !(n.kind === "paragraph" && drop.has(n)))
        .map((n) => {
            if (n.kind === "container") n.children = dropFrom(n.children, drop);
            return n;
        });
}

export function removeEnterGaps(
    model: DocModel,
    info: PageInfo,
    opts: { keywordGapLines: number | null; halfPts: number; exact: boolean },
): { removed: number; keywords: number } {
    const list = flatten(model.body);
    const drop = new Set<ParaNode>();
    const roleOf = (n: ParaNode | TableNode | null) => (n ? info.get(n)?.role : undefined);
    const zoneOf = (n: ParaNode | TableNode | null) => (n ? info.get(n)?.zone : undefined);
    const isHeading = (n: ParaNode | TableNode | null) => Boolean(n && n.kind === "paragraph" && HEADINGS.has(roleOf(n)!));
    let keywords = 0;

    // Kata kunci abstrak: baris kosong di atasnya diganti jarak N baris (juga bila tadinya tanpa Enter).
    if (opts.keywordGapLines !== null) {
        list.forEach((n, i) => {
            if (n.kind !== "paragraph" || zoneOf(n) !== "front" || !KEYWORDS_RE.test(textOf(n))) return;
            for (let k = i - 1; k >= 0 && isGap(list[k]); k--) drop.add(list[k] as ParaNode);
            setGapBefore(n, opts.keywordGapLines!, opts.halfPts, opts.exact);
            keywords++;
        });
    }

    let i = 0;
    while (i < list.length) {
        if (!isGap(list[i])) { i++; continue; }
        let j = i;
        while (j < list.length && isGap(list[j])) j++;
        const prev = i > 0 ? list[i - 1] : null;
        const next = j < list.length ? list[j] : null;
        const zone = zoneOf(list[i]) ?? zoneOf(prev) ?? "main";

        let remove = false;
        if (zone !== "cover") {
            if (isHeading(prev)) {
                remove = true;                                   // 1. setelah judul
            } else if (zone === "main" && isHeading(next)) {
                remove = true;                                   // sebelum judul subbab di dalam bab
            } else if ((zone === "front" || zone === "main") &&
                prev?.kind === "paragraph" && next?.kind === "paragraph" &&
                !(opts.keywordGapLines === null && KEYWORDS_RE.test(textOf(next))) &&
                textOf(prev).length >= PROSE_MIN && textOf(next).length >= PROSE_MIN) {
                remove = true;                                   // 2. di antara paragraf teks
            }
        }
        if (remove) for (let k = i; k < j; k++) drop.add(list[k] as ParaNode);
        i = j;
    }

    if (drop.size) model.body = dropFrom(model.body, drop);
    return { removed: drop.size, keywords };
}
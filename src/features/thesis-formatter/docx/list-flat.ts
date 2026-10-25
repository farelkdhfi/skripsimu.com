/**
 * Tata letak daftar "flat" (mis. pedoman UNPAD FAPET):
 *
 *   3.1.1    Bahan Penelitian
 *   A.       Perlakuan              ← butir daftar / judul daftar
 *   (1)      Ransum Basal           ← judul tingkat 4
 *   a.       Sub-butir              ← tingkat berapa pun: penanda tetap di margin
 *            Penelitian ini ...     ← paragraf di bawahnya = paragraf biasa (baris pertama menjorok)
 *   bahan baku pakan ...
 *
 * Semua penanda di margin kiri; teks (termasuk baris lanjutan) mulai di indentasi alinea.
 * Dijalankan SETELAH modul daftar & judul tingkat 4: indentasi paragraf menimpa indentasi
 * definisi numbering, jadi definisi daftar (dan preset lain) tidak perlu diubah.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import type { ListLayout } from "./lists";
import { isHeadingRole, type Role } from "../structure/roles";
import type JSZip from "jszip";

export function flattenLists(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    layout: ListLayout,
    minorIds: Set<number>,
    opts: { indent: number },
): number {
    let count = 0;
    const ind = opts.indent;
    nodes.forEach((n, i) => {
        if (n.kind !== "paragraph") return;
        const it = outline.items[i];
        if (!it || (it.zone !== "main" && it.zone !== "front")) return;
        const l = layout.get(n.id);
        const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";

        const marked = minorIds.has(n.id) ||
            Boolean(l && !("restore" in l) && l.hanging !== undefined) ||
            (n.hasNumbering && !isHeadingRole(it.role as Role));

        if (marked) {
            // Penanda di margin, teks & baris lanjutan di indentasi alinea.
            n.xml = withProps(n.xml, "w:pPr", {
                "w:tabs": `<w:tabs><w:tab w:val="left" w:pos="${ind}"/></w:tabs>`,
                "w:ind": `<w:ind w:left="${ind}" w:right="0" w:hanging="${ind}"/>`,
            }, PPR_ORDER);
            count++;
            return;
        }

        if (l && !("restore" in l) && l.hanging === undefined) {
            // Paragraf penjelasan di bawah butir/judul daftar → paragraf biasa (tidak menjorok satu blok).
            if (/<w:jc w:val="(center|right|end)"/.test(pPr) || /<m:oMath/.test(n.xml)) return;
            const fl = /<w:ind\b[^>]*w:firstLine="(\d+)"/.exec(pPr)?.[1] ?? "0";
            n.xml = withProps(n.xml, "w:pPr", {
                "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="${fl}"/>`,
            }, PPR_ORDER);
            count++;
        }
    });
    return count;
}

/**
 * Penanda daftar rata tengah/kanan (lvlJc center/right) diposisikan dengan titik tengah/ujung
 * kanannya di posisi penanda — begitu penanda ada di margin, "(1)" separuhnya menjorok ke
 * margin kiri. Penanda daftar di naskah selalu rata kiri. Hanya penomoran yang dipakai
 * paragraf bagian awal & isi yang diubah (daftar di tabel/lampiran tidak disentuh).
 */
export async function alignListMarkers(
    zip: JSZip,
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
): Promise<number> {
    const part = zip.file("word/numbering.xml");
    if (!part) return 0;
    let xml = await part.async("string");

    const used = new Set<string>();
    nodes.forEach((n, i) => {
        const z = outline.items[i]?.zone;
        if (n.kind !== "paragraph" || (z !== "main" && z !== "front")) return;
        const id = /<w:numPr>[\s\S]*?<w:numId\s+w:val="(\d+)"/.exec(n.xml)?.[1];
        if (id && id !== "0") used.add(id);
    });
    if (!used.size) return 0;

    let count = 0;
    const fix = (block: string) =>
        block.replace(/<w:lvlJc\s+w:val="(?:center|right|end)"\s*\/>/g, () => { count++; return `<w:lvlJc w:val="left"/>`; });
    const abstracts = new Set<string>();
    xml = xml.replace(/<w:num\b[^>]*\bw:numId="(\d+)"[^>]*>[\s\S]*?<\/w:num>/g, (block, id: string) => {
        if (!used.has(id)) return block;
        const a = /<w:abstractNumId\s+w:val="(\d+)"/.exec(block)?.[1];
        if (a) abstracts.add(a);
        return fix(block);
    });
    xml = xml.replace(/<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>[\s\S]*?<\/w:abstractNum>/g,
        (block, id: string) => (abstracts.has(id) ? fix(block) : block));
    if (count) zip.file("word/numbering.xml", xml);
    return count;
}
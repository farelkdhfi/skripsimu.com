/**
 * Penekanan (miring, tebal, garis bawah) yang tidak pada tempatnya.
 *  - Judul tidak pernah miring / bergaris bawah (run + style).
 *  - Paragraf isi yang seluruhnya miring / bergaris bawah dikembalikan normal;
 *    penekanan sebagian (istilah asing, dsb.) dipertahankan.
 *  - Butir daftar: miring seluruh butir dibuang; tebal/garis bawah seluruh butir
 *    hanya dipertahankan bila semua butir daftar itu memakainya (seragam).
 * Sampul, lampiran, catatan tabel, rumus, dan kutipan (style Quote) tidak disentuh.
 */
import type JSZip from "jszip";
import { withProps, RPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole } from "../structure/roles";

type Mark = "i" | "u" | "b";

const ON: Record<Mark, RegExp> = {
    i: /<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/,
    u: /<w:u\s+w:val="(?!none)[^"]*"[^>]*\/>/,
    b: /<w:b(?:\s+w:val="(?:1|true|on)")?\s*\/>/,
};
const CHAR_STYLE: Record<Mark, RegExp> = {
    i: /<w:rStyle w:val="(?:Emphasis|SubtleEmphasis|IntenseEmphasis)"\s*\/>/,
    u: /(?!)/,
    b: /<w:rStyle w:val="(?:Strong)"\s*\/>/,
};
const OFF: Record<Mark, Record<string, string>> = {
    i: { "w:i": `<w:i w:val="0"/>`, "w:iCs": `<w:iCs w:val="0"/>` },
    u: { "w:u": `<w:u w:val="none"/>` },
    b: { "w:b": `<w:b w:val="0"/>`, "w:bCs": `<w:bCs w:val="0"/>` },
};
const STRIP: Record<Mark, RegExp> = {
    i: /<w:i(?:Cs)?(?:\s[^>]*)?\/>/g,
    u: /<w:u(?:\s[^>]*)?\/>/g,
    b: /<w:b(?:Cs)?(?:\s[^>]*)?\/>/g,
};

const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const runText = (r: string) => [...r.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
const rPrOf = (r: string) => /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(r)?.[0] ?? "";

/** Porsi karakter teks paragraf yang memakai penekanan tertentu (0–1). */
function share(xml: string, mark: Mark): number {
    let all = 0;
    let on = 0;
    for (const m of xml.matchAll(RUN_RE)) {
        const len = runText(m[0]).replace(/\s/g, "").length;
        if (!len) continue;
        all += len;
        const rPr = rPrOf(m[0]);
        if (ON[mark].test(rPr) || CHAR_STYLE[mark].test(rPr)) on += len;
    }
    return all ? on / all : 0;
}

function strip(xml: string, mark: Mark): string {
    // Run teks
    let out = xml.replace(RUN_RE, (run) => {
        if (!/<w:t[\s>]/.test(run)) return run;
        const cleaned = run.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, (rPr) =>
            rPr.replace(STRIP[mark], "").replace(CHAR_STYLE[mark], ""));
        return withProps(cleaned, "w:rPr", OFF[mark], RPR_ORDER);
    });
    // Tanda paragraf (menentukan tampilan nomor/huruf daftar)
    out = out.replace(/(<w:pPr>[\s\S]*?)(<w:rPr>[\s\S]*?<\/w:rPr>)([\s\S]*?<\/w:pPr>)/, (_m, a: string, rPr: string, c: string) =>
        a + rPr.replace(STRIP[mark], "") + c);
    return out;
}

const isQuote = (xml: string) => /<w:pStyle w:val="(?:Quote|IntenseQuote|Kutipan[^"]*)"/.test(xml);
const isFormula = (p: ParaNode) =>
    /<m:oMath/.test(p.xml) ||
    (/<w:jc w:val="center"/.test(p.xml) && /=/.test(p.text) && p.text.length <= 150);
const numIdOf = (xml: string) => /<w:numPr>[\s\S]*?<w:numId w:val="(\d+)"/.exec(xml)?.[1] ?? null;

export function normalizeEmphasis(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    exclude: Set<number>,
): { headings: number; paragraphs: number; listItems: number } {
    const rep = { headings: 0, paragraphs: 0, listItems: 0 };
    const FULL = 0.9;

    // 1. Judul — semua zona kecuali sampul.
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || !it || it.zone === "cover" || !isHeadingRole(it.role)) return;
        const before = n.xml;
        n.xml = strip(strip(n.xml, "i"), "u");
        if (n.xml !== before) rep.headings++;
    });

    // 2. Paragraf isi & butir daftar — zona isi saja.
    type Item = { p: ParaNode; numId: string };
    const groups: Item[][] = [];
    let cur: Item[] = [];
    let curId: string | null = null;

    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph") return;
        if (!it || it.zone !== "main") { if (cur.length) groups.push(cur); cur = []; curId = null; return; }
        if (it.role !== "body" || exclude.has(n.id) || !n.text.trim() || isQuote(n.xml) || isFormula(n)) return;

        const numId = numIdOf(n.xml);
        if (numId) {
            if (numId !== curId && cur.length) { groups.push(cur); cur = []; }
            curId = numId;
            cur.push({ p: n, numId });
            return;
        }

        let changed = false;
        for (const mark of ["i", "u"] as const) {
            if (share(n.xml, mark) >= FULL) { n.xml = strip(n.xml, mark); changed = true; }
        }
        if (changed) rep.paragraphs++;
    });
    if (cur.length) groups.push(cur);

    // 3. Butir daftar: miring dibuang; tebal/garis bawah harus seragam satu daftar.
    // Butir daftar yang sama (numId sama) bisa terpisah paragraf penjelasan → digabung per numId.
    const byNum = new Map<string, ParaNode[]>();
    for (const g of groups) for (const x of g) {
        const list = byNum.get(x.numId) ?? [];
        list.push(x.p);
        byNum.set(x.numId, list);
    }
    for (const items of byNum.values()) {
        for (const p of items) {
            if (share(p.xml, "i") >= FULL) { p.xml = strip(p.xml, "i"); rep.listItems++; }
        }
        for (const mark of ["b", "u"] as const) {
            const full = items.map((p) => share(p.xml, mark) >= FULL);
            if (full.every(Boolean) || !full.some(Boolean)) continue;
            items.forEach((p, k) => { if (full[k]) { p.xml = strip(p.xml, mark); rep.listItems++; } });
        }
    }

    return rep;
}

/** Style judul, keterangan, dan daftar isi tidak boleh miring / bergaris bawah. */
export async function patchHeadingStyles(zip: JSZip): Promise<void> {
    const f = zip.file("word/styles.xml");
    if (!f) return;
    const xml = await f.async("string");
    const ids = /^(Heading[1-9](Char)?|Title(Char)?|Subtitle(Char)?|JudulBab|JudulAwal|JudulLampiran|Caption|TOC[1-9]|TableofFigures)$/;
    const out = xml.replace(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, (st) => {
        const id = /w:styleId="([^"]+)"/.exec(st)?.[1] ?? "";
        if (!ids.test(id)) return st;
        return st.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, (rPr) =>
            rPr.replace(/<w:i(?:Cs)?(?:\s[^>]*)?\/>/g, "").replace(/<w:u(?:\s[^>]*)?\/>/g, ""));
    });
    if (out !== xml) zip.file("word/styles.xml", out);
}
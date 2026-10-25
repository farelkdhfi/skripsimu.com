/**
 * Ketebalan judul daftar ("A. Perlakuan") & judul tingkat 4 ("(1) Ransum Basal") mengikuti
 * dokumen asli: bila ada bagian yang tebal — teks judul, penanda nomornya, atau style-nya —
 * seluruh judul DAN penandanya ditebalkan; bila tidak ada sama sekali, seluruhnya biasa.
 * Penanda nomor otomatis mengikuti format "akhir paragraf", jadi itu ikut diatur.
 */
import { mergeProps, withProps, PPR_ORDER, RPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole, type Role } from "../structure/roles";

const B_ON = /<w:b(?:\s+w:val="(?:1|true|on)")?\s*\/>/;
const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;

const numIdOf = (xml: string) => /<w:numPr>[\s\S]*?<w:numId\s+w:val="(\d+)"/.exec(xml)?.[1] ?? null;
const ilvlOf = (xml: string) => /<w:numPr>[\s\S]*?<w:ilvl\s+w:val="(\d+)"/.exec(xml)?.[1] ?? "0";
const ownPPr = (xml: string) => {
    const e = xml.indexOf("</w:pPr>");
    return e === -1 ? "" : xml.slice(0, e).replace(/<w:pPrChange\b[\s\S]*$/, "");
};
const rPrOf = (xml: string) => /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(xml)?.[1] ?? "";

/** Style (paragraf/karakter) yang menebalkan teks, termasuk lewat basedOn. */
export function styleBoldOf(stylesXml: string): (id: string | null | undefined) => boolean {
    const own = new Map<string, { bold: boolean | null; basedOn: string | null }>();
    for (const m of stylesXml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
        const r = rPrOf(m[2]);
        const b = /<w:b(?:\s+w:val="([^"]*)")?\s*\/>/.exec(r);
        own.set(m[1], {
            bold: b ? !/^(0|false|off)$/.test(b[1] ?? "") : null,
            basedOn: /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1] ?? null,
        });
    }
    return (id) => {
        let cur = id ?? null;
        for (let d = 0; cur && d < 10; d++) {
            const s = own.get(cur);
            if (!s) return false;
            if (s.bold !== null) return s.bold;
            cur = s.basedOn;
        }
        return false;
    };
}

function lvlBold(numbering: string, numId: string, ilvl: string): boolean {
    const lvlRe = new RegExp(`<w:lvl\\b[^>]*w:ilvl="${ilvl}"[^>]*>[\\s\\S]*?<\\/w:lvl>`);
    const num = new RegExp(`<w:num\\b[^>]*\\bw:numId="${numId}"[^>]*>[\\s\\S]*?<\\/w:num>`).exec(numbering)?.[0];
    if (!num) return false;
    const abs = /<w:abstractNumId\s+w:val="(\d+)"/.exec(num)?.[1];
    const absXml = abs
        ? new RegExp(`<w:abstractNum\\b[^>]*\\bw:abstractNumId="${abs}"[^>]*>[\\s\\S]*?<\\/w:abstractNum>`).exec(numbering)?.[0]
        : undefined;
    const lvl = lvlRe.exec(num)?.[0] ?? (absXml ? lvlRe.exec(absXml)?.[0] : undefined);
    return lvl ? B_ON.test(rPrOf(lvl)) : false;
}

/** Dicatat di awal, sebelum modul lain mengubah format: apakah paragraf tampil tebal di dokumen asli. */
export function originalBold(
    nodes: (ParaNode | TableNode)[],
    numbering: string,
    styleBold: (id: string | null | undefined) => boolean,
): Map<ParaNode, boolean> {
    const out = new Map<ParaNode, boolean>();
    for (const n of nodes) {
        if (n.kind !== "paragraph") continue;
        const pPr = ownPPr(n.xml);
        const body = n.xml.slice(Math.max(0, n.xml.indexOf("</w:pPr>")));
        const paraStyleBold = styleBold(/<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1]);
        const runBold = [...body.matchAll(RUN_RE)].map((m) => m[0]).filter((r) => /<w:t[\s>]/.test(r)).some((r) => {
            const rp = rPrOf(r);
            const direct = /<w:b(?:\s+w:val="([^"]*)")?\s*\/>/.exec(rp);
            if (direct) return !/^(0|false|off)$/.test(direct[1] ?? "");
            return styleBold(/<w:rStyle\s+w:val="([^"]+)"/.exec(rp)?.[1]) || paraStyleBold;
        });
        const id = numIdOf(n.xml);
        const markerBold = Boolean(id && id !== "0") && (B_ON.test(rPrOf(pPr)) || lvlBold(numbering, id!, ilvlOf(n.xml)));
        out.set(n, runBold || markerBold);
    }
    return out;
}

/** Tebal/biasa untuk semua run teks DAN penanda paragraf (penanda nomor otomatis ikut). */
function setBold(xml: string, on: boolean): string {
    const b = on ? "<w:b/>" : `<w:b w:val="0"/>`;
    const bCs = on ? "<w:bCs/>" : `<w:bCs w:val="0"/>`;
    const end = xml.indexOf("</w:pPr>");
    const head = end === -1 ? "" : xml.slice(0, end);
    const body = end === -1 ? xml : xml.slice(end);
    let out = head + body.replace(RUN_RE, (r) =>
        /<w:t[\s>]/.test(r) ? withProps(r, "w:rPr", { "w:b": b, "w:bCs": bCs }, RPR_ORDER) : r);
    const mark = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(ownPPr(out))?.[0] ?? null;
    out = withProps(out, "w:pPr", {
        "w:rPr": `<w:rPr>${mergeProps(mark, { "w:b": b, "w:bCs": bCs }, RPR_ORDER).join("")}</w:rPr>`,
    }, PPR_ORDER);
    return out;
}

/**
 * Judul tingkat 4 (dari modul judul kecil) dan butir daftar yang berfungsi sebagai judul:
 * pendek, tanpa tanda baca akhir, dan tidak diapit butir lain dari daftar yang sama
 * ("A. Perlakuan" berdiri sendiri di antara isi; daftar biasa punya saudara di sebelahnya).
 */
export function unifyHeadingWeight(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    minorIds: Set<number>,
    before: Map<ParaNode, boolean>,
): { changed: number; headings: Set<ParaNode> } {
    const headings = new Set<ParaNode>();
    let changed = 0;
    const near = (i: number, step: 1 | -1) => {
        for (let k = i + step; k >= 0 && k < nodes.length; k += step) {
            const m = nodes[k];
            if (m.kind === "table" || m.text.trim() || m.hasDrawing) return m;
        }
        return null;
    };

    nodes.forEach((n, i) => {
        if (n.kind !== "paragraph") return;
        const it = outline.items[i];
        if (!it || it.zone !== "main" || isHeadingRole(it.role as Role)) return;
        let heading = minorIds.has(n.id);
        if (!heading) {
            const id = numIdOf(n.xml);
            const t = n.text.replace(/\s+/g, " ").trim();
            if (id && id !== "0" && t && t.length <= 80 && t.split(" ").length <= 10 && !/[.;:,?!]$/.test(t)) {
                const sibling = (m: ParaNode | TableNode | null) => Boolean(m && m.kind === "paragraph" && numIdOf(m.xml) === id);
                heading = !sibling(near(i, -1)) && !sibling(near(i, 1));
            }
        }
        if (!heading) return;
        headings.add(n);
        const xml = setBold(n.xml, before.get(n) ?? false);
        if (xml !== n.xml) { n.xml = xml; changed++; }
    });
    return { changed, headings };
}
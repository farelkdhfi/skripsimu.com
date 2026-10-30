/**
 * Rumus dan kalimat sesudahnya dalam SATU paragraf ("Enter" lupa ditekan sesudah rumus):
 *
 *   [rumus Office Math]Selanjutnya, tingkat pertumbuhan jamur dinilai …
 *
 * Word menampilkan rumus seperti ini kecil di dalam kalimat (bukan rumus tampil di barisnya
 * sendiri). Dipecah menjadi dua paragraf: rumus, lalu kalimatnya sebagai alinea baru. Teks dan
 * rumus tidak berubah; hanya batas paragrafnya.
 *
 * Syarat (ketat): paragraf diawali rumus (selain bookmark/run kosong), sesudah rumus ada teks yang
 * diawali huruf kapital dan berbentuk kalimat (≥ 3 kata), dan tidak ada rumus lagi di kalimat itu.
 */
import { parseBlocks, type BlockNode, type DocModel, type ParaNode } from "./document-model";
import { innerXml, openTag, splitElements } from "./xml";
import { maxId } from "./figure-split";
import { paragraphLine, type VerifyRegion } from "./verify";

const textOf = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join("");

export function splitEquationSentences(model: DocModel, skip: Set<ParaNode> = new Set()): { count: number; regions: VerifyRegion[] } {
    const out = { count: 0, regions: [] as VerifyRegion[] };
    let nextId = maxId(model.body) + 1;

    const tryOne = (p: ParaNode): ParaNode[] | null => {
        if (!/<m:oMath\b/.test(p.xml) || /<w:txbxContent\b|<w:drawing\b|<w:pict\b/.test(p.xml)) return null;
        const parts = splitElements(innerXml(p.xml));
        const pPr = parts.find((c) => c.tag === "w:pPr")?.xml ?? "";
        const content = parts.filter((c) => c.tag !== "w:pPr");
        // Awal paragraf: hanya bookmark / run kosong sebelum rumus.
        let k = 0;
        const blank = (c: { tag: string; xml: string }) =>
            /^w:(bookmarkStart|bookmarkEnd|proofErr|permStart|permEnd)$/.test(c.tag) ||
            (c.tag === "w:r" && !textOf(c.xml).trim() && !/<w:(tab|br|drawing|pict|object|sym)\b/.test(c.xml));
        while (k < content.length && blank(content[k])) k++;
        if (k >= content.length || !/^m:oMath(Para)?$/.test(content[k].tag)) return null;
        let e = k + 1;
        while (e < content.length && /^m:oMath(Para)?$/.test(content[e].tag)) e++;
        const rest = content.slice(e);
        if (rest.some((c) => /^m:oMath/.test(c.tag) || /<m:oMath\b/.test(c.xml))) return null;
        const sentence = rest.map((c) => (c.tag === "w:r" ? textOf(c.xml) : c.tag === "w:hyperlink" ? textOf(c.xml) : "")).join("").trim();
        if (!/^[A-ZÀ-Þ]/.test(sentence) || sentence.split(/\s+/).length < 3) return null;
        // Paragraf sesudahnya membawa sectPr (bila ada); paragraf rumus tidak.
        const open = openTag(p.xml);
        const pPrNoSect = pPr.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, "");
        const a = `${open}${pPrNoSect}${content.slice(0, e).map((c) => c.xml).join("")}</w:p>`;
        const b = `<w:p>${pPr}${rest.map((c) => c.xml).join("")}</w:p>`;
        const pa = parseBlocks(a, p.id)[0] as ParaNode;
        const pb = parseBlocks(b, nextId++)[0] as ParaNode;
        return [pa, pb];
    };

    const walk = (list: BlockNode[]) => {
        for (let i = 0; i < list.length; i++) {
            const n = list[i];
            if (n.kind === "container") { walk(n.children); continue; }
            if (n.kind !== "paragraph" || skip.has(n)) continue;
            const before = paragraphLine(n.xml);
            const r = tryOne(n);
            if (!r) continue;
            list.splice(i, 1, ...r);
            out.count++;
            out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: r.map((x) => paragraphLine(x.xml)).filter(Boolean), beforeText: "", afterText: "" });
            i++;
        }
    };
    walk(model.body);
    return out;
}
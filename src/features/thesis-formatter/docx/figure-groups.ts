/**
 * Kelompok gambar buatan tangan → gambar + keterangan biasa.
 *
 * Susunan yang dikenali (di badan teks maupun di sel tabel), dan HANYA bila semua syarat terpenuhi:
 *   [foto melayang sederhana di paragraf tanpa teks]
 *   [0…n paragraf kosong pengganjal]
 *   [paragraf berisi KOTAK TEKS melayang yang isinya keterangan "Ilustrasi/Gambar … N. …"]
 *     (boleh sekaligus memuat baris "Sumber: …")
 *   [0…n paragraf kosong] [opsional "Sumber: …"]
 * Hasil: foto inline di tengah → keterangan sebagai paragraf biasa (field SEQ & bookmark Daftar
 * Ilustrasi ikut) → baris Sumber. Pengganjal dibuang. Posisi keterangan tidak lagi bergantung pada
 * jarak paragraf, jadi tidak bisa tergeser/terpotong oleh margin, spasi, atau style baru.
 *
 * Teks lain di dalam kotak keterangan yang TIDAK terlihat (kotak berukuran tetap hanya memuat
 * baris pertama; sisanya sisa ketikan lama) tidak dimunculkan — dilaporkan. Bila tidak bisa
 * dipastikan mana yang terlihat (kotak membesar mengikuti isi), kelompok dibiarkan (tidak diubah).
 */
import { isEmptyParagraph, parseBlocks, type BlockNode, type DocModel, type ParaNode } from "./document-model";
import { innerXml, splitElements, withProps, PPR_ORDER } from "./xml";
import { inlineOf, maxId } from "./figure-split";
import { fingerprint, paragraphLine, type VerifyRegion } from "./verify";

export type FigureGroupReport = {
    converted: number;
    /** Kelompok yang dikenali tetapi dibiarkan (ragu). */
    skipped: number;
    hidden: { caption: string; text: string }[];
    /** Objek gambar yang dibuang (kotak teks keterangan; Choice + Fallback dihitung seperti di verify). */
    removedDrawings: number;
    regions: VerifyRegion[];
};

/** Keterangan gambar: diawali label + nomor di AWAL paragraf, pendek. Kalimat isi yang hanya MENYEBUT "Ilustrasi 5" tidak lolos. */
export const FIGURE_CAPTION_RE = /^\s*(?:ilustrasi|gambar|grafik|foto|figure|bagan|diagram|peta)\s+\d+(?:[.\-]\d+)*\s*[.:–—-]?\s*\S/i;
const SOURCE_RE = /^\s*sumber\s*[:：]/i;

/** Run tingkat atas paragraf (run di dalam kotak teks tidak ikut — regex biasa salah menutup di run dalam). */
const topRuns = (p: ParaNode): string[] => splitElements(innerXml(p.xml)).filter((c) => c.tag === "w:r").map((c) => c.xml);
const textOf = (xml: string) => paragraphLine(xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "").replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, ""));
const countDrawings = (xml: string) => (xml.match(/<w:drawing[\s>]/g) ?? []).length + (xml.match(/<w:pict[\s>]/g) ?? []).length + (xml.match(/<w:object[\s>]/g) ?? []).length;

const isEmptyPara = (n: BlockNode): n is ParaNode =>
    n.kind === "paragraph" && isEmptyParagraph(n) && !n.hasSectPr;

/** Run berisi SATU foto melayang sederhana (bukan grup/shape/kotak teks/AlternateContent). */
function floatingPictureRun(p: ParaNode): string | null {
    if (textOf(p.xml).trim()) return null;
    const runs = topRuns(p).filter((r) => /<w:drawing\b|<w:pict\b|<w:object\b/.test(r));
    if (runs.length !== 1) return null;
    const r = runs[0];
    if (/<mc:AlternateContent\b|<wps:|<wpg:|<w:txbxContent\b|<w:pict\b|<w:object\b/.test(r)) return null;
    if ((r.match(/<wp:anchor\b/g) ?? []).length !== 1 || !/<pic:pic\b/.test(r)) return null;
    return r;
}

type Caption = { run: string; captionXml: string; captionText: string; hidden: string[] };

/** Run kotak teks keterangan di paragraf ini (satu-satunya objek di paragraf). */
function captionBox(p: ParaNode): Caption | "ambiguous" | null {
    const runs = topRuns(p).filter((r) => /<w:drawing\b|<w:pict\b|<w:object\b/.test(r));
    if (runs.length !== 1) return null;
    const run = runs[0];
    const main = run.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "");
    if (!/<w:txbxContent\b/.test(main) || /<pic:pic\b|<wpg:/.test(main)) return null;
    const boxes = main.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g) ?? [];
    if (boxes.length !== 1) return null;
    const paras = parseBlocks(boxes[0].replace(/^<w:txbxContent>|<\/w:txbxContent>$/g, ""), -1)
        .filter((b): b is ParaNode => b.kind === "paragraph");
    const filled = paras.filter((q) => textOf(q.xml).trim());
    if (!filled.length) return null;
    const first = filled[0];
    const text = textOf(first.xml).trim();
    if (!FIGURE_CAPTION_RE.test(text) || text.length > 250 || /<w:drawing\b|<w:pict\b/.test(first.xml)) return null;
    const rest = filled.slice(1).map((q) => textOf(q.xml).trim());
    // Kotak yang membesar mengikuti isi menampilkan SEMUA baris → tidak bisa dipastikan mana yang dimaksud.
    if (rest.length && /<a:spAutoFit\s*\/>|mso-fit-shape-to-text:\s*t/.test(run)) return "ambiguous";
    return { run, captionXml: first.xml, captionText: text, hidden: rest };
}

export function normalizeFigureGroups(model: DocModel): FigureGroupReport {
    const rep: FigureGroupReport = { converted: 0, skipped: 0, hidden: [], removedDrawings: 0, regions: [] };
    let nextId = maxId(model.body) + 1;

    const processList = (list: BlockNode[]) => {
        for (let i = 0; i < list.length; i++) {
            const p = list[i];
            if (p.kind !== "paragraph") continue;
            const picRun = floatingPictureRun(p);
            if (!picRun) {
                // Foto + kotak keterangan + teks dalam SATU paragraf: urutan tampilnya hanya bisa ditebak dari
                // koordinat yang berbeda acuannya (margin vs paragraf) → ragu, dibiarkan.
                const q = p.xml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "");
                if (/<wp:anchor\b[\s\S]*?<pic:pic\b/.test(q) && /<w:txbxContent\b/.test(q) && FIGURE_CAPTION_RE.test(paragraphLine((/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/.exec(q)?.[0] ?? "").replace(/<\/w:p>[\s\S]*$/, "")))) rep.skipped++;
                continue;
            }

            // Cari paragraf kotak keterangan sesudah foto: hanya paragraf kosong di antaranya.
            let k = i + 1;
            while (k < list.length && isEmptyPara(list[k])) k++;
            const holder = list[k];
            if (!holder || holder.kind !== "paragraph") continue;
            const holderText = textOf(holder.xml).trim();
            if (holderText && !SOURCE_RE.test(holderText)) continue;
            const cap = captionBox(holder);
            if (!cap) continue;
            if (cap === "ambiguous") { rep.skipped++; continue; }

            // Pengganjal sesudah kotak keterangan sampai baris Sumber (bila ada) juga dibuang.
            let s = k + 1;
            while (s < list.length && isEmptyPara(list[s])) s++;
            const sourceAfter = !holderText && list[s]?.kind === "paragraph" && SOURCE_RE.test(textOf((list[s] as ParaNode).xml));
            const dropTo = sourceAfter ? s : k + 1;          // [k+1, dropTo) dibuang bila kosong

            // 1) Foto → inline, paragraf foto di tengah & tidak terpotong spasi "exact".
            let picXml = p.xml.replace(picRun, inlineOf(picRun));
            picXml = withProps(picXml, "w:pPr", {
                "w:keepNext": "<w:keepNext/>",
                "w:spacing": `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`,
                "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
                "w:jc": `<w:jc w:val="center"/>`,
                "w:framePr": null,
            }, PPR_ORDER);
            Object.assign(p, { ...(parseBlocks(picXml, 0)[0] as ParaNode), id: p.id });

            // 2) Keterangan dari kotak teks → paragraf biasa tepat di bawah foto.
            let capXml = withProps(cap.captionXml, "w:pPr", { "w:keepNext": sourceAfter || holderText ? "<w:keepNext/>" : undefined, "w:jc": `<w:jc w:val="center"/>` }, PPR_ORDER);
            capXml = capXml.replace(/\sw14:(?:paraId|textId)="[^"]*"/g, "");
            const caption = parseBlocks(capXml, nextId++)[0] as ParaNode;

            // 3) Kotak teks dibuang dari paragrafnya; paragraf yang jadi kosong ikut dibuang.
            const holderXml = holder.xml.replace(cap.run, "");
            const removedLines = fingerprint(cap.run).text.split("\n").filter(Boolean);
            rep.removedDrawings += countDrawings(cap.run);
            rep.regions.push({ beforeLines: removedLines, afterLines: [paragraphLine(caption.xml)].filter(Boolean), beforeText: "", afterText: "" });
            for (const h of cap.hidden) rep.hidden.push({ caption: cap.captionText, text: h });

            const keepHolder = Boolean(holderText);
            if (keepHolder) Object.assign(holder, { ...(parseBlocks(holderXml, 0)[0] as ParaNode), id: holder.id });

            // Susun ulang daftar: [foto, keterangan, (holder bila berisi Sumber)], pengganjal dibuang.
            const tail: BlockNode[] = [];
            if (keepHolder) tail.push(holder);
            for (let m = k + 1; m < dropTo; m++) if (!isEmptyPara(list[m])) tail.push(list[m]);
            list.splice(i + 1, dropTo - (i + 1), caption, ...tail);
            rep.converted++;
            i += 1 + tail.length;
        }
    };

    const walk = (list: BlockNode[]) => {
        processList(list);
        for (const n of list) {
            if (n.kind === "container") walk(n.children);
            else if (n.kind === "table") for (const r of n.rows) for (const c of r.cells) walk(c.children);
        }
    };
    walk(model.body);
    return rep;
}
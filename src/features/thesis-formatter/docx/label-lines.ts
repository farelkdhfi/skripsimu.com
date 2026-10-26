/**
 * Baris "Label : isi" yang berurutan (identitas di lembar persetujuan/pernyataan,
 * "Nomor/Lampiran/Hal" di surat, Daftar Singkatan) disejajarkan:
 * titik dua satu garis lewat satu tab stop, baris isi yang panjang menjorok
 * sejajar setelah ": ". Hanya di bagian awal & lampiran; minimal 2 baris berurutan.
 */
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { leadingRaw, removeLeading } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";

/** Label 1–30 karakter, titik dua, isi. Label tidak boleh memuat titik dua / angka jam. */
const LV_RAW_RE = /^[ \t\u00a0]*([A-Za-z\u00C0-\u024F\u0370-\u03FF][^:\t\n]{0,29}?)[ \t\u00a0]*[:：][ \t\u00a0]*(?=\S)/;

/**
 * Tab & ganti baris manual di DALAM isi label (dipakai penulis untuk memaksa judul panjang turun
 * baris: "… Postbiotik [tab]terhadap …") diganti spasi biasa — atau dibuang bila sudah berspasi —
 * agar isi mengalir dan baris lanjutannya rata di bawah awal isi (indentasi gantung).
 */
function flattenValue(xml: string, from: number): string {
    const head = xml.slice(0, from);
    const tail = xml.slice(from);
    const plain = (x: string) => x.replace(/<w:tab\s*\/>/g, "\t").replace(/<[^>]+>/g, "");
    return head + tail.replace(/<w:tab\s*\/>|<w:br\b(?![^>]*w:type="(?:page|column)")[^>]*\/>|<w:cr\s*\/>/g, (m, off: number) => {
        const before = plain(tail.slice(0, off));
        const after = plain(tail.slice(off + m.length));
        if (!before.trim() || !after.trim() || /\s$/.test(before) || /^\s/.test(after)) return "";
        return `<w:t xml:space="preserve"> </w:t>`;
    });
}

export function alignLabelLines(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    opts: { halfPts: number },
): { lines: number; regions: VerifyRegion[] } {
    const out = { lines: 0, regions: [] as VerifyRegion[] };
    const charTw = Math.round(opts.halfPts * 5.6);

    const ok = (n: ParaNode | TableNode, i: number): boolean => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || !it || it.role !== "body") return false;
        if (it.zone !== "front" && it.zone !== "back") return false;
        if (n.hasNumbering || n.hasDrawing || n.hasField) return false;
        const t = n.text.replace(/\s+/g, " ").trim();
        if (!t || t.length > 250) return false;
        // Cukup titik dua PERTAMA yang jadi pemisah label; isi boleh memuat titik dua
        // lagi ("Judul Skripsi : … Cairan Rumen: In vitro").
        const raw = leadingRaw(n.xml, 80);
        const m = LV_RAW_RE.exec(raw);
        if (!m) return false;
        // "Pukul 10:30" bukan label.
        return !(/\d$/.test(m[1].trim()) && /^\d/.test(raw.slice(m[0].length)));
    };

    let group: ParaNode[] = [];
    const flush = () => {
        if (group.length >= 2) {
            const labels = group.map((p) => LV_RAW_RE.exec(leadingRaw(p.xml, 80))![1].trim());
            const longest = Math.max(...labels.map((l) => l.length));
            const labelW = Math.min(3400, Math.max(1134, longest * charTw + 170));
            // Titik awal baris pertama = left − hanging + firstLine (bukan w:left mentah:
            // "left=2160 hanging=2160" artinya label mulai di 0). Satu nilai untuk
            // sekelompok baris supaya labelnya mulai di garis yang sama.
            const startOf = (p: ParaNode) => {
                const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
                const ind = /<w:ind\b[^>]*>/.exec(pPr)?.[0] ?? "";
                const v = (a: string) => Number(new RegExp(`w:${a}="(-?\\d+)"`).exec(ind)?.[1] ?? 0);
                return Math.max(0, (v("left") || v("start")) - v("hanging") + v("firstLine"));
            };
            const base = Math.min(...group.map(startOf));
            group.forEach((p, k) => {
                const before = paragraphLine(p.xml);
                const m = LV_RAW_RE.exec(leadingRaw(p.xml, 80))!;
                const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p.xml)?.[1] ?? "";
                let xml = removeLeading(p.xml, m[0].length);
                const e = xml.indexOf("</w:pPr>");
                const at = e !== -1 ? e + 8 : xml.indexOf(">") + 1;
                const labelRun = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(labels[k])}</w:t><w:tab/><w:t xml:space="preserve">: </w:t></w:r>`;
                xml = flattenValue(xml.slice(0, at) + labelRun + xml.slice(at), at + labelRun.length);
                const hang = labelW + charTw * 2;
                p.xml = withProps(xml, "w:pPr", {
                    "w:tabs": `<w:tabs><w:tab w:val="left" w:pos="${base + labelW}"/></w:tabs>`,
                    "w:ind": `<w:ind w:left="${base + hang}" w:right="0" w:hanging="${hang}"/>`,
                    "w:jc": `<w:jc w:val="left"/>`,
                }, PPR_ORDER);
                const after = paragraphLine(p.xml);
                if (after !== before) out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: "", afterText: "" });
                out.lines++;
            });
        }
        group = [];
    };

    nodes.forEach((n, i) => {
        if (ok(n, i)) { group.push(n as ParaNode); return; }
        // Baris kosong tidak memutus kelompok; baris lain memutus.
        if (n.kind === "paragraph" && !n.text.trim() && !n.hasSectPr && !n.hasPageBreak) return;
        flush();
    });
    flush();
    return out;
}
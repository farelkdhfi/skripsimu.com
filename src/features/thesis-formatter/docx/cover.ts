/**
 * Tata letak halaman sampul (cadangan bila template General tidak bisa dipakai),
 * disusun ke urutan baku:
 *
 *   Judul topik                                     ← atas
 *   ── ruang ──
 *   SKRIPSI / USULAN PENELITIAN                     ┐
 *   Diajukan sebagai salah satu syarat…             │ blok tengah
 *   Oleh: / Nama / NIM / Pembimbing                 │
 *   [logo]                                          ┘
 *   ── ruang ──
 *   Prodi / Fakultas / Universitas / Kota / Tahun   ← frame di margin bawah
 *
 * Paragraf sampul DIPINDAH ke urutan ini; teksnya tidak diubah.
 */
import { PPR_ORDER, RPR_ORDER, spacing, withProps } from "./xml";
import type { DocModel, ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import type { CoverStyle } from "../rules/format-rules";
import { paragraphLine } from "./verify";

export type CoverMetrics = { height: number; width: number; style: CoverStyle };
export type CoverReport = { pages: number; reordered: boolean; coverLines: number };

const INSTITUTION_RE = /(program\s+studi|prodi|jurusan|departemen|fakultas|universitas|institut|sekolah\s+tinggi|politeknik|akademi|kementerian)/i;
const YEAR_RE = /^\d{4}$/;
const NOT_INSTITUTION_RE = /(oleh|nim\b|npm\b|nrp\b|skripsi|tesis|disusun|diajukan|pembimbing|:)/i;
const DOC_TYPE_RE = /^\(?((usulan|proposal|penelitian|skripsi|tesis|thesis|disertasi|tugas|akhir|seminar|hasil|laporan)[\s\/\-]*)+\)?$/i;
const PURPOSE_RE = /(diajukan\s+(?!oleh)|syarat|guna\s+memperoleh|untuk\s+memperoleh|memperoleh\s+gelar|untuk\s+memenuhi|disusun\s+untuk)/i;
const PURPOSE_CONT_RE = /^(memperoleh|gelar|sarjana|magister|doktor|pada\s+program|sebagai|\()/i;
const OLEH_RE = /^((disusun|ditulis|dibuat|diajukan)\s+)?oleh\s*:?$/i;
const ID_RE = /^(nim|npm|nrp|no\.?\s*bp|stambuk|nomor\s+induk)\b/i;
const PBB_RE = /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/;
const EMU_PER_TWIP = 635;

type Group = "doc" | "title" | "purpose" | "logo" | "author" | "inst";

function attrs(xml: string, tag: string): string {
    const end = xml.indexOf("</w:pPr>");
    if (end === -1) return "";
    return new RegExp(`<${tag}(\\s[^>]*)/?>`).exec(xml.slice(0, end))?.[1] ?? "";
}
const num = (a: string, n: string) => {
    const m = new RegExp(`${n}="(-?\\d+)"`).exec(a);
    return m ? Number(m[1]) : undefined;
};

function maxRunSize(xml: string, dflt: number): number {
    const sizes = [...xml.matchAll(/<w:r[\s>][\s\S]*?<\/w:r>/g)]
        .map((m) => /<w:sz\s+w:val="(\d+)"/.exec(m[0])?.[1])
        .filter((v): v is string => Boolean(v))
        .map(Number);
    return sizes.length ? Math.max(...sizes) : dflt;
}

function setSize(p: ParaNode, half: number) {
    p.xml = p.xml
        .replace(/<w:sz\s+w:val="\d+"\s*\/>/g, `<w:sz w:val="${half}"/>`)
        .replace(/<w:szCs\s+w:val="\d+"\s*\/>/g, `<w:szCs w:val="${half}"/>`);
}

/** Tampilkan kapital lewat format (w:caps); teks asli tidak diubah. */
function setCaps(p: ParaNode) {
    p.xml = p.xml.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) =>
        withProps(run, "w:rPr", { "w:caps": "<w:caps/>" }, RPR_ORDER));
}

function setPara(p: ParaNode, before: number, after: number, frame?: string) {
    p.xml = withProps(p.xml, "w:pPr", {
        "w:spacing": spacing({ before: Math.max(0, Math.round(before)), after: Math.max(0, Math.round(after)), line: 240, rule: "auto" }),
        "w:pageBreakBefore": null,
        "w:framePr": frame ?? null,
    }, PPR_ORDER);
}

function shrinkLogo(p: ParaNode, maxTwips: number) {
    const max = maxTwips * EMU_PER_TWIP;
    p.xml = p.xml.replace(/<wp:inline\b[\s\S]*?<\/wp:inline>/g, (block) => {
        const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/.exec(block);
        if (!ext) return block;
        const cx = Number(ext[1]);
        const cy = Number(ext[2]);
        const k = max / Math.max(cx, cy);
        if (k >= 1) return block;
        const ncx = Math.floor(cx * k);
        const ncy = Math.floor(cy * k);
        return block
            .replace(ext[0], `<wp:extent cx="${ncx}" cy="${ncy}"/>`)
            .replace(new RegExp(`<a:ext\\s+cx="${cx}"\\s+cy="${cy}"\\s*/>`, "g"), `<a:ext cx="${ncx}" cy="${ncy}"/>`);
    });
}

function heightOf(p: ParaNode, m: CoverMetrics): number {
    const sp = attrs(p.xml, "w:spacing");
    const before = num(sp, "w:before") ?? 0;
    const after = num(sp, "w:after") ?? 0;
    if (p.hasDrawing) {
        const cy = [...p.xml.matchAll(/<wp:extent\s+cx="\d+"\s+cy="(\d+)"/g)].reduce((a, x) => a + Number(x[1]) / EMU_PER_TWIP, 0);
        return cy + before + after + 60;
    }
    const half = maxRunSize(p.xml, m.style.textHalf);
    const lineH = half * 10 * 1.17;
    const t = p.text.replace(/\t/g, "    ");
    const upper = /[A-Z]/.test(t) && t === t.toUpperCase();
    const cpl = Math.max(10, m.width / (half * (upper ? 6.5 : 5.5)));
    return Math.max(1, Math.ceil(t.length / cpl)) * lineH + before + after;
}

export function layoutCover(
    model: DocModel,
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    m: CoverMetrics,
): CoverReport {
    const s = m.style;
    const rep: CoverReport = { pages: 0, reordered: false, coverLines: 0 };

    // Halaman fisik di zona sampul + jumlah baris teksnya (untuk verifikasi).
    const pages: ParaNode[][] = [];
    let cur: ParaNode[] = [];
    let hasTable = false;
    const flush = () => { if (cur.length && !hasTable) pages.push(cur); cur = []; hasTable = false; };
    for (let i = 0; i < nodes.length; i++) {
        if (outline.items[i]?.zone !== "cover") break;
        const n = nodes[i];
        if (n.kind === "table") { hasTable = true; continue; }
        if (paragraphLine(n.xml)) rep.coverLines++;
        const pPrEnd = n.xml.indexOf("</w:pPr>");
        if (cur.length && pPrEnd !== -1 && PBB_RE.test(n.xml.slice(0, pPrEnd))) flush();
        cur.push(n);
        if (n.hasPageBreak || n.hasSectPr) flush();
    }
    flush();

    for (const page of pages) {
        if (page.some((p) => /<wp:anchor\b/.test(p.xml))) continue;

        const isBoundary = (p: ParaNode) => p.hasPageBreak || p.hasSectPr || p.hasField || /<w:bookmarkStart/.test(p.xml);
        const tails = page.filter(isBoundary);
        if (tails.some((p) => p.text.trim() || p.hasDrawing)) continue;
        const content = page.filter((p) => !isBoundary(p) && (p.text.trim() || p.hasDrawing));
        if (!content.length) continue;

        // Blok institusi: dari baris paling bawah ke atas.
        const inst = new Set<ParaNode>();
        let steps = 0;
        let seenKeyword = false;
        for (let k = content.length - 1; k >= 0; k--) {
            const p = content[k];
            const t = p.text.trim();
            if (p.hasDrawing || NOT_INSTITUTION_RE.test(t) || t.length > 80) break;
            const kw = INSTITUTION_RE.test(t);
            if (!(YEAR_RE.test(t) || kw || (!seenKeyword && steps < 3 && t.length <= 40 && !/\d/.test(t)))) break;
            if (kw) seenKeyword = true;
            inst.add(p);
            steps++;
        }
        if (!seenKeyword) inst.clear();

        // Kelompokkan sisanya, urutan relatif di dalam kelompok tetap.
        const groups: Record<Group, ParaNode[]> = { doc: [], title: [], purpose: [], logo: [], author: [], inst: [] };
        let inAuthor = false;
        let last: Group | null = null;
        for (const p of content) {
            const t = p.text.trim();
            let g: Group;
            if (p.hasDrawing) g = "logo";
            else if (inst.has(p)) g = "inst";
            else if (OLEH_RE.test(t)) { g = "author"; inAuthor = true; }
            else if (ID_RE.test(t) || (inAuthor && (t.includes(":") || (t.length <= 60 && !PURPOSE_RE.test(t) && !DOC_TYPE_RE.test(t))))) g = "author";
            else if (DOC_TYPE_RE.test(t)) { g = "doc"; inAuthor = false; }
            else if (PURPOSE_RE.test(t) || (last === "purpose" && PURPOSE_CONT_RE.test(t))) { g = "purpose"; inAuthor = false; }
            else { g = "title"; inAuthor = false; }
            groups[g].push(p);
            last = g;
        }
        // Baris nama tanpa "Oleh:" — baris pendek tepat sebelum NIM ikut kelompok penulis.
        for (const id of groups.author.filter((p) => ID_RE.test(p.text.trim()))) {
            const k = content.indexOf(id) - 1;
            const prev = content[k];
            if (prev && groups.title.includes(prev) && prev.text.trim().length <= 60) {
                groups.title.splice(groups.title.indexOf(prev), 1);
                groups.author.unshift(prev);
            }
        }

        const ordered = [...groups.title, ...groups.doc, ...groups.purpose, ...groups.author, ...groups.logo, ...groups.inst];

        // Susun ulang di body (hanya bila halaman ini berurutan di tingkat atas).
        const idxs = page.map((p) => model.body.indexOf(p));
        if (idxs.some((i) => i === -1)) continue;
        const lo = Math.min(...idxs);
        const hi = Math.max(...idxs);
        if (hi - lo + 1 !== page.length) continue;
        model.body.splice(lo, page.length, ...ordered, ...tails);
        if (ordered.some((p, k) => p !== content[k])) rep.reordered = true;

        // Ukuran font.
        for (const p of groups.title) setSize(p, s.titleHalf);
        for (const p of [...groups.doc, ...groups.purpose, ...groups.author]) setSize(p, s.textHalf);
        for (const p of groups.inst) {
            setSize(p, s.institutionHalf);
            if (s.institutionCaps) setCaps(p);
        }

        // Jenis dokumen selalu 12 pt.
        for (const p of groups.doc) setSize(p, 24);

        // Jarak dasar: rapat di dalam kelompok.
        const gap = s.textHalf * 10;
        for (const p of ordered) if (!inst.has(p)) setPara(p, 0, 0);
        groups.logo.forEach((p) => shrinkLogo(p, Math.round(s.logoMaxTwips * 1.25)));

        // Institusi → frame di margin bawah.
        const frame = `<w:framePr w:w="${Math.round(m.width)}" w:wrap="notBeside" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="center" w:yAlign="bottom"/>`;
        groups.inst.forEach((p) => setPara(p, 0, 0, frame));

        // Blok tengah di antara judul dan institusi; ruang di atas & di bawah logo seimbang.
        const middle = [...groups.doc, ...groups.purpose, ...groups.author, ...groups.logo];
        if (middle.length) {
            const h = (list: ParaNode[]) => list.reduce((sum, p) => sum + heightOf(p, m), 0);
            const A = h(groups.title);
            const G = h(middle) + (groups.doc.length && groups.purpose.length ? gap : 0);
            const D = h(groups.inst) + h(tails);
            const hasLogo = groups.logo.length > 0;
            const weight = {
                top: groups.title.length ? 1 : 0,
                author: groups.author.length && middle[0] !== groups.author[0] ? 0.7 : 0,
                logo: hasLogo && middle[0] !== groups.logo[0] ? 1 : 0,
                bottom: 1,
            };
            const minGap = {
                top: weight.top ? 360 : 0,
                author: weight.author ? 240 : 0,
                logo: weight.logo ? 360 : 0,
                bottom: 480,
            };
            const minSum = minGap.top + minGap.author + minGap.logo + minGap.bottom;
            const wSum = weight.top + weight.author + weight.logo + weight.bottom;
            const room = Math.max(0, m.height - A - G - D - 60);
            const k = room < minSum ? room / minSum : 1;
            const extra = Math.max(0, room - minSum);
            const gapOf = (key: keyof typeof weight) => minGap[key] * k + (extra * weight[key]) / wSum;
            setPara(middle[0], gapOf("top"), 0);
            if (groups.doc.length && groups.purpose.length) setPara(groups.purpose[0], gap, 0);
            if (weight.author) setPara(groups.author[0], gapOf("author"), 0);
            if (weight.logo) setPara(groups.logo[0], gapOf("logo"), 0);
        }

        rep.pages++;
    }

    return rep;
}
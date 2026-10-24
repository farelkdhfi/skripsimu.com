/**
 * Membangun outline dokumen: peran + zona untuk setiap node, mengikuti
 * dokumen yang diunggah. Tidak ada template generik di sini — urutan bagian
 * sepenuhnya berasal dari file asli.
 */

import type { ParaNode, TableNode } from "../docx/document-model";
import {
    LIST_MARKER_RE, captionKind, classifyByRule, frontKindOf, isHeadingRole, scanStructure,
    type FrontKind, type LayoutHint, type Role, type Signals, type Zone,
} from "./roles";

export type OutlineItem = {
    role: Role;
    zone: Zone;
    /** Petunjuk tata letak asli, diteruskan ke lapisan aturan. */
    hint: LayoutHint;
    /** Node ini "BAB I" dan judul babnya ada di node berikutnya. */
    pairedTitleNext?: boolean;
    /** Node ini judul bab yang mengikuti label "BAB I" di node sebelumnya. */
    continuesChapterLabel?: boolean;
    /** Halaman baru sudah dimulai oleh node sebelumnya — jangan tambah break lagi. */
    breakAlreadyPresent?: boolean;
    /** Judul ini sengaja berbagi halaman dengan judul sebelumnya (motto + persembahan). */
    sharesPage?: boolean;
};

export type Outline = {
    items: OutlineItem[];
    zones: Zone[];
    /** Indeks node pertama pada zona "main" (awal BAB I). */
    mainStart: number;
};

export type OutlineInput = ParaNode | TableNode;

function signalsOf(n: OutlineInput, chapterIds: Set<number>): Signals {
    return n.kind === "table"
        ? { text: "", hasDrawing: false, isTable: true }
        : {
            text: n.text, hasDrawing: n.hasDrawing, isTable: false,
            captionStyled: /caption|keterangan/i.test(n.styleId ?? ""),
            chapterHint: chapterIds.has(n.id),
        };
}

/**
 * @param aiRoles peran hasil AI per indeks node; hanya dipakai untuk node yang
 *   tidak bisa diputuskan oleh aturan, sehingga AI tidak pernah menimpa
 *   penanda struktural yang sudah pasti.
 */

const AI_DEGREE_RE = /,\s*(?:S|M|Dr|Ir|Drs|Dra|Prof|Ph)\.?\s?[A-Za-z]|\b(?:Dr|Prof|Ir|Drs|Dra|Hj?)\.\s/i;
const AI_DATE_RE = /\b(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)\s+\d{4}\b/i;
const AI_ID_RE = /^(nip|nidn|nik|nim|npm|nuptk)\b/i;

/**
 * Peran judul dari AI hanya diterima bila masuk akal: baris bergelar, tanggal,
 * NIP/NIDN, baris berakhiran koma/titik dua, atau "(Nama)" bukan judul.
 * Di luar isi (sampul/awal/lampiran), AI tidak boleh membuat bab/subbab.
 */
function plausibleAiHeading(text: string, role: Role, zone: Zone): boolean {
    const t = text.replace(/\s+/g, " ").trim();
    if (!t || t.length > 120) return false;
    if (AI_DEGREE_RE.test(t) || AI_DATE_RE.test(t) || AI_ID_RE.test(t)) return false;
    if (/[,:;]$/.test(t) || /^\(.*\)$/.test(t)) return false;
    if (zone !== "main" && (role === "chapter" || role === "chapter_label" || role === "section" || role === "subsection")) return false;
    if ((zone === "front" || zone === "cover") && (t.includes(",") || t.split(" ").length > 8)) return false;
    return true;
}

export function buildOutline(nodes: OutlineInput[], aiRoles: (Role | null)[] = [], chapterIds: Set<number> = new Set()): Outline {
    const signals = nodes.map((n) => signalsOf(n, chapterIds));
    const { zones, inList } = scanStructure(signals);

    const roles: Role[] = signals.map((s, i) => {
        if (s.isTable) return "body";
        const rule = classifyByRule(s, zones[i], inList[i]);
        if (rule !== null && rule !== "empty") return rule;
        if (rule === "empty") return "empty";
        const ai = aiRoles[i];
        if (ai) {
            const fallback: Role = zones[i] === "cover" ? "cover" : "body";
            if (isHeadingRole(ai) && !plausibleAiHeading(s.text, ai, zones[i])) return fallback;
            // Peran sampul hanya sah di zona sampul (di luar itu membuat baris jadi rata tengah).
            if (ai === "cover" && zones[i] !== "cover") return "body";
            // Entri daftar isi hanya sah di dalam blok DAFTAR ISI/TABEL/GAMBAR.
            if (ai === "list_entry" && !inList[i]) return fallback;
            return ai;
        }
        return zones[i] === "cover" ? "cover" : "body";
    });

    // Keterangan panjang (judul + penjelasan + "Sumber: …", > 200 karakter) dikenali dari
    // letaknya: tepat di bawah/atas gambar, atau tepat di atas/bawah tabel.
    {
        const blankAt = (k: number) => { const n = nodes[k]; return n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing; };
        const near = (i: number, step: 1 | -1) => {
            for (let k = i + step; k >= 0 && k < nodes.length; k += step) if (!blankAt(k)) return k;
            return -1;
        };
        const isFig = (k: number) => k !== -1 && nodes[k].kind === "paragraph" && (nodes[k] as ParaNode).hasDrawing;
        const isTbl = (k: number) => k !== -1 && nodes[k].kind === "table";
        roles.forEach((role, i) => {
            const n = nodes[i];
            if (role !== "body" || n.kind !== "paragraph" || n.hasDrawing) return;
            if (zones[i] !== "main" && zones[i] !== "back") return;
            const t = n.text.trim();
            if (t.length <= 200 || t.length > 700) return;
            const kind = captionKind(t);
            if (kind === "figure_caption" && (isFig(near(i, -1)) || isFig(near(i, 1)))) roles[i] = kind;
            else if (kind === "table_caption" && (isTbl(near(i, -1)) || isTbl(near(i, 1)))) roles[i] = kind;
        });
    }

    // Daftar pustaka: paragraf setelah judulnya adalah entri referensi,
    // sampai bertemu judul lain (mis. LAMPIRAN).
    let inRefs = false;
    for (let i = 0; i < roles.length; i++) {
        if (roles[i] === "reference_heading") { inRefs = true; continue; }
        if (!inRefs) continue;
        if (roles[i] === "chapter" || roles[i] === "chapter_label" || roles[i] === "appendix_heading" ||
            roles[i] === "front_heading" || roles[i] === "list_heading") inRefs = false;
        else if (roles[i] === "body") {
            // Catatan/keterangan di bawah judul Daftar Pustaka bukan entri pustaka
            // (tidak diberi indentasi gantung).
            const n = nodes[i];
            const t = n.kind === "paragraph" ? n.text.trim() : "";
            if (t.length >= 25 && !/^(catatan|keterangan|note|sumber)\b/i.test(t)) roles[i] = "reference";
        }
    }

    // Jenis halaman awal diwarisi paragraf di bawah judulnya, sampai judul
    // berikutnya. Di luar zona depan nilainya null.
    let kind: FrontKind | null = null;
    const items: OutlineItem[] = roles.map((role, i) => {
        const n = nodes[i];
        if (role === "front_heading" && n.kind === "paragraph") kind = frontKindOf(n.text);
        else if (role === "list_heading" || zones[i] !== "front") kind = null;
        return {
            role,
            zone: zones[i],
            hint: n.kind === "paragraph"
                ? {
                    align: n.align, tabCount: n.tabCount,
                    leadingTabsOnly: n.tabCount > 0 && /^\t/.test(n.text) && !/\t/.test(n.text.replace(/^\t+/, "")),
                    listMarker: !n.hasNumbering && LIST_MARKER_RE.test(n.text.replace(/^\t+/, "")),
                    allBold: n.allBold,
                    length: n.text.length, frontKind: zones[i] === "front" ? kind : null,
                }
                : { align: null, tabCount: 0, leadingTabsOnly: false, listMarker: false, allBold: false, length: 0, frontKind: null },
        };
    });

    // Pasangkan "BAB I" dengan judul bab di baris berikutnya, tanpa menggabungkan
    // teksnya — keduanya tetap paragraf terpisah seperti di dokumen asli.
    for (let i = 0; i < items.length; i++) {
        if (items[i].role !== "chapter_label") continue;
        let j = i + 1;
        while (j < items.length && items[j].role === "empty") j++;
        const next = nodes[j];
        if (j < items.length && next?.kind === "paragraph" && next.text.trim() && next.text.trim().length <= 90 &&
            (items[j].role === "body" || items[j].role === "cover" || items[j].role === "chapter")) {
            items[i].pairedTitleNext = true;
            items[j].role = "chapter";
            items[j].continuesChapterLabel = true;
            items[i].hint.pairedChapter = true;
            items[j].hint.pairedChapter = true;
        }
    }

    // Motto dan persembahan lazim berbagi satu halaman ("MOTTO DAN PERSEMBAHAN"
    // di daftar isi). Bila dokumen asli menaruhnya berdampingan tanpa pindah
    // halaman, jangan dipisahkan.
    const pagedHeadings = (i: number) => {
        const n = nodes[i];
        return n.kind === "paragraph" && items[i].role === "front_heading" ? n.text.trim().toLowerCase() : null;
    };
    let lastFront: { idx: number; text: string } | null = null;
    for (let i = 0; i < items.length; i++) {
        const t = pagedHeadings(i);
        if (!t) continue;
        const n = nodes[i] as ParaNode;
        if (lastFront && /^(motto|moto)\b/.test(lastFront.text) && /^persembahan\b/.test(t) &&
            !n.hadPageBreakBefore && !nodes.slice(lastFront.idx + 1, i).some((x) => x.kind === "paragraph" && (x.hasPageBreak || x.hasSectPr))) {
            items[i].sharesPage = true;
        }
        lastFront = { idx: i, text: t };
    }

    // Deteksi page break yang SUDAH ada, agar kita tidak menambah break kedua
    // (penyebab klasik halaman kosong).
    let seenText = false;
    for (let i = 0; i < items.length; i++) {
        const self = nodes[i];
        const hadText = seenText;
        if (self.kind === "table" || (self.kind === "paragraph" && self.text.trim())) seenText = true;

        // Tidak ada isi apa pun sebelum node ini, jadi memulai halaman baru di sini
        // hanya akan menghasilkan halaman kosong di depan dokumen.
        if (!hadText) { items[i].breakAlreadyPresent = true; continue; }

        let k = i - 1;
        while (k >= 0 && nodes[k].kind === "paragraph" && !(nodes[k] as ParaNode).text.trim() &&
            !(nodes[k] as ParaNode).hasDrawing && !(nodes[k] as ParaNode).hasPageBreak &&
            !(nodes[k] as ParaNode).hasSectPr) k--;
        if (k < 0) { items[i].breakAlreadyPresent = true; continue; }
        const prev = nodes[k];
        if (prev.kind === "paragraph" && (prev.hasPageBreak || prev.hasSectPr)) items[i].breakAlreadyPresent = true;
        if (items[k]?.pairedTitleNext && items[i].continuesChapterLabel) items[i].breakAlreadyPresent = true;
        if (items[i].sharesPage) items[i].breakAlreadyPresent = true;
    }

    const mainStart = zones.findIndex((z) => z === "main");
    return { items, zones, mainStart: mainStart === -1 ? nodes.length : mainStart };
}

export function outlineStats(nodes: OutlineInput[], outline: Outline) {
    const count = (r: Role) => outline.items.filter((it) => it.role === r).length;
    const chapters = outline.items.filter(
        (it) => it.role === "chapter" && !it.continuesChapterLabel,
    ).length + outline.items.filter((it) => it.role === "chapter_label" && it.pairedTitleNext).length;
    return {
        chapters,
        sections: count("section"),
        subsections: count("subsection"),
        figures: nodes.filter((n) => n.kind === "paragraph" && n.hasDrawing).length,
        figureCaptions: count("figure_caption"),
        tableCaptions: count("table_caption"),
        tables: nodes.filter((n) => n.kind === "table").length,
        references: count("reference"),
        frontMatter: count("front_heading") + count("list_heading"),
        appendices: count("appendix_heading"),
        referenceHeadings: count("reference_heading"),
    };
}
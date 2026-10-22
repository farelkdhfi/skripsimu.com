/**
 * Membangun outline dokumen: peran + zona untuk setiap node, mengikuti
 * dokumen yang diunggah. Tidak ada template generik di sini — urutan bagian
 * sepenuhnya berasal dari file asli.
 */

import type { ParaNode, TableNode } from "../docx/document-model";
import {
    LIST_MARKER_RE, classifyByRule, frontKindOf, scanStructure,
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

function signalsOf(n: OutlineInput): Signals {
    return n.kind === "table"
        ? { text: "", hasDrawing: false, isTable: true }
        : { text: n.text, hasDrawing: n.hasDrawing, isTable: false };
}

/**
 * @param aiRoles peran hasil AI per indeks node; hanya dipakai untuk node yang
 *   tidak bisa diputuskan oleh aturan, sehingga AI tidak pernah menimpa
 *   penanda struktural yang sudah pasti.
 */
export function buildOutline(nodes: OutlineInput[], aiRoles: (Role | null)[] = []): Outline {
    const signals = nodes.map(signalsOf);
    const { zones, inList } = scanStructure(signals);

    const roles: Role[] = signals.map((s, i) => {
        if (s.isTable) return "body";
        const rule = classifyByRule(s, zones[i], inList[i]);
        if (rule !== null && rule !== "empty") return rule;
        if (rule === "empty") return "empty";
        const ai = aiRoles[i];
        if (ai) return ai;
        return zones[i] === "cover" ? "cover" : "body";
    });

    // Daftar pustaka: paragraf setelah judulnya adalah entri referensi,
    // sampai bertemu judul lain (mis. LAMPIRAN).
    let inRefs = false;
    for (let i = 0; i < roles.length; i++) {
        if (roles[i] === "reference_heading") { inRefs = true; continue; }
        if (!inRefs) continue;
        if (roles[i] === "chapter" || roles[i] === "chapter_label" || roles[i] === "appendix_heading" ||
            roles[i] === "front_heading" || roles[i] === "list_heading") inRefs = false;
        else if (roles[i] === "body") roles[i] = "reference";
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
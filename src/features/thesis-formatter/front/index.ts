/**
 * Bagian depan (sampul s.d. abstract) dari TEMPLATE DOCX preset.
 *
 * Alur:
 *  1. Wilayah bagian depan naskah = awal dokumen sampai BAB I. Blok DAFTAR ISI/TABEL/ILUSTRASI/
 *     LAMPIRAN di dalamnya dipertahankan (disusun ulang modul daftar) dan diletakkan sesudah
 *     bagian depan; sisanya dibaca sebagai bahan.
 *  2. Data dibaca dengan aturan, dan — bila tersedia — dengan AI yang hanya menunjuk baris.
 *     Tunjukan AI yang lolos pemeriksaan dipakai; selebihnya hasil aturan. Isian pengguna
 *     (form konfirmasi) menimpa keduanya.
 *  3. Template diisi (lihat template.ts) dan menggantikan seluruh bagian depan naskah. Template
 *     SELALU dipakai: data yang tidak ditemukan diisi titik-titik dan dilaporkan, tidak dibatalkan.
 *
 * Integritas: data dari naskah harus ada di teks naskah asli; baris & gambar naskah yang tidak
 * dipakai dilaporkan.
 */
import JSZip from "jszip";
import { unifyChapterHeadings } from "../docx/chapter-heads";
import { flatten, parseBlocks, parseDocument, serializeBlock, serializeDocument, type BlockNode, type DocModel, type ParaNode, type RawNode } from "../docx/document-model";
import { fingerprint, type VerifyIssue, type VerifyRegion } from "../docx/verify";
import { buildOutline } from "../structure/outline";
import type { FormatConfig } from "../schemas/format-config";
import { presetById } from "../presets";
import { readBlocks, richText } from "./read";
import { isChapterStart, segmentPages } from "./segment";
import { splitMergedLines } from "./split";
import { collapseRepeat, emptyFields as emptyFieldsOf, extractFront, plainRich, type Extraction } from "./extract";
import { applyPointers, buildListing, fieldOk, type AiData, type FrontAI } from "./ai";
import { fillTemplate, loadTemplate, missingNamespaces, TemplateError } from "./template";
import {
    FIELD_KEYS, FIELD_LABEL, FRONT_PAGE_LABEL,
    type FrontFieldInput, type FrontFieldKey, type FrontFields, type FrontPageKind, type FrontPreviewLine,
    type Item, type LineItem, type Pic, type Rich,
} from "./types";

export type { FrontAI } from "./ai";

/** Penanda versi kode bagian depan: tercetak di log & rekaman diagnosis, untuk memastikan kode yang jalan. */
export const FRONT_VERSION = "2026-10-29.1";

/** Asal data: dibaca dari naskah (aturan / AI) atau diisi pengguna. */
/** Asal data: aturan / AI (dari naskah), pengguna (form), atau bawaan (diisi otomatis, mis. tempat & tanggal). */
export type FieldSource = "aturan" | "ai" | "pengguna" | "bawaan" | null;

/** Ringkasan untuk UI (form konfirmasi) dan laporan. */
export type FrontReport = {
    built: boolean;
    /** Alasan template tidak dipakai (null bila dipakai, atau preset tanpa template). */
    reason: string | null;
    /** Halaman yang dikenali di naskah, sesuai urutan naskah. */
    found: FrontPageKind[];
    /** Halaman template yang isinya tidak ditemukan di naskah. */
    missing: FrontPageKind[];
    fields: Record<FrontFieldKey, string | null>;
    sources: Record<FrontFieldKey, FieldSource>;
    /** Baris naskah yang tidak dipakai template. */
    unmapped: { page: string; text: string }[];
};

export type FrontResult = FrontReport & {
    notes: string[];
    issues: VerifyIssue[];
    /** Kata pengantar s.d. abstract sudah memulai nomor romawi (section sesudahnya melanjutkan). */
    romanStarted: boolean;
    /**
     * Blok teks yang sengaja diganti, untuk verifikasi: bagian depan (naskah → template) dan
     * tiap judul halaman daftar SENDIRI-SENDIRI — judul "DAFTAR TABEL" juga muncul sebagai baris
     * Daftar Isi, jadi bila digabung ke blok bagian depan, baris Daftar Isi yang ikut terambil.
     */
    regions: VerifyRegion[];
    delta: { drawings: number; tables: number; rows: number; cells: number };
    removedImageIds: string[];
    tocRows: { text: string; bookmark: string }[];
    preview: FrontPreviewLine[];
};

const simple = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

function asText(f: FrontFields): Record<FrontFieldKey, string | null> {
    const out = {} as Record<FrontFieldKey, string | null>;
    for (const k of FIELD_KEYS) {
        const v = f[k];
        out[k] = v === null ? null : typeof v === "string" ? v : richText(v).replace(/\s+/g, " ").trim();
    }
    return out;
}

const nulls = <T,>(v: T) => Object.fromEntries(FIELD_KEYS.map((k) => [k, v])) as Record<FrontFieldKey, T>;

const empty = (reason: string | null): FrontResult => ({
    built: false, reason, found: [], missing: [], fields: nulls(null), sources: nulls(null),
    unmapped: [], notes: [], issues: [], romanStarted: false, regions: [],
    delta: { drawings: 0, tables: 0, rows: 0, cells: 0 }, removedImageIds: [], tocRows: [], preview: [],
});

/* ------------------------------------------------------------------ */
/* Wilayah bagian depan                                                 */
/* ------------------------------------------------------------------ */

type Region = { end: number; keep: number[] };

const ROMAN: Record<string, number> = { i: 1, v: 5, x: 10, l: 50 };
const romanValue = (r: string) => [...r.toLowerCase()].reduce((acc, c, i, a) =>
    (ROMAN[a[i + 1]] ?? 0) > ROMAN[c] ? acc - ROMAN[c] : acc + ROMAN[c], 0);
/** Nomor bab di awal baris ("BAB II", "II KAJIAN …", "2. …"); null bila tidak bernomor ("PENDAHULUAN"). */
function chapterNo(text: string): number | null {
    const m = /^\s*(?:(?:bab|chapter)\s+)?([ivxl]+|\d{1,2})(?=[\s.)]|$)/i.exec(text);
    if (!m) return null;
    return /\d/.test(m[1]) ? Number(m[1]) : romanValue(m[1]);
}
/** Paragraf isi (kalimat utuh), bukan baris daftar/judul/keterangan. */
function isProse(text: string): boolean {
    const t = text.replace(/\s+/g, " ").trim();
    if (t.split(" ").length < 5) return false;
    if (/^(tabel|table|gambar|ilustrasi|grafik|figure|lampiran)\s*\.?\s*\d/i.test(t)) return false;   // keterangan / entri daftar
    if (/(\.{2,}|\u2026|\s)\s*[\divxlcdm]+\s*$/i.test(text.trimEnd()) && /(\.{2,}|\u2026|\t)/.test(text)) return false; // entri daftar
    return /[.!?]["\u201d)]?$/.test(t) || /[a-z]\.\s+[A-Z]/.test(t);
}

/**
 * Batas akhir bagian depan = awal BAB I yang ASLI, plus blok daftar (DAFTAR ISI dkk.) sebelumnya.
 *
 * BAB I asli: baris awal bab bernomor I/1 (atau tanpa nomor, mis. "PENDAHULUAN") yang DIIKUTI ISI
 * (paragraf kalimat) sebelum baris bab lain. Baris bab di Daftar Isi — dengan atau tanpa nomor
 * halaman, field Word atau teks biasa — tidak pernah diikuti paragraf isi sebelum baris bab
 * berikutnya, jadi tidak terpilih. Isi field TOC / content control Daftar Isi Word dilewati.
 * Tidak ada BAB I yang memenuhi → gagal (dilaporkan), TIDAK menebak.
 */
function findRegion(model: DocModel): Region | string {
    const tops: number[] = [];
    const nodes = model.body.flatMap((b, i) => {
        const ns = flatten([b]);
        ns.forEach(() => tops.push(i));
        return ns;
    });
    // Judul bab romawi tanpa kata BAB ("I" ↵ "PENDAHULUAN") dikenali modul judul bab; dijalankan
    // pada salinan supaya naskah belum berubah.
    const chapterIds = unifyChapterHeadings(parseDocument(serializeDocument(model))).chapterIds;
    const outline = buildOutline(nodes, [], chapterIds);
    const textOf = (k: number) => (nodes[k].kind === "paragraph" ? (nodes[k] as ParaNode).text : "");

    // Isi Daftar Isi Word: di dalam field TOC (sampai fldChar penutupnya) atau content control
    // "Table of Contents".
    const inToc = new Array<boolean>(nodes.length).fill(false);
    let depth = 0;
    nodes.forEach((n, k) => {
        const top = model.body[tops[k]];
        if (top.kind === "container" && /docPartGallery\s+w:val="Table of Contents"/.test(top.open)) inToc[k] = true;
        if (n.kind !== "paragraph") { if (depth > 0) inToc[k] = true; return; }
        const begins = (n.xml.match(/w:fldCharType="begin"/g) ?? []).length;
        const ends = (n.xml.match(/w:fldCharType="end"/g) ?? []).length;
        if (depth === 0 && /<w:instrText[^>]*>\s*TOC\b/.test(n.xml)) {
            inToc[k] = true;
            depth = Math.max(0, begins - ends);         // field TOC masih terbuka sesudah paragraf ini
        } else if (depth > 0) {
            inToc[k] = true;
            depth = Math.max(0, depth + begins - ends);
        }
    });

    const candidate = (k: number) => {
        if (inToc[k] || nodes[k].kind !== "paragraph") return false;
        const t = textOf(k).trim();
        if (!t) return false;
        const role = outline.items[k]?.role;
        let next: string | null = null;
        for (let j = k + 1; j < nodes.length; j++) if (textOf(j).trim()) { next = textOf(j); break; }
        return role === "chapter" || role === "chapter_label" || isChapterStart(t, next) ||
            (outline.items[k]?.zone === "main" && outline.items[k - 1]?.zone !== "main");
    };
    /** Baris bab berikutnya (bukan judul bab ini): bab bernomor ≥ II, atau awal bab lain di luar 2 baris judul. */
    const otherChapter = (k: number, from: number, titleLines: number) => {
        const t = textOf(k).trim();
        const no = chapterNo(t);
        if (no !== null && no >= 2 && (/^(bab|chapter)\b/i.test(t) || /^[ivxl]+\.?\s+[A-Z]/.test(t) || /^[ivxl]+\.?$/i.test(t))) return true;
        if (outline.items[k]?.role === "list_heading") return true;
        return k > from && titleLines >= 2 && candidate(k);
    };

    /** Baris berisi berikutnya adalah entri/judul daftar. */
    const listNext = (k: number) => {
        for (let j = k + 1; j < nodes.length; j++) {
            if (nodes[j].kind === "paragraph" && !textOf(j).trim()) continue;
            const r = outline.items[j]?.role;
            return inToc[j] || r === "list_entry" || r === "list_heading";
        }
        return false;
    };

    const keep = new Set<number>();
    let listing = false;
    for (let k = 0; k < nodes.length; k++) {
        const role = outline.items[k]?.role;
        if (inToc[k]) { keep.add(tops[k]); continue; }
        if (role === "list_heading") { listing = true; keep.add(tops[k]); continue; }
        if (listing && (role === "list_entry" || role === "empty" || nodes[k].kind === "table")) { keep.add(tops[k]); continue; }
        const isCand = candidate(k);
        // Di dalam daftar, baris tanpa nomor halaman ("LEMBAR PENGESAHAN", "II KAJIAN …") yang
        // masih diikuti entri daftar tetap bagian daftar. Hanya menentukan baris mana yang dibaca
        // sebagai bahan bagian depan — batas BAB I ditentukan di bawah.
        if (listing && !isCand && !isProse(textOf(k)) && listNext(k)) { keep.add(tops[k]); continue; }
        // "DAFTAR PUSTAKA"/"LAMPIRAN" sebelum BAB I hanya mungkin baris Daftar Isi (baris terakhirnya).
        if (listing && /^(daftar\s+pustaka|lampiran)\s*$/i.test(textOf(k).trim())) { keep.add(tops[k]); continue; }
        if (!isCand) { listing = false; continue; }
        const no = chapterNo(textOf(k).trim());
        // Cari isi bab sebelum bab berikutnya.
        let body = false;
        let titleLines = 0;
        let j = k + 1;
        for (; j < nodes.length; j++) {
            if (inToc[j]) break;
            const t = textOf(j);
            if (!t.trim()) continue;
            if (otherChapter(j, k, titleLines)) break;
            titleLines++;
            if (isProse(t) && outline.items[j]?.role !== "list_entry") { body = true; break; }
        }
        if (body && (no === null || no === 1)) {
            const end = tops[k];
            // end 0 = naskah langsung dimulai dari BAB I: template tetap disisipkan di depan.
            return { end, keep: [...keep].filter((i) => i < end).sort((a, b) => a - b) };
        }
        // Bukan BAB I asli: baris bab di Daftar Isi (tanpa isi) → bagian daftar.
        if (!body) {
            for (let m = k; m < j; m++) keep.add(tops[m]);
            listing = true;
            k = j - 1;
        }
    }
    return "awal BAB I tidak ditemukan (tidak ada baris BAB I yang diikuti isi bab), jadi batas bagian depan tidak bisa ditentukan.";
}

/* ------------------------------------------------------------------ */
/* Pembacaan                                                            */
/* ------------------------------------------------------------------ */

type Read = {
    region: Region;
    blocks: BlockNode[];
    items: Item[];
    x: Extraction;
    found: FrontPageKind[];
    skipped: { reason: string; items: Item[] }[];
    blankPages: number;
    /** Jenis halaman tempat setiap item (untuk memeriksa tunjukan AI). */
    pageOf: Map<Item, FrontPageKind>;
    /** Teks tiap halaman bagian depan (huruf & angka saja) untuk cek silang antarhalaman. */
    pageNorms: string[];
    /** Kata di XML bagian depan yang tidak terbaca pembaca. */
    unread: string[];
};

/** Teks footer/header per relationship id (baris berisi selain nomor halaman). */
export type PartLines = Map<string, LineItem[]>;

/**
 * Baris teks di footer setiap section. Naskah hasil konversi PDF (dan sebagian naskah ketikan)
 * menaruh bagian bawah sampul — program studi, fakultas, universitas, kota & TAHUN — di footer;
 * tanpa ini tahun tidak pernah terbaca. Nomor halaman diabaikan.
 */
export async function footerLines(zip: JSZip): Promise<PartLines> {
    const out: PartLines = new Map();
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("string");
    if (!rels) return out;
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const tag = m[0];
        if (!/\/footer"/.test(tag)) continue;
        const id = /\bId="([^"]+)"/.exec(tag)?.[1];
        const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
        if (!id || !target) continue;
        const xml = await zip.file(`word/${target.replace(/^\/?word\//, "").replace(/^\//, "")}`)?.async("string");
        if (!xml) continue;
        const inner = /<w:ftr\b[^>]*>([\s\S]*)<\/w:ftr>/.exec(xml)?.[1] ?? "";
        const items = readBlocks(parseBlocks(inner, -1), false);
        const lines = items.filter((x): x is LineItem => x.k === "line" &&
            !/^(?:halaman|page)?\s*[-–]?\s*[\divxlcdm]+\s*[-–]?$/i.test(x.text.trim()));
        // Salinan kembar (mis. kotak teks + cadangannya) cukup sekali.
        const seen = new Set<string>();
        out.set(id, lines.filter((l) => { const k = squashNorm(l.text); if (seen.has(k)) return false; seen.add(k); return true; }));
    }
    return out;
}

function footerOf(sect: string, footers: PartLines): LineItem[] {
    const ref = (type: string) => new RegExp(`<w:footerReference\\b[^>]*w:type="${type}"[^>]*r:id="([^"]+)"|<w:footerReference\\b[^>]*r:id="([^"]+)"[^>]*w:type="${type}"`).exec(sect);
    const m = (/<w:titlePg(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(sect) ? ref("first") : null) ?? ref("default");
    const id = m?.[1] ?? m?.[2];
    return id ? footers.get(id) ?? [] : [];
}

/**
 * Pemeriksa kelengkapan pembaca: kata di XML bagian depan (di luar salinan cadangan mc:Fallback dan
 * teks terhapus) yang tidak muncul di baris hasil baca. Bila ada, ada bentuk Word yang belum
 * didukung pembaca — dicatat di log & rekaman diagnosis supaya langsung ketahuan.
 */
function unreadText(blocks: BlockNode[], items: Item[]): string[] {
    const xml = blocks.map(serializeBlock).join("")
        .replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, "")
        .replace(/<w:del\b[\s\S]*?<\/w:del>/g, "")
        .replace(/<w:instrText\b[\s\S]*?<\/w:instrText>/g, "");
    const words = (t: string) => t.toLowerCase().match(/[a-z\u00C0-\u024F0-9]{3,}/g) ?? [];
    const got = new Set(items.flatMap((it) => (it.k === "line" ? words(it.text) : [])));
    // Kata bisa terpecah di beberapa run ("P" + "engaruh"): teks digabung per paragraf dulu.
    const want = xml.split(/<w:p[\s>]|<\/w:p>|<w:tab\/>|<w:br\b[^>]*\/>/).flatMap((chunk) =>
        words([...chunk.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("")));
    return [...new Set(want.filter((w) => !got.has(w)))];
}

function readNaskah(model: DocModel, footers: PartLines = new Map()): Read | string {
    const region = findRegion(model);
    if (typeof region === "string") return region;
    const keep = new Set(region.keep);
    const blocks: BlockNode[] = [];
    let items: Item[] = [];
    for (let i = 0; i < region.end; i++) {
        if (keep.has(i)) { items.push({ k: "break" }); continue; }
        blocks.push(model.body[i]);
        const tmp: Item[] = [];
        readBlocks([model.body[i]], false, tmp);
        // Footer section ini = bagian bawah halaman terakhirnya: disisipkan sebelum pindah halaman.
        const b = model.body[i];
        const sect = b.kind === "paragraph" ? /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(b.xml)?.[0] : undefined;
        const foot = sect ? footerOf(sect, footers).map((l) => ({ ...l })) : [];
        if (foot.length) {
            const at = tmp.length && tmp[tmp.length - 1].k === "break" ? tmp.length - 1 : tmp.length;
            tmp.splice(at, 0, ...foot);
        }
        items.push(...tmp);
    }
    items = splitMergedLines(items).items;
    const unread = unreadText(blocks, items);
    if (unread.length) console.warn(`[front-template] teks bagian depan yang TIDAK terbaca pembaca (${unread.length}): ${unread.slice(0, 12).map((w) => `"${w}"`).join(", ")}`);
    const seg = segmentPages(items);
    const x = extractFront(seg.pages, {
        institution: [], docType: "SKRIPSI", purpose: "Diajukan untuk memperoleh gelar Sarjana",
        pernyataan: { body: [], intro: "Yang bertanda tangan di bawah ini, saya:" },
    });
    for (const it of seg.blankShapes) if (it.k === "pic") x.dropped.push({ pic: it.pic, why: "bentuk/garis di halaman kosong" });
    const pageOf = new Map<Item, FrontPageKind>();
    for (const p of seg.pages) for (const it of p.items) pageOf.set(it, p.kind);
    const pageNorms = [...seg.pages.map((p) => p.items), ...seg.skipped.map((p) => p.items)]
        .map((its) => squashNorm(its.map((it) => (it.k === "line" ? it.text : "")).join(" ")));
    return { region, blocks, items, x, found: seg.pages.map((p) => p.kind), skipped: seg.skipped, blankPages: seg.blankPages, pageOf, pageNorms, unread };
}

/* ------------------------------------------------------------------ */
/* Penggabungan aturan + AI + isian pengguna                            */
/* ------------------------------------------------------------------ */

type Merged = {
    fields: FrontFields;
    sources: Record<FrontFieldKey, FieldSource>;
    kataPengantar: ParaNode[] | null;
    abstrak: { paras: ParaNode[]; keywords: ParaNode | null } | null;
    abstract: { paras: ParaNode[]; keywords: ParaNode | null } | null;
    signatures: AiData["signatures"];
    pernyataanPics: Pic[];
    kataPengantarPics: Pic[];
    disagree: string[];
};

/**
 * Nilai akhir tiap data dari dua sumber (aturan & AI), keduanya lewat pemeriksaan umum yang sama.
 *  - satu lolos → itu; dua lolos & sama → itu;
 *  - dua lolos & beda → bacaan aturan (deterministik & teruji), perbedaannya dicatat;
 *  - isian yang memuat isian lain (judul Inggris berisi judul Indonesia, judul berisi nama) tidak lolos;
 *  - tidak ada yang lolos → kosong (titik-titik, diisi lewat form). Lebih baik kosong daripada salah.
 */
function merge(x: Extraction, ai: AiData | null, input: FrontFieldInput, pageNorms: string[]): Merged {
    const fields = emptyFieldsOf();
    const sources = nulls<FieldSource>(null);
    const disagree: string[] = [];
    const txt = (r: Rich | string | null | undefined) => (r == null ? "" : typeof r === "string" ? r : richText(r).replace(/\s+/g, " ").trim());
    void pageNorms;
    // Judul Inggris dari halaman abstract bila sampul tidak memuatnya.
    const ruleOf = (k: FrontFieldKey): Rich | string | null => {
        let v = x.fields[k] ?? (k === "judulInggris" && x.abstract?.title?.length ? x.abstract.title : null);
        if (v && typeof v !== "string") v = collapseRepeat(v);
        return v;
    };
    // Nama & NPM lebih dulu (penanda tangan & judul dibandingkan dengannya), lalu judul, judul Inggris.
    const first: FrontFieldKey[] = ["nama", "npm", "judul", "judulInggris"];
    const order: FrontFieldKey[] = [...first, ...FIELD_KEYS.filter((k) => !first.includes(k))];
    /** Isian tidak boleh memuat isian lain (judul Inggris berisi judul Indonesia, judul berisi nama/NPM, …). */
    const clash = (k: FrontFieldKey, v: string): string | null => {
        if (k !== "judul" && k !== "judulInggris") return null;
        const n = squashNorm(v);
        for (const o of ["judul", "judulInggris", "nama", "npm"] as const) {
            if (o === k) continue;
            const ov = fields[o];
            const on = ov ? squashNorm(txt(ov)) : "";
            if (on.length >= 6 && n.includes(on)) return FIELD_LABEL[o].toLowerCase();
        }
        return null;
    };
    const ok = (k: FrontFieldKey, v: Rich | string | null | undefined, nama: string | null, from: string) => {
        if (v == null) return false;
        if (!fieldOk(k, txt(v), k === "nama" ? null : nama)) return false;
        const c = clash(k, txt(v));
        if (c) { disagree.push(`${FIELD_LABEL[k]}: bacaan ${from} "${txt(v).slice(0, 80)}" memuat ${c}; tidak dipakai.`); return false; }
        return true;
    };
    for (const k of order) {
        const nama = typeof fields.nama === "string" ? fields.nama : null;
        const rule = ruleOf(k);
        const fromAi = ai?.fields[k];
        const ruleOk = ok(k, rule, nama, "aturan");
        const aiOk = ok(k, fromAi, nama, "AI");
        if (rule !== null && !ruleOk) disagree.push(`${FIELD_LABEL[k]}: bacaan aturan "${txt(rule).slice(0, 80)}" tidak lolos pemeriksaan, tidak dipakai.`);
        // Aturan (deterministik, teruji) didahulukan; AI mengisi bila aturan kosong/tidak lolos.
        // Bila keduanya lolos tetapi berbeda, aturan dipakai dan perbedaannya dicatat.
        let pick: { v: Rich | string; src: FieldSource } | null = null;
        if (ruleOk && aiOk) {
            pick = { v: rule!, src: simple(txt(rule)) === simple(txt(fromAi)) ? "ai" : "aturan" };
            if (pick.src === "aturan") disagree.push(`${FIELD_LABEL[k]}: AI membaca "${txt(fromAi).slice(0, 80)}", aturan membaca "${txt(rule).slice(0, 80)}" — dipakai hasil aturan.`);
        } else if (ruleOk) pick = { v: rule!, src: "aturan" };
        else if (aiOk) pick = { v: fromAi!, src: "ai" };
        if (pick) { (fields as Record<string, unknown>)[k] = pick.v; sources[k] = pick.src; }
    }
    // Satu orang tidak menandatangani dua jabatan: nama kembar → yang kedua dikosongkan.
    const signers = ["pembimbingUtama", "pembimbingAnggota", "ketuaSidang", "dekan"] as const;
    const seen = new Set<string>();
    for (const k of signers) {
        const v = fields[k];
        if (!v) continue;
        const n = squashNorm(v);
        if (seen.has(n)) {
            disagree.push(`${FIELD_LABEL[k]}: nama "${v}" sudah dipakai jabatan lain; dikosongkan.`);
            fields[k] = null;
            sources[k] = null;
        } else seen.add(n);
    }

    for (const k of FIELD_KEYS) {
        const v = input[k];
        if (typeof v !== "string") continue;
        const t = v.replace(/\s+/g, " ").trim();
        if (!t) continue;
        (fields as Record<string, unknown>)[k] = k === "judul" || k === "judulInggris" ? plainRich(t) : t;
        sources[k] = "pengguna";
    }
    const absOf = (a: Extraction["abstrak"]) => (a ? { paras: a.paras, keywords: a.keywords } : null);
    // Isi prosa (kata pengantar, abstrak, abstract): TIDAK BOLEH ada teks yang hilang. Rentang AI
    // dipakai hanya bila memuat semua teks bacaan aturan; kalau AI memotong (mis. alinea terakhir
    // tertinggal), bacaan aturan yang dipakai.
    const len = (ps: ParaNode[] | null | undefined) => (ps ?? []).reduce((n, p) => n + squashNorm(p.text).length, 0);
    const prose = <T,>(fromAi: T | null | undefined, rule: T | null, paras: (v: T) => ParaNode[], label: string): T | null => {
        if (!fromAi) return rule;
        if (rule && len(paras(fromAi)) < len(paras(rule))) {
            disagree.push(`${label}: rentang AI lebih pendek dari isi naskah; dipakai bacaan aturan agar tidak ada teks yang hilang.`);
            return rule;
        }
        return fromAi;
    };
    return {
        fields, sources, disagree,
        kataPengantar: prose(ai?.kataPengantar, x.kataPengantar?.paras ?? null, (v) => v, "Kata pengantar"),
        abstrak: prose(ai?.abstrak, absOf(x.abstrak), (v) => v.paras, "Abstrak"),
        abstract: prose(ai?.abstract, absOf(x.abstract), (v) => v.paras, "Abstract"),
        signatures: { ...x.signatures, ...ai?.signatures },
        pernyataanPics: ai?.pernyataanPics?.length ? ai.pernyataanPics : x.pernyataanPics,
        kataPengantarPics: ai?.kataPengantarPics?.length ? ai.kataPengantarPics : x.kataPengantar?.closingPics ?? [],
    };
}

type AiRead = { data: AiData | null; note: string | null; listing: string; answer: unknown };

async function aiRead(ai: FrontAI | undefined, items: Item[], pageOf: Map<Item, FrontPageKind>): Promise<AiRead> {
    const listing = buildListing(items, (it) => pageOf.get(it));
    const base = { listing: listing.text, answer: null as unknown };
    if (!ai || !listing.lines.size) return { ...base, data: null, note: null };
    try {
        const ptr = await ai(listing.text);
        if (!ptr) return { ...base, data: null, note: "Pembacaan AI tidak tersedia; data bagian depan dibaca dengan aturan saja." };
        console.info("[front-template] jawaban AI:", JSON.stringify(ptr));
        const data = applyPointers(ptr, listing, (it) => pageOf.get(it));
        if (data.rejected.length) console.warn("[front-template] tunjukan AI ditolak:", data.rejected.join("; "));
        return { ...base, answer: ptr, data, note: null };
    } catch (e) {
        console.warn("[front-template] AI gagal:", e instanceof Error ? e.message : e);
        return { ...base, data: null, note: "Pembacaan AI gagal; data bagian depan dibaca dengan aturan saja." };
    }
}

/**
 * Rekaman diagnosis: bila ada isian kosong/ditolak/berbeda atau pemeriksaan gagal, daftar baris
 * bagian depan + jawaban AI + keputusan dikirim ke "penampung" (bukan naskahnya). Bawaan: dicetak
 * ke console (browser: DevTools → Console). Di Node (front/node.ts) disimpan sebagai berkas JSON di
 * folder front-debug. Berkas ini bisa diputar ulang sebagai tes (`frontAI` di test-docs).
 */
export type DiagnosisSink = (payload: Record<string, unknown>) => Promise<void> | void;
let diagnosisSink: DiagnosisSink = (payload) => {
    console.info("[front-template] rekaman diagnosis (salin JSON ini bila melaporkan masalah):", JSON.stringify(payload));
};
export function setDiagnosisSink(fn: DiagnosisSink): void { diagnosisSink = fn; }

async function saveDiagnosis(payload: Record<string, unknown>): Promise<void> {
    try { await diagnosisSink(payload); } catch (e) {
        console.warn("[front-template] rekaman diagnosis gagal disimpan:", e instanceof Error ? e.message : e);
    }
}

function missingPages(m: Merged): FrontPageKind[] {
    const out: FrontPageKind[] = [];
    if (!m.kataPengantar?.length) out.push("kataPengantar");
    if (!m.abstrak?.paras.length) out.push("abstrak");
    if (!m.abstract?.paras.length) out.push("abstract");
    return out;
}

/** Pembacaan tanpa mengubah dokumen: untuk form konfirmasi data sebelum format diterapkan. */
export async function analyzeFront(model: DocModel, config: FormatConfig, ai?: FrontAI, footers: PartLines = new Map()): Promise<FrontReport> {
    const spec = presetById(config.presetId).front;
    if (!spec) return empty(null);
    const r = readNaskah(model, footers);
    if (typeof r === "string") return empty(r);
    const { data } = await aiRead(ai, r.items, r.pageOf);
    const m = merge(r.x, data, {}, r.pageNorms);
    const res = empty(null);
    res.found = r.found;
    res.missing = missingPages(m);
    res.fields = asText(m.fields);
    res.sources = m.sources;
    res.unmapped = unmappedOf(r);
    return res;
}

/** Sama dengan analyzeFront, langsung dari berkas DOCX (untuk endpoint form konfirmasi). */
export async function analyzeFrontDocx(input: ArrayBuffer | Uint8Array, config: FormatConfig, ai?: FrontAI): Promise<FrontReport> {
    const zip = await JSZip.loadAsync(input);
    const xml = await zip.file("word/document.xml")?.async("string");
    if (!xml) return empty("file bukan DOCX yang sah.");
    return analyzeFront(parseDocument(xml), config, ai, await footerLines(zip));
}

function unmappedOf(r: Read): { page: string; text: string }[] {
    const out = [...r.x.unmapped];
    for (const s of r.skipped) {
        for (const it of s.items) if (it.k === "line") out.push({ page: "Halaman lain", text: it.text });
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Penyusunan                                                           */
/* ------------------------------------------------------------------ */

const SMALL_WORDS = new Set(["dan", "di", "ke", "dari", "pada", "dalam", "terhadap", "untuk", "yang", "dengan", "atau", "serta", "oleh", "sebagai", "tentang", "antara", "bagi", "melalui"]);

/* ------------------------------------------------------------------ */
/* Halaman daftar                                                       */
/* ------------------------------------------------------------------ */

const LIST_ITEM = {
    tables: /^\s*(?:tabel|table)\s*\.?\s*\d+(?:[.\-]\d+)*\s*(?:[.:\u2013\u2014-]\s*\S|\s+[A-Z(\u201c"])/i,
    figures: /^\s*(?:gambar|ilustrasi|grafik|bagan|foto|figure)\s*\.?\s*\d+(?:[.\-]\d+)*\s*(?:[.:\u2013\u2014-]\s*\S|\s+[A-Z(\u201c"])/i,
    appendices: /^\s*lampiran\s+(?:\d{1,2}|[A-Z]|[IVXL]{1,5})(?![A-Za-z0-9])\s*(?:[.:\u2013\u2014-]\s*\S|\s+[A-Z(\u201c"])/i,
};

/**
 * Judul halaman daftar yang disisipkan sesudah template: DAFTAR ISI selalu; daftar tabel/ilustrasi/
 * lampiran hanya bila isi skripsi memuat keterangannya. Isinya disusun modul daftar (column-lists /
 * toc) seperti biasa. Section halaman daftar dibuat baru (mulai di halaman baru), dan section BAB I
 * juga dipaksa mulai di halaman baru walau naskah menyetelnya "continuous".
 */
function listPages(model: DocModel, end: number, heads: { contents: string; tables: string; figures: string; appendices: string }): { xml: string; headings: string[] } {
    const rest = flatten(model.body.slice(end));
    const has = (re: RegExp) => rest.some((n) => n.kind === "paragraph" && n.text.length <= 300 && re.test(n.text));
    const headings = [heads.contents];
    if (has(LIST_ITEM.tables)) headings.push(heads.tables);
    if (has(LIST_ITEM.figures)) headings.push(heads.figures);
    if (has(LIST_ITEM.appendices)) headings.push(heads.appendices);

    // Section asal BAB I (sectPr pertama sesudah bagian depan) dipakai sebagai dasar section daftar.
    const nextSect = rest.find((n): n is ParaNode => n.kind === "paragraph" && n.hasSectPr);
    const base = (nextSect ? /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(nextSect.xml)?.[0] : null) ?? model.bodySectPr ?? "<w:sectPr/>";
    const noContinuous = (s: string) => s.replace(/<w:type\b[^>]*\/>/g, "");
    if (nextSect) nextSect.xml = nextSect.xml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, (x) => noContinuous(x));
    else if (model.bodySectPr) model.bodySectPr = noContinuous(model.bodySectPr);

    const xml = headings.map((h, i) =>
        `<w:p><w:pPr>${i ? "<w:pageBreakBefore/>" : ""}</w:pPr><w:r><w:t xml:space="preserve">${h}</w:t></w:r></w:p>`).join("") +
        `<w:p><w:pPr>${noContinuous(base)}</w:pPr></w:p>`;
    return { xml, headings };
}

const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

const squashNorm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "");

/** Paragraf teks baku template (bukan isian contoh, bukan isi prosa) yang tidak muncul di hasil. */
function templateDrift(pages: { kind: FrontPageKind; xml: string }[], spec: NonNullable<ReturnType<typeof presetById>["front"]>, outText: string): string[] {
    const out = squashNorm(outText);
    const samples = Object.values(spec.sample).filter(Boolean).map((v) => squashNorm(v!));
    const missing: string[] = [];
    const walk = (blocks: BlockNode[], acc: string[]) => {
        for (const b of blocks) {
            if (b.kind === "paragraph") acc.push(b.text);
            else if (b.kind === "table") for (const row of b.rows) for (const c of row.cells) walk(c.children, acc);
            else if (b.kind === "container") walk(b.children, acc);
        }
        return acc;
    };
    for (const p of pages) {
        const prosa = p.kind === "kataPengantar" || p.kind === "abstrak" || p.kind === "abstract";
        for (const t of walk(parseBlocks(p.xml, 0), [])) {
            const n = squashNorm(t);
            if (!n || samples.some((s) => n.includes(s))) continue;
            // Halaman prosa: hanya judul halamannya yang baku.
            if (prosa && !/^(kata\s+pengantar|abstrak|abstract|penulis)$/i.test(t.trim())) continue;
            if (!out.includes(n)) missing.push(t.trim());
        }
    }
    return missing;
}

function titleNotes(judul: string): string[] {
    const out: string[] = [];
    const words = judul.split(/\s+/).filter(Boolean);
    const wrong = [...new Set(words.slice(1).map((w) => w.replace(/[^A-Za-z]/g, "")).filter((w) => w && SMALL_WORDS.has(w.toLowerCase()) && /^[A-Z]/.test(w)))];
    if (wrong.length) out.push(`Judul memakai huruf kapital pada kata sambung/kata depan (${wrong.join(", ")}). Pedoman: kata sambung & kata depan ditulis huruf kecil. Judul tidak diubah otomatis — perbaiki bila perlu.`);
    if (words.length > 20) out.push(`Judul terdiri atas ${words.length} kata; pedoman membatasi maksimal 20 kata.`);
    const letters = judul.replace(/[^A-Za-z]/g, "");
    if (letters.length > 10 && letters.replace(/[^A-Z]/g, "").length / letters.length > 0.9) {
        out.push("Judul di sampul naskah ditulis huruf kapital semua; pedoman: huruf kapital di awal tiap kata (kecuali kata sambung/kata depan). Tidak diubah otomatis.");
    }
    return out;
}

export async function buildFrontMatter(
    zip: JSZip,
    model: DocModel,
    config: FormatConfig,
    input: FrontFieldInput = {},
    ai?: FrontAI,
): Promise<FrontResult> {
    const preset = presetById(config.presetId);
    const spec = preset.front;
    if (!spec) return empty(null);
    if (!spec.docTypes.includes(config.docType)) {
        const reason = `template bagian depan preset ini baru tersedia untuk ${spec.docTypes.join("/")}; jenis dokumen yang dipilih ${config.docType}.`;
        console.warn(`[front-template] ${preset.id}: template tidak dipakai — ${reason}`);
        return empty(reason);
    }
    // Mulai sini template WAJIB dipakai. Tidak ada jalan kembali ke penyusunan bagian depan cara lama:
    // kegagalan = error yang jelas (dan tercatat di log), bukan hasil yang diam-diam berantakan.
    const fail = (msg: string, status: number, cause?: unknown): never => {
        console.error(`[front-template] ${preset.id}: GAGAL — ${msg}`, cause ?? "");
        throw Object.assign(new Error(`Bagian depan tidak dapat disusun dari template ${preset.name}: ${msg}`), { status });
    };
    const r = readNaskah(model, await footerLines(zip));
    if (typeof r === "string") return fail(`${r} Pastikan BAB I diawali judul bab (mis. "BAB I" / "I PENDAHULUAN").`, 422);

    let tpl!: Awaited<ReturnType<typeof loadTemplate>>;
    try {
        tpl = await loadTemplate(spec);
    } catch (e) {
        fail(e instanceof TemplateError ? e.message : "berkas template tidak dapat dibaca.", 500, e);
    }

    const aiRes = await aiRead(ai, r.items, r.pageOf);
    const m = merge(r.x, aiRes.data, input, r.pageNorms);

    // Tempat & tanggal yang tidak ada di naskah diisi otomatis (bukan tanggal contoh template).
    const now = new Date();
    for (const k of ["tempatTanggalPernyataan", "tempatTanggalKataPengantar"] as const) {
        if (m.fields[k]) continue;
        m.fields[k] = `${spec.defaultPlace}, ${BULAN[now.getMonth()]} ${m.fields.tahun ?? now.getFullYear()}`;
        m.sources[k] = "bawaan";
    }

    const lineMap = new Map<ParaNode, LineItem[]>();
    for (const it of r.items) if (it.k === "line") lineMap.set(it.para, [...(lineMap.get(it.para) ?? []), it]);

    const bodyXml = model.body.map(serializeBlock).join("");
    const maxOf = (re: RegExp) => Math.max(0, ...[...bodyXml.matchAll(re)].map((x) => Number(x[1])));
    const rootTag = /<w:document\b[^>]*>/.exec(model.head)?.[0] ?? "";
    let fill!: Awaited<ReturnType<typeof fillTemplate>>;
    try {
        fill = await fillTemplate(tpl, spec, {
            fields: m.fields,
            linesOf: (p) => lineMap.get(p) ?? [],
            kataPengantar: m.kataPengantar,
            abstrak: m.abstrak,
            abstract: m.abstract,
            signatures: m.signatures,
            pernyataanPics: m.pernyataanPics,
            kataPengantarPics: m.kataPengantarPics,
        }, {
            zip, rootTag,
            maxDocPr: maxOf(/<wp:docPr\b[^>]*\bid="(\d+)"/g),
            maxBookmark: maxOf(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g),
            tocField: config.listStyle !== "columns",
        });
    } catch (e) {
        fail(e instanceof TemplateError ? e.message : "template tidak dapat diisi.", 500, e);
    }

    // Namespace yang dipakai template/tanda tangan tetapi belum dideklarasikan naskah.
    const ns = missingNamespaces(rootTag, fill.xml, [tpl.rootTag, ...[...m.pernyataanPics, ...m.kataPengantarPics, ...Object.values(m.signatures)].map((p) => p?.xml ?? ""), model.head]);
    if (ns && rootTag) model.head = model.head.replace(rootTag, rootTag.replace(/>$/, `${ns}>`));

    // Halaman daftar (DAFTAR ISI dst.) selalu disusun formatter sesudah template; daftar buatan
    // naskah dalam bentuk apa pun (teks, field, content control) ikut terbuang bersama bagian depan.
    const lists = listPages(model, r.region.end, spec.listHeadings);

    /* ---- integritas ---- */
    const regionXml = model.body.slice(0, r.region.end).map(serializeBlock).join("");
    const before = fingerprint(regionXml);
    const after = fingerprint(fill.xml + lists.xml);
    const text = asText(m.fields);
    const issues: VerifyIssue[] = [];
    const original = simple(before.text);
    for (const k of FIELD_KEYS) {
        if (m.sources[k] !== "aturan" && m.sources[k] !== "ai") continue;
        const v = simple(text[k] ?? "");
        if (v && !original.includes(v)) {
            issues.push({ level: "error", message: `Bagian depan: data "${FIELD_LABEL[k]}" (${text[k]}) tidak ditemukan di teks naskah asli.` });
        }
    }
    const pool = [...after.embeddedImageIds];
    const removedImageIds: string[] = [];
    for (const id of before.embeddedImageIds) {
        const at = pool.indexOf(id);
        if (at === -1) removedImageIds.push(id);
        else pool.splice(at, 1);
    }

    /* ---- catatan ---- */
    const notes: string[] = [];
    const missing = missingPages(m);
    notes.push(`Bagian depan disusun dari template ${preset.name}: ${spec.pages.map((k) => FRONT_PAGE_LABEL[k]).join(" → ")}. Tata letak, huruf, jarak, logo, dan nomor halaman mengikuti template; dari naskah hanya diambil teks dan tanda tangan.`);
    const viaAi = FIELD_KEYS.some((k) => m.sources[k] === "ai");
    notes.push(`Data bagian depan${viaAi ? " (dibaca dengan bantuan AI, disalin dari teks naskah)" : ""}: ${FIELD_KEYS.filter((k) => text[k]).map((k) => `${FIELD_LABEL[k]}: ${text[k]}${m.sources[k] === "pengguna" ? " (isian pengguna)" : ""}`).join("; ")}.`);
    if (aiRes.note) notes.push(aiRes.note);
    if (aiRes.data?.rejected.length) notes.push(`Tunjukan AI untuk ${aiRes.data.rejected.join(", ")} tidak lolos pemeriksaan; dipakai pembacaan aturan.`);
    if (fill.placeholders.length) {
        notes.push(`Data berikut tidak ditemukan di naskah dan diisi titik-titik: ${fill.placeholders.map((k) => FIELD_LABEL[k]).join(", ")}. Lengkapi lewat isian data bagian depan atau langsung di Word.`);
    }
    for (const k of missing) notes.push(`Isi ${FRONT_PAGE_LABEL[k].toLowerCase()} tidak ditemukan di naskah; halaman ${FRONT_PAGE_LABEL[k]} template dibiarkan tanpa isi.`);
    notes.push(...m.disagree);
    notes.push(...r.x.conflicts);
    const unmapped = unmappedOf(r);
    if (unmapped.length) {
        notes.push(`${unmapped.length} baris bagian depan naskah bukan isian template dan TIDAK dicantumkan: ${unmapped.slice(0, 8).map((u) => `${u.page}: "${u.text.slice(0, 60)}"`).join("; ")}${unmapped.length > 8 ? "; …" : ""}. Periksa apakah baris ini memang perlu.`);
    }
    const why = [...new Set([...r.x.dropped.filter((d) => !fill.usedPics.includes(d.pic)).map((d) => d.why), ...fill.unusedPics.map(() => "tanda tangan/materai yang tidak mendapat tempat di template")])];
    const droppedCount = r.x.dropped.filter((d) => !fill.usedPics.includes(d.pic)).length + fill.unusedPics.length;
    if (droppedCount) notes.push(`${droppedCount} gambar/bentuk bagian depan naskah tidak dipakai: ${why.join("; ")}.`);
    if (fill.usedPics.length) notes.push(`${fill.usedPics.length} tanda tangan/materai dari naskah dipasang pada posisi tanda tangan template.`);
    if (r.blankPages) notes.push(`${r.blankPages} halaman kosong di bagian depan naskah dihapus.`);
    if (m.fields.judul) notes.push(...titleNotes(richText(m.fields.judul)));
    if (m.abstrak?.paras.length) {
        const words = m.abstrak.paras.map((p) => p.text).join(" ").split(/\s+/).filter(Boolean).length;
        if (words < 100 || words > 200) notes.push(`Abstrak berisi ${words} kata; pedoman: 100–200 kata.`);
        if (m.abstrak.paras.length > 1) notes.push(`Abstrak terdiri atas ${m.abstrak.paras.length} alinea; pedoman: satu alinea. Tidak digabung otomatis.`);
        const kw = m.abstrak.keywords?.text.replace(/^[^:：]*[:：]/, "").split(/[,;]/).map((s) => s.trim()).filter(Boolean) ?? [];
        if (kw.length > 5) notes.push(`Kata kunci berjumlah ${kw.length}; pedoman: maksimal 5.`);
    }
    notes.push(...fill.warnings);
    const auto = FIELD_KEYS.filter((k) => m.sources[k] === "bawaan");
    if (auto.length) notes.push(`${auto.map((k) => FIELD_LABEL[k]).join(", ")} tidak ditemukan di naskah dan diisi otomatis (${auto.map((k) => text[k]).join("; ")}). Sesuaikan bila perlu.`);

    // Pemeriksaan: teks baku template (yang bukan isian) harus utuh di hasil.
    const drift = templateDrift(tpl.pages, spec, after.text);
    for (const t of drift) issues.push({ level: "error", message: `Bagian depan menyimpang dari template: teks "${t.slice(0, 80)}" tidak ada di hasil.` });

    // Cek akhir isian: di satu halaman tidak boleh ada baris yang sama dua kali (judul dobel dsb.),
    // dan judul Inggris tidak boleh memuat judul Indonesia (atau sebaliknya).
    for (const p of fill.pageTexts) {
        const seenLine = new Set<string>();
        for (const l of p.lines) {
            const k = squashNorm(l);
            if (k.length < 20) continue;
            if (seenLine.has(k)) {
                const msg = `Baris muncul dua kali di ${FRONT_PAGE_LABEL[p.kind]}: "${l.slice(0, 80)}".`;
                console.error(`[front-template] ${preset.id}: ${msg}`);
                issues.push({ level: "error", message: msg });
            }
            seenLine.add(k);
        }
    }
    {
        const j = squashNorm(text.judul ?? "");
        const e = squashNorm(text.judulInggris ?? "");
        if (j && e && (j.includes(e) || e.includes(j))) {
            const msg = "Judul dan judul bahasa Inggris saling memuat; periksa isian judul.";
            console.error(`[front-template] ${preset.id}: ${msg}`);
            issues.push({ level: "error", message: msg });
        }
    }

    if (fill.placeholders.length || aiRes.data?.rejected.length || m.disagree.length || issues.length || r.x.conflicts.length || r.unread.length) {
        await saveDiagnosis({
            versi: FRONT_VERSION, preset: preset.id, nama: text.nama,
            isian: text, sumber: m.sources, kosong: fill.placeholders,
            ditolakAI: aiRes.data?.rejected ?? [], perbedaan: m.disagree, catatan: r.x.conflicts,
            masalah: issues.map((i) => i.message), tidakTerbaca: r.unread, jawabanAI: aiRes.answer, daftarBaris: aiRes.listing,
        });
    }
    console.info(`[front-template] ${preset.id}: versi ${FRONT_VERSION}; disusun dari template; AI ${ai ? (aiRes.data ? "dipakai" : "gagal/tidak tersedia") : "tidak aktif"}; ` +
        `kosong: ${fill.placeholders.join(", ") || "-"}; otomatis: ${auto.join(", ") || "-"}; menyimpang: ${drift.length}; ` +
        `sumber: ${FIELD_KEYS.map((k) => `${k}=${m.sources[k] ?? "kosong"}`).join(" ")}.`);

    const preview: FrontPreviewLine[] = fill.pageTexts.flatMap((p) =>
        p.lines.map((t) => ({ role: p.kind, zone: p.kind === "kataPengantar" || p.kind === "abstrak" || p.kind === "abstract" ? "front" : "cover", text: t })));

    // Seluruh bagian depan naskah (sebelum BAB I) diganti template + halaman daftar.
    const raw: RawNode = { kind: "raw", xml: fill.xml };
    const listBlocks = parseBlocks(lists.xml, 9_000_000);
    model.body.splice(0, r.region.end, raw, ...listBlocks);
    console.info(`[front-template] ${preset.id}: halaman daftar: ${lists.headings.join(", ")}.`);

    // Cek akhir: sesudah template + halaman daftar harus langsung BAB I (tidak ada sisa bagian depan
    // naskah, mis. Daftar Isi otomatis Word yang tertinggal).
    const restText = flatten(model.body.slice(1 + listBlocks.length))
        .filter((n): n is ParaNode => n.kind === "paragraph" && n.text.trim() !== "").slice(0, 2).map((n) => n.text.trim());
    if (restText.length && !isChapterStart(restText[0], restText[1] ?? null)) {
        const msg = `Sesudah bagian depan tidak langsung BAB I: "${restText[0].slice(0, 80)}". Ada sisa bagian depan naskah yang tidak terbuang.`;
        console.error(`[front-template] ${preset.id}: ${msg}`);
        issues.push({ level: "error", message: msg });
    }

    return {
        built: true, reason: null, found: r.found, missing,
        fields: text, sources: m.sources, unmapped,
        notes, issues,
        romanStarted: true,
        regions: [
            { beforeLines: before.text.split("\n").filter(Boolean), afterLines: fingerprint(fill.xml).text.split("\n").filter(Boolean), beforeText: "", afterText: "" },
            ...lists.headings.map((h) => ({ beforeLines: [], afterLines: [h], beforeText: "", afterText: "" })),
        ],
        delta: {
            drawings: after.drawings - before.drawings,
            tables: after.tables - before.tables,
            rows: after.rows - before.rows,
            cells: after.cells - before.cells,
        },
        removedImageIds,
        tocRows: fill.tocRows,
        preview,
    };
}
/**
 * Lampiran dibekukan: isinya dikembalikan PERSIS seperti naskah asli.
 *
 * Isi lampiran berbeda-beda di tiap naskah (tabel data mentah, formulir, foto, surat, biodata),
 * jadi aturan format isi tidak bisa diterapkan dengan aman. Wilayah lampiran = dari judul
 * "LAMPIRAN" pertama SETELAH Daftar Pustaka sampai akhir dokumen (termasuk Riwayat Hidup).
 *
 * Cara kerjanya transaksional: semua modul tetap berjalan seperti biasa (judul lampiran tetap
 * terbaca untuk Daftar Lampiran, perkiraan halaman tetap dihitung), lalu di akhir setiap blok
 * lampiran diganti kembali dengan XML aslinya. Yang dibawa dari hasil format hanya:
 *   - pengaturan halaman section (ukuran kertas, margin, nomor halaman: header/footer),
 *   - penanda Daftar Lampiran (field TC & bookmark yang tidak tampil),
 *   - pindah halaman sebelum judul lampiran pertama,
 *   - gambar/tabel yang lebih lebar dari area teks diperkecil proporsional (isinya tetap).
 * Batas wilayah tidak pasti → tidak dibekukan, dilaporkan.
 */
import { parseBlocks, serializeBlock, serializeDocument, type BlockNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { APPENDIX_HEADING_RE, CHAPTER_RE, REFERENCE_HEADING_RE } from "../structure/roles";
import { fitDrawings, type ImageReport } from "./images";
import { fitTableWidth } from "./layout-groups";
import { isLandscape } from "./package-parts";
import { fingerprint, removeGroupIdx, type DocFingerprint, type VerifyIssue, type VerifyOptions } from "./verify";
import { withProps, PPR_ORDER } from "./xml";

export type AppendixFreeze = {
    /** Blok tingkat atas wilayah lampiran (objek di model.body). */
    blocks: BlockNode[];
    /** XML asli tiap blok. */
    original: string[];
    /** Judul lampiran pertama (untuk laporan). */
    title: string;
};

export type FreezeResult = { frozen: AppendixFreeze | null; note: string | null };

/** Entri daftar isi lama ("LAMPIRAN ........ 45", field PAGEREF), bukan judul. */
const tocLike = (p: ParaNode) => /PAGEREF|w:anchor="_Toc|\.{5,}|…{2,}/.test(p.xml) || /\t\s*\d+\s*$/.test(p.text) || /^toc/i.test(p.styleId ?? "");
const plain = (p: ParaNode) => p.text.replace(/\s+/g, " ").trim();
const headingish = (t: string) => t.length > 0 && t.length <= 80 && !/[.;,]$/.test(t);

function nestedParagraphs(n: BlockNode, out: ParaNode[] = []): ParaNode[] {
    if (n.kind === "paragraph") out.push(n);
    else if (n.kind === "container") n.children.forEach((c) => nestedParagraphs(c, out));
    else if (n.kind === "table") for (const r of n.rows) for (const c of r.cells) c.children.forEach((x) => nestedParagraphs(x, out));
    return out;
}

/** Panggil pada model ASLI (sebelum modul lain). */
export function findAppendixRegion(model: DocModel): FreezeResult {
    const body = model.body;
    const isPara = (n: BlockNode): n is ParaNode => n.kind === "paragraph";
    // Judul Daftar Pustaka (bukan entri daftar isi) — yang terakhir.
    let ref = -1;
    body.forEach((n, i) => { if (isPara(n) && !tocLike(n) && REFERENCE_HEADING_RE.test(plain(n)) && headingish(plain(n))) ref = i; });
    const appendixAfter = (from: number) => {
        for (let i = from + 1; i < body.length; i++) {
            const n = body[i];
            if (isPara(n) && !tocLike(n) && APPENDIX_HEADING_RE.test(plain(n)) && headingish(plain(n))) return i;
        }
        return -1;
    };
    const nestedAppendix = (from: number) => body.slice(from + 1).some((n) => !isPara(n) &&
        nestedParagraphs(n).some((p) => !tocLike(p) && APPENDIX_HEADING_RE.test(plain(p)) && headingish(plain(p))));

    if (ref === -1) {
        const loose = appendixAfter(-1);
        // Ada judul lampiran tetapi Daftar Pustaka tidak ditemukan: batasnya tidak pasti.
        const hasLater = loose !== -1 && body.slice(loose + 1).some((n) => isPara(n) && CHAPTER_RE.test(plain(n)));
        return { frozen: null, note: loose !== -1 && !hasLater ? "Lampiran tidak dibekukan: judul Daftar Pustaka tidak ditemukan, jadi awal lampiran tidak bisa dipastikan. Isi lampiran ikut dirapikan — periksa bagian lampiran." : null };
    }
    const a = appendixAfter(ref);
    if (a === -1) {
        return { frozen: null, note: nestedAppendix(ref) ? "Lampiran tidak dibekukan: judul LAMPIRAN berada di dalam tabel/kotak, batasnya tidak bisa dipastikan. Isi lampiran ikut dirapikan — periksa bagian lampiran." : null };
    }
    // Sesudah awal lampiran tidak boleh ada judul bab / Daftar Pustaka lagi (berarti salah tebak).
    const odd = body.slice(a + 1).find((n) => {
        if (!isPara(n) || tocLike(n)) return false;
        const t = plain(n);
        if (!headingish(t)) return false;
        return REFERENCE_HEADING_RE.test(t) || (CHAPTER_RE.test(t) && t === t.toUpperCase());
    });
    if (odd) {
        return { frozen: null, note: `Lampiran tidak dibekukan: setelah judul "${plain(body[a] as ParaNode).slice(0, 40)}" masih ada judul "${plain(odd as ParaNode).slice(0, 40)}", jadi batas lampiran tidak pasti. Isi lampiran ikut dirapikan — periksa bagian lampiran.` };
    }
    const blocks = body.slice(a);
    return { frozen: { blocks, original: blocks.map(serializeBlock), title: plain(body[a] as ParaNode) }, note: null };
}

/** Semua paragraf (termasuk di tabel) wilayah beku — modul lain tidak boleh menjadikannya kandidat perubahan. */
export function frozenParagraphs(f: AppendixFreeze | null): Set<ParaNode> {
    const s = new Set<ParaNode>();
    if (f) for (const b of f.blocks) nestedParagraphs(b).forEach((p) => s.add(p));
    return s;
}

const SECT_RE = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g;
const TC_RE = /<w:r><w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> TC "[\s\S]*?<\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="end"\/><\/w:r>/g;
const BM_START_RE = /<w:bookmarkStart\b[^>]*\/>/g;

/** Penanda tak tampil (TC, bookmark) yang ditambahkan formatter ke paragraf ini → dibawa ke XML asli. */
function carryMarkers(orig: string, now: string): string {
    let out = orig;
    const tcs = (now.match(TC_RE) ?? []).filter((t) => !orig.includes(t));
    if (tcs.length) {
        const i = out.lastIndexOf("</w:p>");
        if (i !== -1) out = out.slice(0, i) + tcs.join("") + out.slice(i);
    }
    const origNames = new Set([...orig.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]+)"/g)].map((m) => m[1]));
    const marks: string[] = [];
    for (const s of now.match(BM_START_RE) ?? []) {
        const name = /w:name="([^"]+)"/.exec(s)?.[1];
        const id = /w:id="([^"]+)"/.exec(s)?.[1];
        if (!name || !id || origNames.has(name)) continue;
        marks.push(`${s}<w:bookmarkEnd w:id="${id}"/>`);
    }
    if (marks.length) {
        const e = out.indexOf("</w:pPr>");
        const at = e !== -1 ? e + 8 : out.indexOf(">") + 1;
        out = out.slice(0, at) + marks.join("") + out.slice(at);
    }
    return out;
}

const startsPage = (xml: string) =>
    /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "");

export type RestoreReport = {
    restored: number;
    imagesScaled: number;
    tablesFitted: number;
    floatingOverflow: number;
    /** Untuk verifikasi: teks & jumlah objek wilayah lampiran asli dan sebelum dikembalikan. */
    beforeText: string;
    preRestoreText: string;
    /** Teks SELURUH dokumen sesaat sebelum lampiran dikembalikan. */
    preRestoreDocText: string;
    preRestore: ReturnType<typeof fingerprint>;
    originalFp: ReturnType<typeof fingerprint>;
    problem: string | null;
};

/**
 * Kembalikan wilayah lampiran ke XML asli. Dipanggil PALING AKHIR, sebelum dokumen diserialisasi
 * untuk verifikasi. `usable`/`landscapeUsable`: lebar area teks halaman tegak/landscape (twip).
 */
export function restoreAppendix(model: DocModel, f: AppendixFreeze, usable: number, landscapeUsable: number): RestoreReport {
    const rep: RestoreReport = {
        restored: 0, imagesScaled: 0, tablesFitted: 0, floatingOverflow: 0,
        beforeText: "", preRestoreText: "", preRestoreDocText: "", preRestore: fingerprint(""), originalFp: fingerprint(""), problem: null,
    };
    const body = model.body;
    const present = f.blocks.map((b) => body.indexOf(b)).filter((i) => i !== -1);
    if (!present.length) { rep.problem = "wilayah lampiran tidak ditemukan lagi setelah diformat"; return rep; }
    const start = Math.min(...present);
    const pre = body.slice(start);
    const preXml = pre.map(serializeBlock).join("");
    const origXml = f.original.join("");

    // Section di wilayah lampiran: pengaturan halaman HASIL format (margin, kertas, nomor halaman), urut.
    const preSects = preXml.match(SECT_RE) ?? [];
    const origSects = origXml.match(SECT_RE) ?? [];
    if (preSects.length !== origSects.length) {
        rep.problem = `jumlah section lampiran berubah (${origSects.length} → ${preSects.length})`;
        return rep;
    }
    let sk = 0;
    const nowOf = new Map<BlockNode, string>(pre.map((b) => [b, serializeBlock(b)]));
    const xmls = f.original.map((x, i) => {
        let out = x.replace(SECT_RE, () => preSects[sk++]);
        const b = f.blocks[i];
        const now = nowOf.get(b);
        if (b.kind === "paragraph" && now) out = carryMarkers(out, now);
        if (i === 0 && b.kind === "paragraph" && now && startsPage(now) && !startsPage(out))
            out = withProps(out, "w:pPr", { "w:pageBreakBefore": "<w:pageBreakBefore/>" }, PPR_ORDER);
        return out;
    });

    // Lebar area teks per blok: section tempat blok itu berada (sectPr berikutnya, atau sectPr body).
    const widthAt: number[] = new Array(xmls.length);
    let cur = model.bodySectPr ?? "";
    for (let i = xmls.length - 1; i >= 0; i--) {
        const s = xmls[i].match(SECT_RE);
        if (s) cur = s[s.length - 1];
        widthAt[i] = isLandscape(cur) ? landscapeUsable : usable;
    }

    let nextId = Math.max(0, ...body.map((b) => (b.kind === "paragraph" || b.kind === "table" ? b.id : 0))) + 1000;
    const img: ImageReport = { scaled: 0, floatingOverflow: 0 };
    const restored: BlockNode[] = [];
    xmls.forEach((x, i) => {
        const nodes = parseBlocks(x, nextId);
        nextId += 1000;
        for (const n of nodes) {
            if (n.kind === "paragraph" && n.hasDrawing) {
                const y = fitDrawings(n.xml, widthAt[i], img);
                if (y !== n.xml) Object.assign(n, { ...(parseBlocks(y, n.id)[0] as ParaNode), id: n.id });
            } else if (n.kind === "table") {
                if (fitTableWidth(n as TableNode, widthAt[i])) rep.tablesFitted++;
                // Gambar di sel tidak boleh melebihi lebar halaman.
                for (const r of n.rows) for (const c of r.cells) for (const p of c.children) {
                    if (p.kind === "paragraph" && p.hasDrawing) p.xml = fitDrawings(p.xml, widthAt[i], img);
                }
            }
            restored.push(n);
        }
    });
    rep.imagesScaled = img.scaled;
    rep.floatingOverflow = img.floatingOverflow;
    rep.preRestore = fingerprint(preXml);
    rep.preRestoreDocText = fingerprint(serializeDocument(model)).text;
    rep.originalFp = fingerprint(origXml);
    rep.beforeText = rep.originalFp.text;
    rep.preRestoreText = rep.preRestore.text;
    body.splice(start, body.length - start, ...restored);
    rep.restored = f.blocks.length;
    return rep;
}

/**
 * Verifikasi dengan lampiran dibekukan, dalam dua bagian:
 *   1. Lampiran: teks hasil harus IDENTIK dengan asli (baris per baris, urutan sama).
 *   2. Selebihnya: dibandingkan seperti biasa, tanpa baris lampiran. Baris lampiran yang tercatat
 *      di blok laporan modul (perubahan yang sudah dibatalkan) dikeluarkan dari blok itu memakai
 *      pencocokan yang SAMA dengan verifikasi (posisi, bukan sekadar teks), dan jumlah tabel/gambar
 *      yang ditambah/dibuang modul di lampiran tidak dihitung lagi.
 */
export function verifyWithFreeze(
    before: DocFingerprint, after: DocFingerprint, opts: VerifyOptions, r: RestoreReport,
): { before: DocFingerprint; after: DocFingerprint; opts: VerifyOptions; issues: VerifyIssue[] } {
    const issues: VerifyIssue[] = [];
    const lines = (t: string) => t.split("\n");
    const tailOrig = lines(r.originalFp.text).filter(Boolean);
    const A = lines(before.text);
    const P = lines(r.preRestoreDocText);
    const B = lines(after.text);
    const endsWith = (x: string[], tail: string[]) => tail.length <= x.length && tail.every((l, i) => x[x.length - tail.length + i] === l);
    if (!endsWith(A, tailOrig)) {
        issues.push({ level: "warn", message: "Pemeriksaan lampiran: wilayah lampiran asli tidak berada di akhir teks dokumen; dibandingkan sebagai satu dokumen." });
        return { before, after, opts, issues };
    }
    if (!endsWith(B, tailOrig)) {
        issues.push({ level: "error", message: "Teks lampiran hasil tidak sama dengan naskah asli." });
        return { before, after, opts, issues };
    }
    const tailPre = lines(r.preRestore.text).filter(Boolean);
    const cutA = A.length - tailOrig.length;
    const cutP = P.length - tailPre.length;

    // Tiru urutan verifikasi (blok besar dulu), catat indeks yang diambil tiap blok.
    const regions = (opts.regions ?? []).map((g) => ({ ...g }));
    const strip = (doc: string[], cut: number, side: "beforeLines" | "afterLines") => {
        let idx = doc.map((_, i) => i);
        for (const g of [...regions].sort((x, y) => y[side].length - x[side].length)) {
            const gone = removeGroupIdx(idx.map((i) => doc[i]), g[side]);
            if (!gone.size) continue;
            const keep: string[] = [];
            const picked = [...gone].sort((x, y) => x - y).map((k) => idx[k]);
            // Baris blok yang diambil dari wilayah lampiran dibuang dari blok; sisanya tetap.
            const tailKeys = new Set(picked.filter((i) => i >= cut).map((i) => doc[i].toLowerCase().replace(/\s+/g, "")));
            const tally = new Map<string, number>();
            for (const i of picked) if (i >= cut) { const k = doc[i].toLowerCase().replace(/\s+/g, ""); tally.set(k, (tally.get(k) ?? 0) + 1); }
            for (const l of g[side]) {
                const k = l.toLowerCase().replace(/\s+/g, "");
                const c = tally.get(k) ?? 0;
                if (tailKeys.has(k) && c > 0) { tally.set(k, c - 1); continue; }
                keep.push(l);
            }
            g[side] = keep;
            idx = idx.filter((_, k) => !gone.has(k));
        }
    };
    strip(A, cutA, "beforeLines");
    strip(P, cutP, "afterLines");

    const d = (k: "tables" | "rows" | "cells" | "drawings" | "equations") => r.preRestore[k] - r.originalFp[k];
    const add = opts.added ?? { tables: 0, rows: 0, cells: 0 };
    return {
        before: { ...before, text: A.slice(0, cutA).join("\n") },
        after: { ...after, text: B.slice(0, B.length - tailOrig.length).join("\n") },
        opts: {
            ...opts,
            regions,
            added: {
                tables: add.tables - d("tables"),
                rows: add.rows - d("rows"),
                cells: add.cells - d("cells"),
                drawings: (add.drawings ?? 0) - d("drawings"),
                equations: (add.equations ?? 0) - d("equations"),
            },
        },
        issues,
    };
}
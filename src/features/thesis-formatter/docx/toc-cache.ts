/**
 * Isi tampilan (cached result) field DAFTAR TABEL / DAFTAR GAMBAR.
 *
 * Field TOC sudah ditandai dirty, tetapi penampil yang tidak memperbarui field
 * (Google Docs, WPS, LibreOffice, pratinjau — atau Word bila pembaruan ditolak)
 * tetap menampilkan isi tersimpan, yang sebelumnya masih daftar lama dokumen
 * asli. Di sini isinya disusun ulang dari entri TC yang benar-benar ada.
 */
import type JSZip from "jszip";
import { parseBlocks, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { paragraphLine, type VerifyRegion } from "./verify";
import { stripFields } from "./captions";
import { escapeXml, withProps, PPR_ORDER, RPR_ORDER } from "./xml";
import type { Outline } from "../structure/outline";
import { leadingRaw, removeLeading } from "./lists";

type Flag = "t" | "g" | "l";
type Body = DocModel["body"];
type Block = { i0: number; i1: number; open: string };

/** Baris lanjutan entri daftar sejajar dengan awal judul (± lebar "Gambar 4.12 "). */
const HANG = 1191;
/** Entri daftar tabel/gambar: style Table of Figures (hanging indent ada di style). */
const entryPPr = (pPr: string) => {
    const style = `<w:pStyle w:val="TableofFigures"/>`;
    if (!pPr) return `<w:pPr>${style}</w:pPr>`;
    return pPr
        .replace(/<w:pStyle\b[^>]*\/>/, "")
        .replace(/<w:ind\b[^>]*\/>/, "")
        .replace(/<w:tabs>[\s\S]*?<\/w:tabs>/, "")
        .replace(/^<w:pPr>/, `<w:pPr>${style}`);
};

const decodeXml = (s: string) =>
    s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const runsOf = (xml: string) => [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
const instrOf = (xml: string) =>
    [...xml.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((m) => decodeXml(m[1])).join("");
const textOf = (xml: string) =>
    [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXml(m[1])).join("");
const pPrOf = (xml: string) => /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
const keyOf = (t: string) =>
    t.replace(/^\s*(?:tabel|table|gambar|figure)\s*\.?\s*(?:(?:[A-Za-z]{1,2}[.\-]?)?\d[\d.\-]*)?\s*[.:]?\s*/i, "")
        .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

function collectTc(body: Body, flag: Flag): string[] {
    const re = new RegExp(String.raw`\bTC\s+"([^"]*)"[^"]*?\\f\s*${flag}(?![a-z])`, "g");
    const out: string[] = [];
    for (const n of body) {
        if (n.kind !== "paragraph" || !/instr/.test(n.xml)) continue;
        const ins = instrOf(n.xml) + " " + [...n.xml.matchAll(/w:instr="([^"]*)"/g)].map((m) => decodeXml(m[1])).join(" ");
        for (const m of ins.matchAll(re)) out.push(m[1].replace(/\s+/g, " ").trim());
    }
    return out;
}

/** Teks paragraf ber-style tertentu, urut dokumen (untuk Daftar Lampiran). */
function collectStyled(body: Body, styleId: string): string[] {
    const re = new RegExp(`<w:pStyle w:val="${styleId}"/>`);
    return body
        .filter((n): n is ParaNode => n.kind === "paragraph" && re.test(n.xml))
        .map((n) => n.text.replace(/\s+/g, " ").trim())
        .filter(Boolean);
}

function locateBy(body: Body, tocRe: RegExp): Block | null {
    for (let i = 0; i < body.length; i++) {
        const n = body[i];
        if (n.kind !== "paragraph" || !n.xml.includes("TOC")) continue;
        const runs = runsOf(n.xml);
        for (let k = 0; k < runs.length; k++) {
            if (!/w:fldCharType="begin"/.test(runs[k][0])) continue;
            let ins = "";
            let s = k + 1;
            for (; s < runs.length && !/w:fldCharType="(?:separate|end)"/.test(runs[s][0]); s++) ins += instrOf(runs[s][0]);
            if (s >= runs.length || !/separate/.test(runs[s][0]) || !tocRe.test(ins)) continue;
            if (textOf(n.xml.slice(0, runs[k].index)).trim()) return null;
            const openEnd = runs[s].index! + runs[s][0].length;
            const open = n.xml.slice(runs[k].index!, openEnd);

            let depth = 1;
            const scan = (xml: string, from: number) => {
                for (const m of xml.slice(from).matchAll(/w:fldCharType="(begin|end)"/g)) {
                    depth += m[1] === "begin" ? 1 : -1;
                    if (depth === 0) return from + m.index!;
                }
                return -1;
            };
            let at = scan(n.xml, openEnd);
            let j = i;
            while (at === -1 && ++j < body.length) {
                const b = body[j];
                if (b.kind !== "paragraph") return null;
                at = scan(b.xml, 0);
            }
            if (at === -1) return null;
            const tail = (body[j] as ParaNode).xml.slice(at);
            if (textOf(tail.slice(tail.indexOf("</w:r>") + 6)).trim()) return null;
            return { i0: i, i1: j, open };
        }
    }
    return null;
}

function locate(body: Body, flag: Flag): Block | null {
    const label = flag === "t" ? "Tabel" : flag === "g" ? "Gambar" : null;
    const alt = label
        ? String.raw`|\\c\s*"?${label}"?`
        : String.raw`|\\t\s*"Judul Lampiran|\\c\s*"?Lampiran"?`;
    return locateBy(body, new RegExp(String.raw`^\s*TOC\b[\s\S]*(?:\\f\s*${flag}(?![a-z])` + alt + ")"));
}

function oldEntries(body: Body, b: Block) {
    const out: { label: string; key: string; page: string; used: boolean }[] = [];
    for (let i = b.i0; i <= b.i1; i++) {
        const t = (body[i] as ParaNode).text.replace(/\s+$/, "");
        const m = /^(.*?)(?:\t|\.{2,}|…|\s{2,})+\s*([0-9]+|[ivxlcdm]+)$/i.exec(t) ?? /^(.*?\D)(\d{1,3})$/.exec(t);
        if (m) out.push({ label: m[1].trim(), key: keyOf(m[1]), page: m[2], used: false });
        else if (t.trim()) out.push({ label: t.trim(), key: keyOf(t), page: "", used: false });
    }
    return out;
}

/** Heading yang dikumpulkan Word untuk TOC \o "1-3" (+ TC \f c), urut dokumen. */
function collectHeadings(body: Body): { text: string; level: number }[] {
    const out: { text: string; level: number }[] = [];
    for (const n of body) {
        if (n.kind !== "paragraph") continue;
        const tc = [...instrOf(n.xml).matchAll(/\bTC\s+"([^"]*)"[^"]*?\\f\s*c(?![a-z])(?:[^"]*?\\l\s*(\d))?/g)];
        if (tc.length) {
            for (const m of tc) out.push({ text: m[1].replace(/\s+/g, " ").trim(), level: Number(m[2] ?? 1) });
            continue;
        }
        const pPr = pPrOf(n.xml);
        const st = /<w:pStyle w:val="Heading([1-3])"/.exec(pPr)?.[1];
        const ol = /<w:outlineLvl w:val="([0-2])"/.exec(pPr)?.[1];
        const level = st ? Number(st) : ol != null ? Number(ol) + 1 : 0;
        const text = n.text.replace(/\s+/g, " ").trim();
        if (level && text) out.push({ text, level });
    }
    return out;
}

/** Isi Daftar Isi yang tersimpan disusun ulang dari heading yang benar-benar ada. */
function rebuildContents(body: Body, regions: VerifyRegion[], startId: number): number {
    const b = locateBy(body, /^\s*TOC\b[\s\S]*(?:\\o|\\f\s*c(?![a-z]))/);
    if (!b) return 0;
    const heads = collectHeadings(body);
    if (!heads.length) return 0;

    const olds = oldEntries(body, b);
    const norm = (t: string) => t.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    const items = heads.map((h) => {
        const k = norm(h.text);
        const exact = olds.find((o) => !o.used && norm(o.label) === k);
        const loose = exact ?? olds.find((o) => {
            const ok = norm(o.label);
            return !o.used && ok.length >= 5 && (k.startsWith(ok) || ok.startsWith(k));
        });
        if (loose) loose.used = true;
        return { ...h, page: loose?.page ?? "" };
    });

    const last = body[b.i1] as ParaNode;
    const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(pPrOf(last.xml))?.[0] ?? "";
    const xml = items.map((e, idx) => {
        const isFirst = idx === 0;
        const isLast = idx === items.length - 1;
        let inner =
            `<w:r><w:t xml:space="preserve">${escapeXml(e.text)}</w:t></w:r><w:r><w:tab/></w:r>` +
            (e.page ? `<w:r><w:t>${escapeXml(e.page)}</w:t></w:r>` : "");
        if (isFirst) inner = b.open + inner;
        if (isLast) inner += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
        return `<w:p><w:pPr><w:pStyle w:val="TOC${Math.min(3, Math.max(1, e.level))}"/>${isLast ? sect : ""}</w:pPr>${inner}</w:p>`;
    }).join("");

    const fresh = parseBlocks(xml, startId) as ParaNode[];
    const removed = body.slice(b.i0, b.i1 + 1) as ParaNode[];
    regions.push({
        beforeLines: removed.map((p) => paragraphLine(p.xml)).filter(Boolean),
        afterLines: fresh.map((p) => paragraphLine(p.xml)).filter(Boolean),
        beforeText: "",
        afterText: "",
    });
    body.splice(b.i0, b.i1 - b.i0 + 1, ...fresh);
    return fresh.length;
}

export function refreshListCaches(model: DocModel, opts: { dropTc?: Set<number> } = {}) {
    const regions: VerifyRegion[] = [];
    const counts = { tables: 0, figures: 0, appendices: 0 };
    const body = model.body;

    // TC di keterangan lampiran tidak ikut Daftar Tabel (lampiran punya daftarnya sendiri).
    if (opts.dropTc?.size) {
        for (const n of body) if (n.kind === "paragraph" && opts.dropTc.has(n.id)) n.xml = stripFields(n.xml, /^\s*TC\b/i, false);
    }

    let nextId = 1_000_000 + body.reduce(
        (m, n) => (n.kind === "paragraph" || n.kind === "table" ? Math.max(m, n.id) : m), 0);
    const contents = rebuildContents(body, regions, nextId);
    nextId += contents + 10;
    // Cadangan bila Daftar Isi tidak bisa disusun ulang: minimal buang entri bagian awal.
    if (!contents) pruneFrontEntries(body, regions);
    for (const flag of ["t", "g", "l"] as const) {
        const entries = flag === "l" ? collectStyled(body, "JudulLampiran") : collectTc(body, flag);
        const b = entries.length ? locate(body, flag) : null;
        if (!b) continue;

        const olds = oldEntries(body, b);

        const items = entries.map((text) => {
            const k = keyOf(text);
            let best: (typeof olds)[number] | null = null;
            for (const o of olds) {
                if (o.used || o.key.length < 5 || k.length < 5) continue;
                if (!(k.startsWith(o.key) || o.key.startsWith(k))) continue;
                if (!best || o.key.length > best.key.length) best = o;
            }
            if (best) best.used = true;
            return { text, page: best?.page ?? "" };
        });

        const first = body[b.i0] as ParaNode;
        const last = body[b.i1] as ParaNode;
        const noSect = (p: string) => p.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, "");
        const firstPPr = noSect(pPrOf(first.xml));
        const midPPr = firstPPr.replace(/<w:pageBreakBefore(?:\s[^>]*)?\/>/, "");
        const lastPPr = /<w:sectPr\b/.test(pPrOf(last.xml)) ? pPrOf(last.xml) : midPPr;
        const rPr = (() => {
            for (const m of runsOf(first.xml + last.xml)) {
                if (/<w:t[\s>]/.test(m[0]) && !/fldChar|instrText/.test(m[0]))
                    return (/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m[0])?.[0] ?? "")
                        .replace(/<w:rStyle\b[^>]*\/>/, "").replace(/<w:color\b[^>]*\/>/, "").replace(/<w:u\b[^>]*\/>/, "");
            }
            return "";
        })();
        const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;

        const label = flag === "t" ? "Tabel" : flag === "g" ? "Gambar" : null;
        const open = b.open.replace(
            /(<w:instrText[^>]*>)([\s\S]*?)(<\/w:instrText>)/g,
            (_all, a: string, t: string, c: string) =>
                a + (label
                    ? t.replace(/\\f\s*[tg](?![a-z])/, `\\c "${label}"`)
                    : t.replace(/\\f\s*l(?![a-z])|\\t\s*"Judul Lampiran[^"]*"/, `\\c "Lampiran"`)) + c,
        );

        const xml = items.map((e, idx) => {
            const isFirst = idx === 0;
            const isLast = idx === items.length - 1;
            const pPr = isLast ? (isFirst && !/<w:sectPr\b/.test(lastPPr) ? firstPPr : lastPPr) : isFirst ? firstPPr : midPPr;
            let inner = run(`<w:t xml:space="preserve">${escapeXml(e.text)}</w:t>`) + run("<w:tab/>") + (e.page ? run(`<w:t>${e.page}</w:t>`) : "");
            if (isFirst) inner = open + inner;
            if (isLast) inner += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
            return `<w:p>${entryPPr(pPr)}${inner}</w:p>`;
        }).join("");

        const fresh = parseBlocks(xml, nextId) as ParaNode[];
        nextId += fresh.length + 10;
        const removed = body.slice(b.i0, b.i1 + 1) as ParaNode[];
        regions.push({
            beforeLines: removed.map((p) => paragraphLine(p.xml)).filter(Boolean),
            afterLines: fresh.map((p) => paragraphLine(p.xml)).filter(Boolean),
            beforeText: "",
            afterText: "",
        });
        body.splice(b.i0, b.i1 - b.i0 + 1, ...fresh);
        if (flag === "t") counts.tables = items.length;
        else if (flag === "g") counts.figures = items.length;
        else counts.appendices = items.length;
    }
    // TC sudah tidak dibutuhkan (daftar tabel/gambar memakai \c, lampiran memakai \t).
    // Dibiarkan, LibreOffice/WPS memasukkan semuanya ke Daftar Isi.
    for (const n of body) {
        if (n.kind === "paragraph" && n.xml.includes("TC ")) {
            n.xml = stripFields(n.xml, /^\s*TC\b[\s\S]*\\f\s*[tgl](?![a-z])/i, false);
        }
    }
    return { ...counts, regions };
}

/**
 * Judul tiap lampiran ("Lampiran 1. Kuesioner…") tidak boleh ber-outline level,
 * supaya Daftar Isi hanya memuat "LAMPIRAN". Isinya tetap masuk Daftar Lampiran (TC \f l).
 */
/** Bagian awal sebelum Kata Pengantar — tidak dicantumkan di Daftar Isi. */
const FRONT_SKIP_RE = /\b(halaman\s+(judul|sampul)|sampul|persetujuan|pengesahan|pernyataan|motto|moto|persembahan|berita\s+acara)\b/i;

const APPX_NUM_RE = /^\s*(lampiran)\s+([0-9]{1,2}|[A-Za-z]|[IVXLCivxlc]{1,6})(?![A-Za-z0-9])/i;
/** "LAMPIRAN 4  JADWAL PENELITIAN" / "Lampiran 1." → label + nomor (+ pemisah) di awal. */
const APPX_HEAD_RE = /^[ \t\u00a0]*(lampiran|appendix)[ \t\u00a0]+([0-9]{1,2}|[A-Za-z]|[IVXLCivxlc]{1,6})(?![A-Za-z0-9])[ \t\u00a0]*[.:\-\u2013\u2014]?[ \t\u00a0]*/i;
const SOFT_BREAK_RE = /<w:br(?![^>]*w:type="page")[^>]*\/>|<w:cr\s*\/>/;

const toRoman = (n: number) =>
    [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]].reduce((acc, [v, s]) => {
        while (n >= (v as number)) { acc += s; n -= v as number; }
        return acc;
    }, "");
function fromRoman(s: string): number {
    const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
    const a = s.toLowerCase();
    let t = 0;
    for (let k = 0; k < a.length; k++) { const c = v[a[k]] ?? 0; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
    return t;
}

const seqField = (rPr: string, instr: string, result: string) =>
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
    (result ? `<w:r>${rPr}<w:t>${escapeXml(result)}</w:t></w:r>` : "") +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;

const tcRun = (text: string) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> TC "${escapeXml(text)}" \\f c \\l 1 </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

type NumFmt = "ARABIC" | "ROMAN" | "ALPHABETIC";

/**
 * Nomor lampiran berurutan 1..n walau ditulis campur ("1, 2, III, 4, 5, VI") →
 * format mayoritas. null bila ada yang tanpa nomor atau urutannya tidak 1..n.
 */
function appendixFormat(tokens: (string | null)[]): { fmt: NumFmt; upper: boolean } | null {
    if (!tokens.length || tokens.some((t) => t === null)) return null;
    const kinds: NumFmt[] = [];
    for (let i = 0; i < tokens.length; i++) {
        const v = tokens[i]!;
        const want = i + 1;
        if (/^\d+$/.test(v)) { if (Number(v) !== want) return null; kinds.push("ARABIC"); continue; }
        if (/^[ivxlc]+$/i.test(v) && fromRoman(v) === want) { kinds.push("ROMAN"); continue; }
        if (v.length === 1 && v.toUpperCase().charCodeAt(0) - 64 === want) { kinds.push("ALPHABETIC"); continue; }
        return null;
    }
    const count = (k: NumFmt) => kinds.filter((x) => x === k).length;
    const fmt = (["ARABIC", "ROMAN", "ALPHABETIC"] as NumFmt[]).reduce((a, b) => (count(b) > count(a) ? b : a));
    const letters = (tokens as string[]).filter((t) => /[A-Za-z]/.test(t));
    return { fmt, upper: letters.every((t) => t === t.toUpperCase()) };
}

const fmtValue = (n: number, f: { fmt: NumFmt; upper: boolean }) =>
    f.fmt === "ARABIC" ? String(n)
        : f.fmt === "ROMAN" ? (f.upper ? toRoman(n) : toRoman(n).toLowerCase())
            : String.fromCharCode((f.upper ? 64 : 96) + n);

/**
 * Judul lampiran diberi field SEQ Lampiran supaya Daftar Lampiran (TOC \c "Lampiran")
 * bekerja di Word versi bahasa apa pun dan di LibreOffice/WPS.
 * Nomor berurutan (campuran 1/III/C pun) → diseragamkan ke format mayoritas lewat SEQ.
 * Selain itu → SEQ tersembunyi (\h) disisipkan, teks tidak diubah.
 */
function numberAppendixHeadings(paras: ParaNode[]) {
    if (!paras.length) return;
    const spec = appendixFormat(paras.map((p) => APPX_NUM_RE.exec(p.text)?.[2] ?? null));

    paras.forEach((p, i) => {
        let xml = stripFields(p.xml, /^\s*SEQ\s+Lampiran\b/i, true);
        const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(xml)?.[1] ?? "";
        const insertAt = (x: string) => { const e = x.indexOf("</w:pPr>"); return e !== -1 ? e + 8 : x.indexOf(">") + 1; };

        const m = spec ? APPX_NUM_RE.exec(leadingRaw(xml, 40)) : null;
        if (spec && m) {
            const value = fmtValue(i + 1, spec);
            const sw = spec.fmt === "ARABIC" ? "ARABIC" : spec.upper ? spec.fmt : spec.fmt.toLowerCase();
            xml = removeLeading(xml, m[0].length);
            const label = m[0].slice(0, m[0].length - m[2].length);
            const at = insertAt(xml);
            xml = xml.slice(0, at) +
                `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r>` +
                seqField(rPr, `SEQ Lampiran \\* ${sw}`, value) +
                xml.slice(at);
            p.text = p.text.replace(APPX_NUM_RE, (all: string, _l: string, tok: string) => all.slice(0, all.length - tok.length) + value);
        } else {
            const at = insertAt(xml);
            xml = xml.slice(0, at) + seqField(rPr, "SEQ Lampiran \\h", "") + xml.slice(at);
        }
        p.xml = xml;
        p.hasField = true;
    });
}

/** Paragraf pendek berbentuk judul ("Kisi-Kisi dan Contoh Butir Tes"), bukan kalimat/tabel/daftar. */
function titleLike(q: ParaNode): boolean {
    const t = q.text.replace(/\s+/g, " ").trim();
    return t.length >= 3 && t.length <= 150 && !q.hasDrawing && !q.hasNumbering && !q.hasField &&
        !q.hasSectPr && !q.hasPageBreak && !/[.!?;,:]$/.test(t) && t.split(" ").length <= 14 &&
        !/^[a-z]/.test(t) && !/^(tabel|table|gambar|figure|lampiran|appendix)\b/i.test(t) &&
        !/\b(berikut|adalah|merupakan)\b/i.test(t);
}

/**
 * Judul lampiran diseragamkan jadi "LAMPIRAN n" ↵ "JUDUL" dalam SATU paragraf
 * (Daftar Lampiran memuat nomor + judulnya):
 *  - "LAMPIRAN 4  JADWAL PENELITIAN" → ganti baris setelah nomor
 *  - "LAMPIRAN 1" + paragraf judul di bawahnya → judul digabung dengan ganti baris
 * Mengembalikan baris paragraf yang ikut digabung (untuk verifikasi), atau null bila tidak berubah.
 */
function shapeAppendixHeading(p: ParaNode, body?: Body): string[] | null {
    const m = APPX_HEAD_RE.exec(leadingRaw(p.xml, 80));
    if (!m) return null;
    const tm = APPX_HEAD_RE.exec(p.text);
    const rest = tm ? p.text.slice(tm[0].length).replace(/\s+/g, " ").trim() : "";
    const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p.xml)?.[1] ?? "";

    if (rest) {
        if (SOFT_BREAK_RE.test(p.xml)) return null;
        let xml = removeLeading(p.xml, m[0].length);
        const e = xml.indexOf("</w:pPr>");
        const at = e !== -1 ? e + 8 : xml.indexOf(">") + 1;
        xml = xml.slice(0, at) +
            `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(`${m[1]} ${m[2]} `)}</w:t><w:br/></w:r>` + xml.slice(at);
        p.xml = xml;
        p.text = `${m[1]} ${m[2]} ${rest}`;
        return [];
    }

    if (!body) return null;
    const i = body.indexOf(p);
    if (i === -1) return null;
    let k = i + 1;
    while (k < body.length) {
        const b = body[k];
        if (b.kind === "paragraph" && !b.text.trim() && !b.hasDrawing && !b.hasSectPr && !b.hasPageBreak && !b.hasField) { k++; continue; }
        break;
    }
    const q = body[k];
    if (!q || q.kind !== "paragraph" || !titleLike(q)) return null;

    const caps = /<w:caps\/>/.test(rPr);
    const qs = q.xml.indexOf("</w:pPr>");
    const content = q.xml
        .slice(qs !== -1 ? qs + 8 : q.xml.indexOf(">") + 1, q.xml.length - "</w:p>".length)
        .replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) =>
            withProps(run, "w:rPr", { "w:b": "<w:b/>", "w:bCs": "<w:bCs/>", ...(caps ? { "w:caps": "<w:caps/>" } : {}) }, RPR_ORDER));
    p.xml = p.xml.replace(/<\/w:p>$/, `<w:r>${rPr}<w:t xml:space="preserve"> </w:t><w:br/></w:r>${content}</w:p>`);
    p.text = `${p.text.trim()} ${q.text.replace(/\s+/g, " ").trim()}`;
    const line = paragraphLine(q.xml);
    body.splice(k, 1);
    return [line];
}

/**
 * Judul yang tidak boleh masuk Daftar Isi diberi style tanpa outline level:
 *  - judul tiap lampiran → "Judul Lampiran" (tetap masuk Daftar Lampiran)
 *  - halaman judul, persetujuan, pengesahan, pernyataan, motto, persembahan → "Judul Bagian Awal"
 * Bila lampiran tidak punya halaman pembatas "LAMPIRAN", lampiran pertama diberi
 * entri TC "LAMPIRAN" supaya Daftar Isi tetap memuat LAMPIRAN.
 */
export function demoteFromToc(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    body?: Body,
): { appendix: number; front: number; shaped: number; regions: VerifyRegion[] } {
    const appx: ParaNode[] = [];
    const out = { appendix: 0, front: 0, shaped: 0, regions: [] as VerifyRegion[] };
    let hasDivider = false;
    const demote = (n: ParaNode, styleId: string) => {
        const before = n.xml;
        n.xml = withProps(
            n.xml.replace(/<w:outlineLvl\b[^>]*\/>/g, ""),
            "w:pPr",
            { "w:pStyle": `<w:pStyle w:val="${styleId}"/>` },
            PPR_ORDER,
        );
        return n.xml !== before;
    };
    nodes.forEach((n, i) => {
        if (n.kind !== "paragraph") return;
        const role = outline.items[i]?.role;
        if (role === "appendix_heading" && /^\s*lampiran\s*$/i.test(n.text)) { hasDivider = true; return; }
        if (role === "appendix_heading" && !/^\s*(daftar\s+)?lampiran\s*$/i.test(n.text)) {
            if (demote(n, "JudulLampiran")) out.appendix++;
            appx.push(n);
        } else if (role === "front_heading" && FRONT_SKIP_RE.test(n.text)) {
            if (demote(n, "JudulAwal")) out.front++;
        }
    });

    const before = new Map<ParaNode, string[]>(appx.map((p) => [p, [paragraphLine(p.xml)]]));
    for (const p of appx) {
        const merged = shapeAppendixHeading(p, body);
        if (merged) before.get(p)!.push(...merged);
    }
    numberAppendixHeadings(appx);
    if (appx.length && !hasDivider) appx[0].xml = appx[0].xml.replace(/<\/w:p>$/, `${tcRun("LAMPIRAN")}</w:p>`);

    for (const p of appx) {
        const b = before.get(p)!;
        const a = paragraphLine(p.xml);
        if (b.length === 1 && b[0] === a) continue;
        out.regions.push({ beforeLines: b.filter(Boolean), afterLines: [a].filter(Boolean), beforeText: "", afterText: "" });
        out.shaped++;
    }
    return out;
}

/** Entri bagian awal di isi Daftar Isi yang tersimpan juga dibuang (sebelum baris BAB pertama). */
function pruneFrontEntries(body: Body, regions: VerifyRegion[]): number {
    const b = locateBy(body, /^\s*TOC\b[\s\S]*\\o/);
    if (!b) return 0;
    const last = body[b.i1] as ParaNode;
    const drop = new Set<number>();
    for (let i = b.i0; i <= b.i1; i++) {
        const t = (body[i] as ParaNode).text;
        if (/^\s*BAB\b/i.test(t)) break;
        const label = t.split("\t")[0].trim();
        if (!label || !FRONT_SKIP_RE.test(label)) continue;
        if (i === b.i1 && /<w:sectPr\b/.test(last.xml)) continue;
        drop.add(i);
    }
    if (!drop.size) return 0;
    const keep: number[] = [];
    for (let i = b.i0; i <= b.i1; i++) if (!drop.has(i)) keep.push(i);
    if (!keep.length) return 0;

    // Awal/akhir field TOC dipindah ke entri pertama/terakhir yang tersisa.
    if (drop.has(b.i0)) {
        const p = body[keep[0]] as ParaNode;
        const at = p.xml.indexOf("</w:pPr>");
        const pos = at !== -1 ? at + 8 : p.xml.indexOf(">") + 1;
        p.xml = p.xml.slice(0, pos) + b.open + p.xml.slice(pos);
    }
    if (drop.has(b.i1)) {
        const p = body[keep[keep.length - 1]] as ParaNode;
        p.xml = p.xml.replace(/<\/w:p>$/, `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`);
    }
    regions.push({
        beforeLines: [...drop].map((i) => paragraphLine((body[i] as ParaNode).xml)).filter(Boolean),
        afterLines: [],
        beforeText: "",
        afterText: "",
    });
    for (const i of [...drop].sort((a, c) => c - a)) body.splice(i, 1);
    return drop.size;
}

/**
 * Style "Table of Figures" — dipakai Word untuk TOC \c (daftar tabel/gambar).
 * Hanging indent di sini tidak mengganggu Daftar Isi (TOC 1–3).
 */
export async function patchListEntryStyle(zip: JSZip): Promise<void> {
    const f = zip.file("word/styles.xml");
    if (!f) return;
    const xml = await f.async("string");
    const toc1 = /<w:style\b[^>]*w:styleId="TOC1"[^>]*>[\s\S]*?<\/w:style>/.exec(xml)?.[0] ?? "";
    const pos = /<w:tab\b[^>]*w:val="right"[^>]*w:pos="(\d+)"/.exec(toc1)?.[1]
        ?? /<w:tab\b[^>]*w:pos="(\d+)"[^>]*w:val="right"/.exec(toc1)?.[1] ?? "7937";
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(toc1)?.[0] ?? "";
    const style =
        `<w:style w:type="paragraph" w:styleId="TableofFigures"><w:name w:val="table of figures"/>` +
        `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>` +
        `<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${pos}"/></w:tabs>` +
        `<w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/>` +
        `<w:ind w:left="${HANG}" w:right="0" w:hanging="${HANG}"/><w:jc w:val="left"/></w:pPr>` +
        `${rPr}</w:style>`;
    const re = /<w:style\b[^>]*w:styleId="TableofFigures"[^>]*>[\s\S]*?<\/w:style>/;
    const out = re.test(xml) ? xml.replace(re, style) : xml.replace(/<\/w:styles>/, `${style}</w:styles>`);
    let final = out;
    const h1 = /<w:style\b[^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/.exec(final)?.[0];
    if (h1) {
        for (const [id, name] of [["JudulLampiran", "Judul Lampiran"], ["JudulAwal", "Judul Bagian Awal"]] as const) {
            const clone = h1
                .replace(/w:styleId="Heading1"/, `w:styleId="${id}"`)
                .replace(/<w:name w:val="[^"]*"\/>/, `<w:name w:val="${name}"/>`)
                .replace(/<w:link\b[^>]*\/>/, "")
                .replace(/<w:outlineLvl\b[^>]*\/>/g, "")
                .replace(/<\/w:pPr>/, `<w:outlineLvl w:val="9"/></w:pPr>`);
            const re = new RegExp(String.raw`<w:style\b[^>]*w:styleId="${id}"[^>]*>[\s\S]*?<\/w:style>`);
            final = re.test(final) ? final.replace(re, clone) : final.replace(/<\/w:styles>/, `${clone}</w:styles>`);
        }
    }
    if (final !== xml) zip.file("word/styles.xml", final);
}
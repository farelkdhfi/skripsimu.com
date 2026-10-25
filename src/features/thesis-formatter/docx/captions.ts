/**
 * Keterangan tabel & gambar ala "Insert Caption" Word:
 *  - keterangan tanpa nomor yang menempel pada objek dikenali
 *  - Title Case HANYA untuk kalimat judulnya; penjelasan & "Sumber: …" tetap kalimat biasa
 *  - keterangan gambar panjang: judul tetap di keterangan, penjelasan/sumber jadi baris sendiri
 *    di bawahnya (Daftar Gambar hanya memuat judul)
 *  - tabel → keterangan di atas, gambar → di bawah
 *  - penomoran ulang per bab: "<Label> <bab>.{SEQ}" + style Caption
 *  - rujukan di teks ("pada Gambar 4.5") mengikuti nomor baru
 *  - lampiran: nomor asli dipertahankan, tidak masuk daftar
 *  - bab yang keterangannya tidak semua dikenali tidak dinomori ulang (cegah nomor ganda)
 */
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import { parseBlocks, textOf, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { paragraphLine, type VerifyRegion } from "./verify";
import type { Outline, OutlineItem } from "../structure/outline";
import { captionKind } from "../structure/roles";
import { leadingRaw, removeLeading } from "./lists";
import { retext, titleCaseCaption } from "./tables";

export type CaptionKind = "table" | "figure";

type Spec = {
    role: "table_caption" | "figure_caption";
    label: string;
    words: string;
    refWords: string;
    noun: string;
    /** Keterangan berada di bawah objek (gambar). */
    below: boolean;
};

const SPEC: Record<CaptionKind, Spec> = {
    table: {
        role: "table_caption", label: "Tabel", noun: "tabel", below: false,
        words: String.raw`tabel|table|tbl\.?`,
        refWords: String.raw`Tabel|TABEL|Table|Tbl\.?`,
    },
    figure: {
        role: "figure_caption", label: "Gambar", noun: "gambar", below: true,
        words: String.raw`gambar|gbr\.?|figure|fig\.?|ilustrasi`,
        refWords: String.raw`Gambar|GAMBAR|Figure|Gbr\.?|Fig\.?|Ilustrasi|ILUSTRASI`,
    },
};

/** Nomor keterangan; spasi di sekitar pemisah ditoleransi ("3. 1" = "3.1"). */
const NUM = String.raw`(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[ \t\u00a0]?[.\-][ \t\u00a0]?\d+)*`;
const SP = String.raw`[ \t\u00a0]*`;
/** Pemisah setelah nomor: "Tabel 2.1. …", "Tabel 2.1: …", "Tabel 2.1 – …". */
const SEP = String.raw`[.:\u2013\u2014-]`;

const patterns = (s: Spec) => ({
    starts: new RegExp(String.raw`^\s*(?:${s.words})(?![a-z])`, "i"),
    full: new RegExp(String.raw`^${SP}(?:${s.words})${SP}[.:]?${SP}(?:${NUM})?${SP}${SEP}?${SP}`, "i"),
    text: new RegExp(String.raw`^\s*(?:${s.words})\s*[.:]?\s*(?:${NUM})?\s*${SEP}?\s*`, "i"),
    oldNum: new RegExp(String.raw`^\s*(?:${s.words})\s*\.?\s*(\d+(?:\s?[.\-]\s?\d+)*)`, "i"),
    /** Hanya label berhuruf kapital: "t tabel 1,985" (statistik) tidak ikut. */
    ref: new RegExp(String.raw`\b(${s.refWords})(\s+)(\d+(?:[.\-]\d+)*)(?![.,]?\d)`, "g"),
    refStrip: new RegExp(String.raw`\b(?:${s.words})\s+\d+(?:[.\-]\d+)*`, "gi"),
    unnumbered: new RegExp(String.raw`^(?:${s.words})\s*[.:]?\s+(?!${NUM})\S.{2,150}$`, "i"),
});
const RX = { table: patterns(SPEC.table), figure: patterns(SPEC.figure) };

const NOT_CAPTION_RE = /\b(berikut|di\s*atas|di\s*bawah|menunjukkan|menyajikan|memperlihatkan|menampilkan|menggambarkan|disajikan|ditampilkan|terlihat|merupakan|adalah|tersebut)\b/i;
const STOP = new Set(["dan", "atau", "yang", "pada", "untuk", "dengan", "dari", "tiap", "per", "the", "and", "of"]);
/** Kata yang menunjuk objek SEBELUM kalimat ("Berdasarkan Tabel 2.1", "tabel di atas"). */
const BACK_CUE = /\b(berdasarkan|di\s*atas|tersebut|sebelumnya|telah\s+disajikan|terlihat\s+pada)\b/i;
/** Kata yang menunjuk objek SESUDAH kalimat ("pada Tabel 2.1 berikut"). */
const FWD_CUE = /\b(berikut|di\s*bawah\s*ini)\b/i;

type AnyNode = DocModel["body"][number];

function isObject(s: Spec, n: { kind: string }): boolean {
    return s.below ? n.kind === "paragraph" && (n as ParaNode).hasDrawing : n.kind === "table";
}

const decodeXml = (s: string) =>
    s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const fullText = (xml: string) =>
    [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXml(m[1])).join("");
const words = (t: string) => ` ${t.toLowerCase().replace(/[^a-z0-9\u00c0-\u024f]+/g, " ").trim()} `;
const keywords = (desc: string) =>
    words(desc).trim().split(" ").filter((w) => w.length > 2 && !STOP.has(w)).slice(0, 3);
/** Kalimat menyebut isi objeknya ("Ringkasan model pada Tabel 4.11"). */
function mentions(text: string, desc: string): boolean {
    const kw = keywords(desc);
    if (kw.length < 2) return false;
    const hay = words(text);
    return kw.every((w) => hay.includes(` ${w} `));
}

function fromRoman(s: string): number {
    const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
    const a = s.toLowerCase();
    let t = 0;
    for (let k = 0; k < a.length; k++) { const c = v[a[k]]; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
    return t;
}

/* ------------------------------------------------------------------ */
/* Judul vs penjelasan di dalam satu keterangan                         */
/* ------------------------------------------------------------------ */

const NOTE_LEAD_RE = /(^|[\s.;,])(sumber|source|keterangan|catatan|note)\s*:/i;
/** Token sebelum titik yang BUKAN akhir kalimat (singkatan, inisial, nomor). */
const ABBR_RE = /^(?:dkk|al|dll|dsb|dst|no|nos|vol|hal|fig|gbr|tab|prof|dr|ir|drs|dra|st|vs|etc|jl|kab|kec|pt|cv|rp|[A-Za-z])$/i;

/**
 * Posisi (pada teks lengkap) awal bagian penjelasan keterangan: kalimat kedua,
 * atau "Sumber:/Keterangan:/Catatan:". null bila keterangan hanya berisi judul.
 */
function captionCut(full: string, labelLen: number): number | null {
    const body = full.slice(labelLen);
    const cuts: number[] = [];
    const note = NOTE_LEAD_RE.exec(body);
    if (note && note.index + note[1].length > 0) cuts.push(note.index + note[1].length);
    for (const m of body.matchAll(/(\S+)\.\s+(?=[A-Z(])/g)) {
        const tok = m[1].replace(/^[(["'\u201c]+/, "");
        if (ABBR_RE.test(tok) || /\.\S/.test(tok) || /^\d+$/.test(tok)) continue;
        cuts.push(m.index! + m[0].length);
        break;
    }
    if (!cuts.length) return null;
    const cut = Math.min(...cuts);
    if (body.slice(0, cut).trim().split(/\s+/).length < 2 || !body.slice(cut).trim()) return null;
    return labelLen + cut;
}

/**
 * Pecah paragraf pada posisi teks `offset` (gabungan w:t). Hasil: XML paragraf kepala
 * (pPr asli) + isi ekor (tanpa pPr). null bila titik pecah berada di dalam
 * hyperlink/field/revisi (tidak aman dipecah).
 */
function splitParagraph(xml: string, offset: number): { head: string; tail: string } | null {
    const open = /^<w:p\b[^>]*>/.exec(xml)?.[0];
    if (!open || !xml.endsWith("</w:p>")) return null;
    let start = open.length;
    const pPr = /^<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml.slice(start))?.[0] ?? "";
    start += pPr.length;
    const content = xml.slice(start, xml.length - "</w:p>".length);
    const balanced = (s: string) =>
        ["w:hyperlink", "w:fldSimple", "w:ins", "w:del", "w:smartTag", "w:sdt", "w:customXml"].every((t) =>
            (s.match(new RegExp(`<${t}\\b(?![^>]*\\/>)`, "g")) ?? []).length === (s.match(new RegExp(`</${t}>`, "g")) ?? []).length) &&
        (s.match(/w:fldCharType="begin"/g) ?? []).length === (s.match(/w:fldCharType="end"/g) ?? []).length;

    let pos = 0;
    for (const m of content.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)) {
        const t = [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((x) => decodeXml(x[1])).join("");
        const s = pos;
        pos += t.length;
        if (pos <= offset) continue;
        const at = m.index!;
        let headC: string;
        let tailC: string;
        if (s >= offset) {
            headC = content.slice(0, at);
            tailC = content.slice(at);
        } else {
            const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m[0])?.[0] ?? "";
            const k = offset - s;
            const mk = (x: string) => (x ? `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(x)}</w:t></w:r>` : "");
            headC = content.slice(0, at) + mk(t.slice(0, k).replace(/\s+$/, ""));
            tailC = mk(t.slice(k).replace(/^\s+/, "")) + content.slice(at + m[0].length);
        }
        if (!balanced(headC)) return null;
        return { head: open + pPr + headC + "</w:p>", tail: tailC };
    }
    return null;
}

/* ------------------------------------------------------------------ */
/* Bab                                                                  */
/* ------------------------------------------------------------------ */

/** Nomor bab untuk tiap node di zona isi (null di luar bab). */
function chapterMap(nodes: (ParaNode | TableNode)[], outline: Outline): Map<number, number | null> {
    const out = new Map<number, number | null>();
    let ch: number | null = null;
    let idx = 0;
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (!it) return;
        if (it.zone === "main" && n.kind === "paragraph" &&
            (it.role === "chapter_label" || (it.role === "chapter" && !it.continuesChapterLabel))) {
            idx++;
            const m = /bab\s+([ivxlcdm]+|\d+)/i.exec(n.text);
            ch = m ? (/^\d+$/.test(m[1]) ? Number(m[1]) : fromRoman(m[1])) || idx : idx;
        }
        out.set(n.id, it.zone === "main" ? ch : null);
    });
    return out;
}

/**
 * Lepas field yang instruksinya cocok. keepResult=true: hasil field jadi teks
 * biasa (SEQ lama); false: field dibuang seluruhnya (TC).
 */
export function stripFields(xml: string, instr: RegExp, keepResult: boolean): string {
    const hit = (s: string) => instr.test(decodeXml(s));
    xml = xml.replace(/<w:fldSimple\b[^>]*?\/>/g, (all) => (hit(/w:instr="([^"]*)"/.exec(all)?.[1] ?? "") ? "" : all));
    xml = xml.replace(/<w:fldSimple\b([^>]*[^/])>([\s\S]*?)<\/w:fldSimple>/g, (all, attrs: string, inner: string) =>
        hit(/w:instr="([^"]*)"/.exec(attrs)?.[1] ?? "") ? (keepResult ? inner : "") : all);

    const runs = [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
    type F = { runs: number[]; instr: string; stage: "instr" | "result" };
    const stack: F[] = [];
    const drop = new Set<number>();
    runs.forEach((m, k) => {
        const r = m[0];
        const ch = /<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"/.exec(r)?.[1];
        if (ch === "begin") { stack.push({ runs: [k], instr: "", stage: "instr" }); return; }
        const top = stack[stack.length - 1];
        if (!top) return;
        if (ch === "separate") { top.runs.push(k); top.stage = "result"; return; }
        if (ch === "end") {
            top.runs.push(k);
            stack.pop();
            if (hit(top.instr)) top.runs.forEach((x) => drop.add(x));
            return;
        }
        if (top.stage === "instr") {
            top.instr += /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(r)?.[1] ?? "";
            top.runs.push(k);
        } else if (!keepResult) top.runs.push(k);
    });
    if (!drop.size) return xml;
    let out = "";
    let last = 0;
    runs.forEach((m, k) => {
        if (!drop.has(k)) return;
        out += xml.slice(last, m.index);
        last = m.index! + m[0].length;
    });
    return out + xml.slice(last);
}

/** Ganti potongan teks (posisi pada gabungan w:t) tanpa merusak run/format. */
export function editText(xml: string, edits: { start: number; end: number; text: string }[]): string {
    const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    const toks = [...xml.matchAll(re)].map((m) => decodeXml(m[2]));
    let off = 0;
    const spans = toks.map((t) => { const s = off; off += t.length; return { s, e: off }; });
    const chars = toks.map((t) => t.split(""));
    for (const ed of edits) {
        spans.forEach((sp, ti) => {
            for (let c = Math.max(ed.start, sp.s); c < Math.min(ed.end, sp.e); c++) chars[ti][c - sp.s] = "";
        });
        const ti = spans.findIndex((sp) => ed.start >= sp.s && ed.start < sp.e);
        if (ti !== -1) chars[ti][ed.start - spans[ti].s] = ed.text;
    }
    let k = 0;
    return xml.replace(re, () => `<w:t xml:space="preserve">${escapeXml(chars[k++].join(""))}</w:t>`);
}

/** Keterangan tidak pernah tebal: buang w:b/w:bCs dan style karakter tebal, lalu kunci b=0. */
function unbold(xml: string): string {
    const off = `<w:b w:val="0"/><w:bCs w:val="0"/>`;
    return xml.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) => {
        if (!/<w:t[\s>]/.test(run)) return run;
        let r = run
            .replace(/<w:bCs(?:\s[^>]*)?\/>/g, "")
            .replace(/<w:b(?:\s[^>]*)?\/>/g, "")
            .replace(/<w:rStyle\s+w:val="(?:Strong|[^"]*Bold[^"]*|Judul[^"]*)"\s*\/>/gi, "");
        if (/<w:rPr>/.test(r)) {
            r = r.replace(/<w:rPr>((?:<w:rStyle\b[^>]*\/>)?(?:<w:rFonts\b[^>]*\/>)?)/, `<w:rPr>$1${off}`);
        } else {
            r = r.replace(/^<w:r(\s[^>]*)?>/, (m) => `${m}<w:rPr>${off}</w:rPr>`);
        }
        return r;
    });
}

/**
 * Bangun ulang keterangan + style Caption:
 *  - per bab   : "<Label> <bab>." + {SEQ <Label> \r per bab} + pemisah + judul  ("Tabel 2.1 Judul")
 *  - berurutan : "<Label> " + {SEQ <Label>} + pemisah + judul                    ("Tabel 3. Judul")
 */
function rebuildCaption(p: ParaNode, s: Spec, chapter: number, k: number, o: { global: boolean; sep: string }): boolean {
    const r = RX[s.role === "table_caption" ? "table" : "figure"];
    const before = paragraphLine(p.xml);
    let xml = stripFields(stripFields(p.xml, /^\s*SEQ\b/i, true), /^\s*TC\b/i, false);
    const lm = r.full.exec(leadingRaw(xml, 40));
    if (!lm) return false;
    xml = removeLeading(xml, lm[0].length);
    // Buang run yang teksnya benar-benar kosong setelah label lama dihapus (spasi tetap dipertahankan).
    xml = xml.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) => {
        if (!/<w:t[\s>]/.test(run)) return run;
        const rest = run
            .replace(/^<w:r(?:\s[^>]*)?>|<\/w:r>$/g, "")
            .replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "")
            .replace(/<w:t(?:\s[^>]*)?><\/w:t>|<w:t\s*\/>/g, "");
        return rest === "" ? "" : run;
    });

    const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(xml)?.[1] ?? "";
    const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;
    const head = o.global ? `${s.label} ` : `${s.label} ${chapter}.`;
    const prefix =
        run(`<w:t xml:space="preserve">${escapeXml(head)}</w:t>`) +
        run(`<w:fldChar w:fldCharType="begin"/>`) +
        run(`<w:instrText xml:space="preserve"> SEQ ${s.label} \\* ARABIC${!o.global && k === 1 ? " \\r 1" : ""} </w:instrText>`) +
        run(`<w:fldChar w:fldCharType="separate"/>`) +
        run(`<w:t>${k}</w:t>`) +
        run(`<w:fldChar w:fldCharType="end"/>`) +
        run(`<w:t xml:space="preserve">${escapeXml(o.sep)}</w:t>`);

    const end = xml.indexOf("</w:pPr>");
    const at = end !== -1 ? end + 8 : xml.indexOf(">") + 1;
    xml = xml.slice(0, at) + prefix + xml.slice(at);
    p.xml = withProps(xml, "w:pPr", { "w:pStyle": `<w:pStyle w:val="Caption"/>` }, PPR_ORDER);
    p.text = `${head}${k}${o.sep}${p.text.replace(r.text, "")}`;
    p.hasField = true;
    return paragraphLine(p.xml) !== before;
}

/* ------------------------------------------------------------------ */
/* Keterangan tanpa nomor                                               */
/* ------------------------------------------------------------------ */

/**
 * "Tabel Sebaran populasi…" / "Gambar. Normal P-P Plot…" tanpa nomor baru
 * dianggap keterangan bila menempel pada objeknya dan objek itu belum punya
 * keterangan. Jalankan sebelum detectTableNotes/applyFormat.
 */
export function promoteUnnumberedCaptions(nodes: (ParaNode | TableNode)[], outline: Outline): Set<number> {
    const out = new Set<number>();
    const blank = (k: number) => { const n = nodes[k]; return n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing; };
    const near = (i: number, step: 1 | -1) => {
        for (let k = i + step; k >= 0 && k < nodes.length; k += step) if (!blank(k)) return k;
        return -1;
    };

    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || n.hasDrawing || !it || it.role !== "body") return;
        if (it.zone !== "main" && it.zone !== "back") return;
        const t = n.text.replace(/\s+/g, " ").trim().replace(/[.:]$/, "");
        if (t.length > 160) return;

        for (const kind of ["table", "figure"] as const) {
            const s = SPEC[kind];
            const r = RX[kind];
            if (!r.unnumbered.test(t)) continue;
            const rest = t.replace(r.text, "");
            if (NOT_CAPTION_RE.test(rest) || /[.!?;]\s/.test(rest) || rest.split(" ").length > 16) return;
            const isObj = (k: number) => k !== -1 && isObject(s, nodes[k]);
            const isCap = (k: number) => k !== -1 && outline.items[k]?.role === s.role;
            // Keterangan di atas objek X yang tepat berada di bawah objek lain adalah milik objek lain itu
            // (gambar bertumpuk: [gbr A] [ket A] [gbr B] [ket B]); begitu pula sebaliknya.
            const ownsAbove = (k: number) => { const c = near(k, -1); return isCap(c) && !isObj(near(c, -1)); };
            const ownsBelow = (k: number) => { const c = near(k, 1); return isCap(c) && !isObj(near(c, 1)); };
            const prev = near(i, -1);
            const next = near(i, 1);
            const above = isObj(next) && !ownsBelow(next);
            const below = isObj(prev) && !ownsAbove(prev);
            if (!above && !below) return;
            (it as { role: string }).role = s.role;
            out.add(n.id);
            return;
        }
    });
    return out;
}

/* ------------------------------------------------------------------ */
/* Penataan + penomoran + rujukan                                       */
/* ------------------------------------------------------------------ */

export type CaptionOptions = {
    /** "chapter" = Tabel 2.1, "global" = Tabel 3 (berurutan seluruh naskah). */
    numbering?: "chapter" | "global";
    /** Teks setelah nomor: " " (Tabel 2.1 Judul) atau ". " (Tabel 3. Judul). */
    separator?: string;
    /** Label keterangan; null = ikuti dokumen (gambar: "Ilustrasi" bila itu yang dominan). */
    label?: string | null;
};

/** Label keterangan gambar yang dipakai dokumen ("Ilustrasi" vs "Gambar"). */
export function figureLabelOf(nodes: (ParaNode | TableNode)[], outline: Outline): string {
    let ilu = 0;
    let gam = 0;
    nodes.forEach((n, i) => {
        if (n.kind !== "paragraph" || outline.items[i]?.role !== "figure_caption") return;
        if (/^\s*ilustrasi\b/i.test(n.text)) ilu++;
        else if (/^\s*(gambar|gbr)\b/i.test(n.text)) gam++;
    });
    return ilu > gam ? "Ilustrasi" : "Gambar";
}

export type CaptionReport = {
    kind: CaptionKind;
    /** Label yang dipakai ("Tabel", "Gambar", "Ilustrasi") — juga nama SEQ-nya. */
    label: string;
    titled: number;
    moved: number;
    numbered: number;
    refsUpdated: number;
    /** Keterangan gambar yang penjelasan/sumbernya dipisah ke baris sendiri. */
    split: number;
    /** Rujukan ke nomor yang objeknya tidak ada ("Gambar 2.1"). */
    unresolved: string[];
    uncaptioned: string[];
    /** Bab yang tidak dinomori ulang karena ada keterangan yang tidak dikenali. */
    skipped: number[];
    appendix: Set<number>;
    regions: VerifyRegion[];
};

export function arrangeCaptions(
    model: DocModel,
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    kind: CaptionKind,
    opts: CaptionOptions = {},
): CaptionReport {
    const base = SPEC[kind];
    const r = RX[kind];
    const label = opts.label ?? (kind === "figure" ? figureLabelOf(nodes, outline) : base.label);
    const s: Spec = { ...base, label };
    const global = opts.numbering === "global";
    const sep = opts.separator === "period" ? ". "
        : opts.separator === "space" || opts.separator == null ? " "
            : opts.separator;
    const rep: CaptionReport = {
        kind, label, titled: 0, moved: 0, numbered: 0, refsUpdated: 0, split: 0,
        unresolved: [], uncaptioned: [], skipped: [], appendix: new Set(), regions: [],
    };
    const body = model.body;
    const itemOf = new Map<number, OutlineItem>();
    nodes.forEach((n, i) => { if (outline.items[i]) itemOf.set(n.id, outline.items[i]); });
    const idOf = (n: AnyNode) => (n.kind === "paragraph" || n.kind === "table" ? (n as ParaNode | TableNode).id : -1);
    const roleOf = (n: AnyNode) => (n.kind === "paragraph" ? itemOf.get((n as ParaNode).id)?.role : undefined);
    const isEmpty = (n: AnyNode) =>
        n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing && !n.hasSectPr && !n.hasPageBreak;
    const nextContent = (from: number, step: 1 | -1) => {
        for (let k = from + step; k >= 0 && k < body.length; k += step) if (!isEmpty(body[k])) return k;
        return -1;
    };
    const isObj = (k: number) => k !== -1 && isObject(s, body[k]);
    const isCap = (k: number) => k !== -1 && roleOf(body[k]) === s.role;
    /** Sisi tempat keterangan seharusnya: tabel di atas (-1), gambar di bawah (+1). */
    const dir: 1 | -1 = s.below ? 1 : -1;

    // 1. Title Case untuk KALIMAT JUDUL saja (baris asli dicatat untuk verifikasi).
    const titleOnly = (t: string) => {
        const lab = r.text.exec(t)?.[0].length ?? 0;
        const cut = captionCut(t, lab);
        return cut === null ? titleCaseCaption(t) : titleCaseCaption(t.slice(0, cut)) + t.slice(cut);
    };
    const orig = new Map<ParaNode, { line: string; text: string }>();
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || it?.role !== s.role || it.zone === "front" || it.zone === "cover") return;
        if (!r.starts.test(n.text)) return;
        orig.set(n, { line: paragraphLine(n.xml), text: n.text });
        const xml = retext(n.xml, titleOnly);
        if (xml !== n.xml) { n.xml = xml; n.text = textOf(xml); rep.titled++; }
    });

    // 2. Keterangan di sisi yang salah → pindah ke sisi yang benar.
    const moved = new Set<ParaNode>();
    for (let oi = 0; oi < body.length; oi++) {
        if (!isObj(oi) || isCap(nextContent(oi, dir))) continue;
        const wrong = nextContent(oi, -dir as 1 | -1);
        if (!isCap(wrong)) continue;
        // Keterangan itu milik objek di sisi lainnya.
        if (isObj(nextContent(wrong, -dir as 1 | -1))) continue;
        const cap = body[wrong] as ParaNode;
        body.splice(wrong, 1);
        body.splice(oi, 0, cap);
        if (dir === -1) oi++;
        moved.add(cap);
        rep.moved++;
    }

    // 3. Objek di bab yang tetap tanpa keterangan → dilaporkan (judul tidak dikarang).
    const chapterOf = chapterMap(nodes, outline);
    const objNo = new Map<number, number>();
    body.forEach((n, bi) => {
        if (!isObject(s, n)) return;
        const id = idOf(n);
        if (itemOf.get(id)?.zone !== "main") return;
        if (s.below && (n as ParaNode).text.trim()) return;
        const c = chapterOf.get(id) ?? 0;
        const no = (objNo.get(c) ?? 0) + 1;
        objNo.set(c, no);
        const side = nextContent(bi, dir);
        // Sudah berketerangan, atau bagian dari tabel bersambung / gambar bertumpuk.
        if (isCap(side) || isObj(side)) return;
        rep.uncaptioned.push(c ? `Bab ${c}, ${s.noun} ke-${no}` : `${s.noun} ke-${no} sebelum BAB I`);
    });

    // Bab yang masih punya keterangan bernomor yang TIDAK dikenali sebagai keterangan
    // tidak dinomori ulang: nomor & rujukannya dibiarkan agar tidak muncul nomor ganda.
    const unsafe = new Set<number>();
    body.forEach((n) => {
        if (n.kind !== "paragraph" || roleOf(n) === s.role) return;
        const t = n.text.trim();
        if (!t || t.length > 700 || captionKind(t) !== s.role || !r.oldNum.test(t)) return;
        const c = chapterOf.get(n.id);
        if (c != null) unsafe.add(c);
    });
    rep.skipped = [...unsafe].sort((a, b) => a - b);

    // 4. Penomoran ulang per bab + caption gaya Word (style Caption + SEQ).
    type Entry = { old: string | null; implied: string | null; next: string; pos: number; obj: number; chapter: number; desc: string };
    const mapping: Entry[] = [];
    const counters = new Map<number, number>();
    const renumbered = new Set<ParaNode>();
    const splits: { p: ParaNode; cut: number }[] = [];
    body.forEach((n, bi) => {
        if (n.kind !== "paragraph" || roleOf(n) !== s.role || !r.starts.test(n.text)) return;
        n.xml = unbold(n.xml);
        const c = chapterOf.get(n.id);
        if (c == null) {
            // Lampiran: nomor asli (L.1, L3.1, A-1, …) dipertahankan, hanya style diseragamkan.
            if (itemOf.get(n.id)?.zone === "back") {
                n.xml = withProps(n.xml, "w:pPr", { "w:pStyle": `<w:pStyle w:val="Caption"/>` }, PPR_ORDER);
                rep.appendix.add(n.id);
            }
            return;
        }
        // Berurutan: satu bab bermasalah menghentikan penomoran ulang seluruh naskah (cegah nomor ganda).
        if (global ? unsafe.size > 0 : unsafe.has(c)) return;
        const om = r.oldNum.exec(n.text);
        if (!om) {
            const o = nextContent(bi, -dir as 1 | -1);
            const drawing = o !== -1 && body[o].kind === "paragraph" && (body[o] as ParaNode).hasDrawing;
            if (!isObj(o) && !(kind === "table" && drawing)) return;
        }
        const key = global ? 0 : c;
        const k = (counters.get(key) ?? 0) + 1;
        counters.set(key, k);
        const o = nextContent(bi, -dir as 1 | -1);
        mapping.push({
            old: om ? om[1].replace(/\s+/g, "").replace(/-/g, ".") : null, implied: null,
            next: global ? String(k) : `${c}.${k}`, pos: bi, obj: isObj(o) ? o : bi, chapter: c, desc: n.text.replace(r.text, ""),
        });
        if (rebuildCaption(n, s, c, k, { global, sep })) renumbered.add(n);
        rep.numbered++;
        // Keterangan gambar berisi penjelasan / sumber → dipisah setelah rujukan selesai.
        if (s.below) {
            const full = fullText(n.xml);
            const cut = captionCut(full, r.text.exec(full)?.[0].length ?? 0);
            if (cut !== null) splits.push({ p: n, cut });
        }
    });

    // Keterangan yang dulu tanpa nomor mengisi celah penomoran lama
    // (4.4 → [tanpa nomor] → 4.6 berarti dulu dirujuk sebagai 4.5).
    const usedOld = new Set(mapping.map((m) => m.old).filter(Boolean) as string[]);
    const bump = (o: string, d: number) => {
        const parts = o.split(".");
        const last = Number(parts.pop()) + d;
        return Number.isFinite(last) && last > 0 ? [...parts, String(last)].join(".") : null;
    };
    mapping.forEach((m, i) => {
        if (m.old) return;
        const prev = mapping.slice(0, i).reverse().find((x) => x.chapter === m.chapter && x.old);
        const next = mapping.slice(i + 1).find((x) => x.chapter === m.chapter && x.old);
        const guess = prev ? bump(prev.old!, 1) : next ? bump(next.old!, -1) : null;
        if (guess && !usedOld.has(guess)) m.implied = guess;
    });

    const resolve = (num: string, bi: number, text: string, ch: number | null, ctx: string): Entry | null => {
        const exact = mapping.filter((x) => x.old === num);
        if (exact.length === 1) return exact[0];
        if (exact.length > 1) {
            // Nomor ganda di dokumen asli: utamakan yang isinya disebut di kalimat, lalu
            // arah kalimat ("Berdasarkan Tabel …" → objek sebelumnya; "… berikut" → sesudahnya),
            // lalu objek yang paling dekat dengan kalimat.
            const named = exact.filter((x) => mentions(text, x.desc));
            const pool = named.length ? named : exact;
            const before = pool.filter((x) => x.obj < bi).sort((a, b) => b.obj - a.obj)[0];
            const after = pool.filter((x) => x.obj > bi).sort((a, b) => a.obj - b.obj)[0];
            if (BACK_CUE.test(ctx) && before) return before;
            if (FWD_CUE.test(ctx) && after) return after;
            const near = [before, after].filter((x): x is Entry => Boolean(x));
            if (!near.length) return pool[0];
            return near.sort((a, b) => Math.abs(a.obj - bi) - Math.abs(b.obj - bi))[0];
        }
        const implied = mapping.find((x) => x.implied === num);
        if (implied) return implied;
        const named = mapping.filter((x) => (ch == null || x.chapter === ch) && mentions(text, x.desc));
        return named.length === 1 ? named[0] : null;
    };

    // 5. Rujukan di teks disesuaikan dengan nomor baru.
    const unresolved = new Set<string>();
    body.forEach((n, bi) => {
        if (n.kind !== "paragraph") return;
        const role = roleOf(n);
        if (role === "table_caption" || role === "figure_caption") return;
        const zone = itemOf.get(n.id)?.zone;
        if (zone !== "main" && zone !== "back") return;
        const ch = chapterOf.get(n.id) ?? null;
        if (ch != null && (global ? unsafe.size > 0 : unsafe.has(ch))) return;
        const full = fullText(n.xml);
        r.ref.lastIndex = 0;
        if (!r.ref.test(full)) return;
        r.ref.lastIndex = 0;
        const edits: { start: number; end: number; text: string }[] = [];
        const repl: (string | null)[] = [];
        for (const m of full.matchAll(r.ref)) {
            const num = m[3].replace(/-/g, ".");
            const ctx = full.slice(Math.max(0, m.index! - 60), m.index! + m[0].length + 60);
            const hit = resolve(num, bi, full, ch, ctx);
            if (!hit) { repl.push(null); if (num.includes(".")) unresolved.add(`${s.label} ${num}`); continue; }
            if (hit.next === m[3]) { repl.push(null); continue; }
            repl.push(hit.next);
            const start = m.index! + m[1].length + m[2].length;
            edits.push({ start, end: start + m[3].length, text: hit.next });
        }
        if (!edits.length) return;
        const beforeLine = paragraphLine(n.xml);
        n.xml = editText(n.xml, edits);
        let k = 0;
        n.text = n.text.replace(r.ref, (all, w: string, sp: string) => {
            const v = repl[k++];
            return v ? `${w}${sp}${v}` : all;
        });
        rep.regions.push({
            beforeLines: [beforeLine].filter(Boolean),
            afterLines: [paragraphLine(n.xml)].filter(Boolean),
            beforeText: full.replace(r.refStrip, " "),
            afterText: fullText(n.xml).replace(r.refStrip, " "),
        });
        rep.refsUpdated += edits.length;
    });
    rep.unresolved = [...unresolved];

    // 5b. Penjelasan / "Sumber: …" keterangan gambar → baris sendiri tepat di bawahnya.
    const tails = new Map<ParaNode, ParaNode>();
    let nextId = 3_000_000 + body.reduce((m, x) => (x.kind === "paragraph" || x.kind === "table" ? Math.max(m, x.id) : m), 0);
    for (const { p, cut } of splits) {
        const bi = body.indexOf(p);
        const parts = bi === -1 ? null : splitParagraph(p.xml, cut);
        if (!parts) continue;
        // Baris penjelasan/"Sumber:" mengikuti perataan keterangannya (rata kiri pada pedoman UNPAD).
        const capJc = /<w:jc w:val="(\w+)"\/>/.exec(p.xml.slice(0, Math.max(0, p.xml.indexOf("</w:pPr>"))))?.[1] ?? "center";
        const tailXml =
            `<w:p><w:pPr><w:keepNext w:val="0"/><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/>` +
            `<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="${capJc === "both" ? "left" : capJc}"/></w:pPr>${parts.tail}</w:p>`;
        const tail = (parseBlocks(tailXml, nextId++) as ParaNode[])[0];
        if (!tail || tail.kind !== "paragraph" || !tail.text.trim()) continue;
        p.xml = parts.head;
        p.text = fullText(p.xml).replace(/\s+/g, " ").trim();
        body.splice(bi + 1, 0, tail);
        tails.set(p, tail);
        renumbered.add(p);
        rep.split++;
    }

    // 6. Gambar: objek diikat ke keterangan di bawahnya; keterangan tidak ikut keepNext
    //    (kecuali punya baris penjelasan, yang harus ikut satu halaman).
    if (s.below) {
        body.forEach((n, bi) => {
            if (n.kind !== "paragraph" || roleOf(n) !== s.role) return;
            const o = nextContent(bi, -1);
            if (isObj(o)) {
                const f = body[o] as ParaNode;
                f.xml = withProps(f.xml, "w:pPr", { "w:keepNext": "<w:keepNext/>" }, PPR_ORDER);
            }
            const hasTail = tails.has(n as ParaNode);
            n.xml = withProps(n.xml, "w:pPr", {
                "w:keepNext": hasTail ? "<w:keepNext/>" : `<w:keepNext w:val="0"/>`,
                "w:spacing": `<w:spacing w:before="120" w:after="${hasTail ? 0 : 240}" w:line="240" w:lineRule="auto"/>`,
            }, PPR_ORDER);
        });
    }

    // Region verifikasi untuk keterangan yang dipindah / dinomori ulang / dipecah.
    for (const [p, o] of orig) {
        if (!moved.has(p) && !renumbered.has(p)) continue;
        const tail = tails.get(p);
        rep.regions.push({
            beforeLines: [o.line].filter(Boolean),
            afterLines: [paragraphLine(p.xml), tail ? paragraphLine(tail.xml) : ""].filter(Boolean),
            beforeText: o.text.replace(r.text, ""),
            afterText: fullText(p.xml).replace(r.text, "") + (tail ? ` ${fullText(tail.xml)}` : ""),
        });
    }

    return rep;
}

/* ------------------------------------------------------------------ */
/* Baris lanjutan keterangan tabel sejajar awal judul                   */
/* ------------------------------------------------------------------ */

/** Lebar glyph (per 1000 em) untuk menghitung lebar label "Tabel 4.12 ". */
const GLYPH: Record<string, Record<string, number>> = {
    "Times New Roman": { d: 500, ".": 250, " ": 250, "-": 333, T: 611, a: 444, b: 500, e: 444, l: 278, L: 611 },
    Arial: { d: 556, ".": 278, " ": 278, "-": 333, T: 611, a: 556, b: 556, e: 556, l: 222, L: 556 },
    Calibri: { d: 507, ".": 252, " ": 226, "-": 306, T: 487, a: 479, b: 525, e: 498, l: 229, L: 420 },
};

function labelTwips(label: string, font: string, halfPts: number): number {
    const w = GLYPH[font] ?? GLYPH["Times New Roman"];
    let units = 0;
    for (const ch of label) units += /\d/.test(ch) ? w.d : w[ch] ?? 500;
    return Math.ceil((units / 1000) * halfPts * 10);
}

/** Keterangan tabel (rata kiri): baris kedua dst. sejajar huruf pertama judul. */
export function hangTableCaptions(model: DocModel, font: string): number {
    let count = 0;
    for (const n of model.body) {
        if (n.kind !== "paragraph" || !/<w:pStyle w:val="Caption"\/>/.test(n.xml)) continue;
        if (/<w:jc w:val="(center|right|end)"/.test(n.xml)) continue;
        const m = /^\s*((?:tabel|table|gambar|ilustrasi|figure)\s+[A-Za-z]{0,2}[.\-]?\d+(?:[.\-]\d+)*\.?\s+)/i.exec(n.text);
        if (!m) continue;
        const runs = n.xml.slice(Math.max(0, n.xml.indexOf("</w:pPr>")));
        const half = Number(/<w:sz w:val="(\d+)"/.exec(runs)?.[1] ?? 24);
        const h = labelTwips(m[1], font, half) + 15;
        n.xml = withProps(n.xml, "w:pPr", { "w:ind": `<w:ind w:left="${h}" w:right="0" w:hanging="${h}"/>` }, PPR_ORDER);
        count++;
    }
    return count;
}
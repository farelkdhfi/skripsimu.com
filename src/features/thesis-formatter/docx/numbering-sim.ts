/**
 * Simulasi penomoran otomatis Word: nomor yang DITAMPILKAN Word untuk tiap paragraf bernomor
 * ("2.1.4.1", "a.", "(3)"), dihitung dari numbering.xml + numPr paragraf/style, dalam urutan
 * dokumen. Dipakai untuk (1) menulis nomor otomatis menjadi teks persis seperti yang dilihat
 * pembaca, dan (2) memeriksa bahwa nomor judul yang tampil tidak berubah setelah diformat.
 *
 * Aturan yang ditiru: penghitung per abstractNum (w:num yang memakai abstractNum sama berbagi
 * penghitung kecuali ada startOverride), tingkat lebih dalam di-reset saat tingkat atasnya naik,
 * w:start, w:numFmt (decimal, huruf, romawi), w:isLgl, lvlText "%1.%2".
 */

import type { BlockNode, DocModel, ParaNode } from "./document-model";
import { paragraphLine } from "./verify";

type Lvl = { start: number; fmt: string; text: string; isLgl: boolean };
type Abs = Map<number, Lvl>;

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
const roman = (n: number) => { let s = ""; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const letter = (n: number) => { let s = ""; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };

function fmtNum(n: number, fmt: string): string {
    switch (fmt) {
        case "lowerLetter": return letter(n);
        case "upperLetter": return letter(n).toUpperCase();
        case "lowerRoman": return roman(n);
        case "upperRoman": return roman(n).toUpperCase();
        case "decimalZero": return n < 10 ? `0${n}` : String(n);
        case "bullet": case "none": return "";
        default: return String(n);
    }
}

export type NumRef = { numId: string; ilvl: number };

export type NumberingSim = {
    /** Nomor numPr efektif paragraf (langsung atau dari style), atau null. */
    refOf(xml: string, styleId: string | null): NumRef | null;
    /** Format tingkat (untuk memeriksa bentuk nomornya). */
    levelOf(ref: NumRef): Lvl | null;
    /** Panggil BERURUTAN untuk setiap paragraf dokumen: nomor tampil paragraf ini (atau null). */
    next(xml: string, styleId: string | null): { ref: NumRef; label: string } | null;
    /** abstractNum milik numId (w:num yang abstractNum-nya sama berbagi penghitung). */
    absOf(numId: string): string | null;
};

export function createNumberingSim(numXml: string, stylesXml: string): NumberingSim {
    const abs = new Map<string, Abs>();
    for (const m of numXml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
        const lv: Abs = new Map();
        for (const l of m[2].matchAll(/<w:lvl\b[^>]*w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g)) {
            lv.set(Number(l[1]), {
                start: Number(/<w:start w:val="(\d+)"/.exec(l[2])?.[1] ?? 1),
                fmt: /<w:numFmt w:val="(\w+)"/.exec(l[2])?.[1] ?? "decimal",
                text: /<w:lvlText w:val="([^"]*)"/.exec(l[2])?.[1] ?? "",
                isLgl: /<w:isLgl(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(l[2]),
            });
        }
        abs.set(m[1], lv);
    }
    const nums = new Map<string, { abs: string; overrides: Map<number, number> }>();
    for (const m of numXml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
        const a = /<w:abstractNumId w:val="(\d+)"/.exec(m[2])?.[1];
        if (!a) continue;
        const ov = new Map<number, number>();
        for (const o of m[2].matchAll(/<w:lvlOverride w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvlOverride>/g)) {
            const st = /<w:startOverride w:val="(\d+)"/.exec(o[2])?.[1];
            if (st) ov.set(Number(o[1]), Number(st));
        }
        nums.set(m[1], { abs: a, overrides: ov });
    }
    // numPr di style (rantai basedOn).
    const styles = new Map<string, { np?: string; basedOn?: string }>();
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
        if (!id || !/w:type="paragraph"/.test(m[1])) continue;
        styles.set(id, { np: /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(m[2])?.[1], basedOn: /<w:basedOn w:val="([^"]+)"/.exec(m[2])?.[1] });
    }
    const styleNp = (id: string | null): string | undefined => {
        let s = id;
        for (let d = 0; s && d < 10; d++) {
            const st = styles.get(s);
            if (!st) return undefined;
            if (st.np) return st.np;
            s = st.basedOn ?? null;
        }
        return undefined;
    };

    const refOf = (xml: string, styleId: string | null): NumRef | null => {
        const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
        const direct = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
        const sNp = styleNp(styleId);
        const numId = (direct && /<w:numId w:val="(\d+)"/.exec(direct)?.[1]) ?? (sNp && /<w:numId w:val="(\d+)"/.exec(sNp)?.[1]);
        if (!numId || numId === "0" || !nums.has(numId)) return null;
        const il = (direct && /<w:ilvl w:val="(\d+)"/.exec(direct)?.[1]) ?? (sNp && /<w:ilvl w:val="(\d+)"/.exec(sNp)?.[1]) ?? "0";
        return { numId, ilvl: Number(il) };
    };
    const levelOf = (ref: NumRef): Lvl | null => {
        const n = nums.get(ref.numId);
        return n ? abs.get(n.abs)?.get(ref.ilvl) ?? null : null;
    };

    const counters = new Map<string, number[]>();          // per abstractNum
    const seenNum = new Set<string>();
    const next = (xml: string, styleId: string | null) => {
        const ref = refOf(xml, styleId);
        if (!ref) return null;
        const n = nums.get(ref.numId)!;
        const lv = abs.get(n.abs);
        if (!lv) return null;
        let c = counters.get(n.abs);
        if (!c) { c = new Array(9).fill(0); counters.set(n.abs, c); }
        if (!seenNum.has(ref.numId)) {
            seenNum.add(ref.numId);
            for (const [il, st] of n.overrides) c[il] = st - 1;
        }
        const me = lv.get(ref.ilvl);
        if (!me) return null;
        // Tingkat atas yang belum pernah muncul dianggap sudah bernilai awal (seperti Word/LibreOffice):
        // "2.1.1" tanpa "2.1" sebelumnya → judul tingkat 2 berikutnya tampil "2.2", bukan "2.1".
        for (let k = 0; k < ref.ilvl; k++) if (c[k] === 0) c[k] = lv.get(k)?.start ?? 1;
        c[ref.ilvl] = c[ref.ilvl] === 0 ? me.start : c[ref.ilvl] + 1;
        for (let k = ref.ilvl + 1; k < 9; k++) c[k] = 0;
        const label = me.text.replace(/%(\d)/g, (_m, d: string) => {
            const k = Number(d) - 1;
            const l = lv.get(k);
            const v = c[k] === 0 ? (l?.start ?? 1) : c[k];
            return fmtNum(v, me.isLgl ? "decimal" : (l?.fmt ?? "decimal"));
        });
        return { ref, label };
    };
    return { refOf, levelOf, next, absOf: (id) => nums.get(id)?.abs ?? null };
}

export type Displayed = { label: string; numId: string; abs: string; ilvl: number };

/**
 * Nomor tampil semua paragraf bernomor di model (badan, tabel, container), dalam urutan dokumen.
 * Jalankan pada model ASLI (sebelum bagian depan diganti) supaya hasilnya = yang dilihat pembaca.
 */
export function displayedLabels(model: DocModel, numXml: string, stylesXml: string): { map: Map<ParaNode, Displayed>; sim: NumberingSim } {
    const sim = createNumberingSim(numXml, stylesXml);
    const map = new Map<ParaNode, Displayed>();
    const walk = (list: BlockNode[]) => {
        for (const n of list) {
            if (n.kind === "paragraph") {
                const r = sim.next(n.xml, n.styleId);
                if (r) map.set(n, { label: r.label, numId: r.ref.numId, abs: sim.absOf(r.ref.numId) ?? r.ref.numId, ilvl: r.ref.ilvl });
            } else if (n.kind === "container") walk(n.children);
            else if (n.kind === "table") for (const row of n.rows) for (const c of row.cells) walk(c.children);
        }
    };
    walk(model.body);
    return { map, sim };
}

/** Semua paragraf model (termasuk di tabel/container), urutan dokumen. */
export function allParagraphs(model: DocModel): ParaNode[] {
    const out: ParaNode[] = [];
    const walk = (list: BlockNode[]) => {
        for (const n of list) {
            if (n.kind === "paragraph") out.push(n);
            else if (n.kind === "container") walk(n.children);
            else if (n.kind === "table") for (const row of n.rows) for (const c of row.cells) walk(c.children);
        }
    };
    walk(model.body);
    return out;
}

/** Nomor bertingkat yang tampil di judul (mis. "2.1.4.1 Kadar Air"), untuk diperiksa sesudah format. */
export type ShownHeading = { label: string; line: string };

const DOTTED_LABEL_RE = /^\d+(?:\.\d+)+\.?$/;
const keyOf = (s: string) => s.toLowerCase().replace(/[\s ]+/g, "");

/** Judul bernomor otomatis bertingkat di dokumen (bukan entri Daftar Isi). Panggil pada model ASLI. */
export function shownHeadings(map: Map<ParaNode, Displayed>): ShownHeading[] {
    const out: ShownHeading[] = [];
    for (const [p, d] of map) {
        if (!DOTTED_LABEL_RE.test(d.label.trim())) continue;
        if (/^toc|^daftar/i.test(p.styleId ?? "") || /PAGEREF|w:anchor="_Toc|\.{6,}/.test(p.xml)) continue;
        const line = paragraphLine(p.xml);
        if (!line || line.length > 150 || /[.!?;:,]$/.test(line)) continue;
        out.push({ label: d.label.trim().replace(/\.$/, ""), line });
    }
    return out;
}

/**
 * Nomor judul yang tampil di dokumen HASIL (nomor otomatis disimulasikan ulang dengan numbering &
 * style akhir, ditambah nomor ketikan) harus sama dengan di dokumen asli. Mengembalikan judul yang
 * nomornya berubah/hilang.
 */
export function changedHeadingNumbers(before: ShownHeading[], model: DocModel, numXml: string, stylesXml: string): { label: string; line: string; now: string | null }[] {
    const sim = createNumberingSim(numXml, stylesXml);
    const full: string[] = [];
    for (const p of allParagraphs(model)) {
        const r = sim.next(p.xml, p.styleId);
        const line = paragraphLine(p.xml);
        if (line) full.push(keyOf(r?.label ? `${r.label} ${line}` : line));
    }
    const have = new Set(full);
    const out: { label: string; line: string; now: string | null }[] = [];
    for (const h of before) {
        if (have.has(keyOf(`${h.label} ${h.line}`))) continue;
        const k = keyOf(h.line);
        out.push({ ...h, now: full.find((f) => f.endsWith(k)) ?? null });
    }
    return out;
}
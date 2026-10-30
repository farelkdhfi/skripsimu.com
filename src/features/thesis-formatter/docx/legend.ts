/**
 * Keterangan rumus, keterangan tabel berbentuk "simbol = arti", dan baris hipotesis — disejajarkan.
 *
 *   Yij   =  Respon hasil pengamatan pada perlakuan uji ke-i dan
 *            ulangan ke-j                      ← baris lanjutan sejajar teks sesudah "="
 *   µ     =  Nilai tengah populasi
 *   H0 : P1 = P2 = P3 = P4  :  Semua perlakuan tidak berpengaruh nyata …
 *
 * Satu blok = satu kolom "=" (posisinya dari simbol terlebar). Simbol, rumus, dan formatnya
 * (subskrip dst.) tidak diubah; yang diganti hanya spasi/tab di sekitar pemisah menjadi tab.
 * Baris seperti ini bukan rumus (tidak ditengahkan) dan bukan alinea (tidak menjorok).
 *
 * Blok keterangan rumus dikenali hanya bila tepat di bawah rumus (Office Math atau baris rumus
 * ketikan), boleh diawali "Keterangan:" / "dengan:" / "di mana:".
 */
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import { isEmptyParagraph, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole, type Role } from "../structure/roles";
import { paragraphLine, type VerifyRegion } from "./verify";
import type { TableNotes } from "./tables";

const decode = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

type Tok = { start: number; end: number; kind: "t" | "tab" | "math"; text: string; open: string };

/** Token isi paragraf: teks, tab, dan rumus (satu token utuh). Tidak masuk ke kotak teks. */
function tokens(xml: string): Tok[] {
    const from = Math.max(0, xml.indexOf("</w:pPr>"));
    const re = /<m:oMath\b[\s\S]*?<\/m:oMath>|<w:txbxContent\b[\s\S]*?<\/w:txbxContent>|<mc:Fallback>[\s\S]*?<\/mc:Fallback>|(<w:t(?:\s[^>]*)?>)([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g;
    re.lastIndex = from;
    const out: Tok[] = [];
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
        const s = m[0];
        if (s.startsWith("<m:oMath")) out.push({ start: m.index, end: m.index + s.length, kind: "math", text: [...s.matchAll(/<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g)].map((x) => decode(x[1])).join(""), open: "" });
        else if (s.startsWith("<w:tab")) out.push({ start: m.index, end: m.index + s.length, kind: "tab", text: "\t", open: "" });
        else if (m[1]) out.push({ start: m.index, end: m.index + s.length, kind: "t", text: decode(m[2]), open: m[1] });
    }
    return out;
}

/** Teks seperti dibaca (termasuk tab & isi rumus). */
export function legendText(xml: string): string {
    return tokens(xml).map((t) => t.text).join("");
}

const T = (s: string) => `<w:t xml:space="preserve">${escapeXml(s)}</w:t>`;

/**
 * Ganti spasi/tab di sekitar pemisah ke-`nth` (di teks biasa, bukan di dalam rumus) dengan
 * "⇥pemisah⇥". Mengembalikan XML baru & teks di depan pemisah, atau null bila bentuknya tidak cocok.
 */
export function splitAtSeparator(xml: string, sep: "=" | ":", nth = 1): { xml: string; label: string } | null {
    const toks = tokens(xml);
    let seen = 0;
    let ti = -1;
    let off = -1;
    for (let k = 0; k < toks.length && ti === -1; k++) {
        if (toks[k].kind !== "t") continue;
        for (let c = 0; c < toks[k].text.length; c++) {
            if (toks[k].text[c] === sep && ++seen === nth) { ti = k; off = c; break; }
        }
    }
    if (ti === -1) return null;
    const label = (toks.slice(0, ti).map((t) => t.text).join("") + toks[ti].text.slice(0, off)).replace(/\s+$/, "").replace(/^\s+/, "");
    if (!label) return null;
    const afterText = toks[ti].text.slice(off + 1);
    if (!(afterText + toks.slice(ti + 1).map((t) => t.text).join("")).trim()) return null;

    const edits: { start: number; end: number; repl: string }[] = [];
    // Sesudah pemisah: spasi/tab pembuka dibuang (sampai ketemu isi).
    const head = toks[ti].text.slice(0, off).replace(/[  \t]+$/, "");
    let rest = afterText.replace(/^[  \t]+/, "");
    if (!rest) {
        for (let k = ti + 1; k < toks.length; k++) {
            const t = toks[k];
            if (t.kind === "math") break;
            if (t.kind === "tab") { edits.push({ start: t.start, end: t.end, repl: "" }); continue; }
            const trimmed = t.text.replace(/^[  ]+/, "");
            if (trimmed !== t.text) edits.push({ start: t.start, end: t.end, repl: T(trimmed) });
            if (trimmed) break;
        }
    }
    edits.push({ start: toks[ti].start, end: toks[ti].end, repl: (head ? T(head) : "") + `<w:tab/>${T(sep)}<w:tab/>` + (rest ? T(rest) : "") });
    // Sebelum pemisah: spasi/tab penutup label dibuang.
    if (!head) {
        for (let k = ti - 1; k >= 0; k--) {
            const t = toks[k];
            if (t.kind === "math") break;
            if (t.kind === "tab") { edits.push({ start: t.start, end: t.end, repl: "" }); continue; }
            const trimmed = t.text.replace(/[  ]+$/, "");
            if (trimmed !== t.text) edits.push({ start: t.start, end: t.end, repl: T(trimmed) });
            if (trimmed) break;
        }
    }
    let out = xml;
    for (const e of edits.sort((a, b) => b.start - a.start)) out = out.slice(0, e.start) + e.repl + out.slice(e.end);
    rest = rest || "";
    return { xml: out, label };
}

/** Lebar kira-kira teks (twip) pada ukuran huruf `halfPts`. */
function textWidth(s: string, halfPts: number): number {
    let w = 0;
    // Kira-kira lebar huruf Times New Roman (em): kapital lebar, angka 0,5, spasi & tanda baca sempit.
    for (const ch of s) {
        w += /[MW]/.test(ch) ? 0.89 : /[A-Z]/.test(ch) ? 0.67 : /[mw]/.test(ch) ? 0.72 : /[0-9]/.test(ch) ? 0.5
            : /[=+<>≤≥≠×]/.test(ch) ? 0.56 : /[\s]/.test(ch) ? 0.25 : /[.,:;il|!'()\[\]fjrt]/.test(ch) ? 0.3 : 0.47;
    }
    return Math.round(w * halfPts * 10);
}

const STEP = 142;                                   // 0,25 cm
const roundUp = (v: number) => Math.ceil(v / STEP) * STEP;

/** Label "simbol" yang wajar: pendek, ≤ 3 kata, bukan kalimat. */
const symbolish = (label: string) => label.length <= 22 && label.split(/\s+/).length <= 3 && !/[.,;]$/.test(label);
const LEGEND_HEAD_RE = /^(?:keterangan|ket\.?|dengan|di\s*mana|dimana|where|notes?)\s*:?\s*$/i;
const HYPO_RE = /^\s*H\s*[0-9aAₒ₀₁]\s*:/;

type Line = { p: ParaNode; xml: string; label: string; before: string };

/** Susun satu blok: kolom pemisah & teks sesudahnya sama untuk semua baris. */
function layBlock(lines: Line[], halfPts: number, min: number, regions: VerifyRegion[], extra: Record<string, string | null | undefined> = {}) {
    const sepPos = Math.max(min, roundUp(Math.max(...lines.map((l) => textWidth(l.label, halfPts))) + 170));
    const textPos = sepPos + roundUp(textWidth("= ", halfPts) + 60);
    lines.forEach((l, k) => {
        l.p.xml = withProps(l.xml, "w:pPr", {
            "w:tabs": `<w:tabs><w:tab w:val="left" w:pos="${sepPos}"/><w:tab w:val="left" w:pos="${textPos}"/></w:tabs>`,
            "w:ind": `<w:ind w:left="${textPos}" w:right="0" w:hanging="${textPos}"/>`,
            "w:keepLines": "<w:keepLines/>",
            ...(k < lines.length - 1 ? { "w:keepNext": "<w:keepNext/>" } : {}),
            ...extra,
        }, PPR_ORDER);
        const after = paragraphLine(l.p.xml);
        if (after !== l.before) regions.push({ beforeLines: [l.before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: "", afterText: "" });
    });
}

const wOnly = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join("").replace(/\s+/g, " ").trim();
/** Rumus tampil: Office Math tanpa teks lain (atau hanya nomor rumus), atau baris rumus ketikan. */
function isEquation(p: ParaNode): boolean {
    if (/<m:oMath\b/.test(p.xml)) {
        const v = wOnly(p.xml);
        return !v || /^\(\s*[A-Za-z]?\d+(?:[.\-]\d+)*[a-z]?\s*\)$/.test(v);
    }
    const t = p.text;
    return /=/.test(t) && !/\t/.test(t) && t.trim().length <= 160 && (t.match(/[A-Za-zÀ-ɏ]{4,}/g) ?? []).length <= 2 && !/^\s*\S{1,12}\s*=\s+[A-Z][a-z]/.test(t);
}

export type LegendReport = { blocks: number; lines: number; hypotheses: number; ids: Set<number>; regions: VerifyRegion[] };

/**
 * Keterangan rumus & hipotesis di isi bab. Dipanggil SETELAH modul format paragraf/alur (menimpa
 * indentasi & perataan baris-barisnya).
 */
export function alignLegends(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    opts: { halfPts: number; indent: number; skip: Set<number>; line: number; exact: boolean },
): LegendReport {
    // Satu blok satu jarak baris (sama dengan teks isi): naskah sering mencampur spasi 1 dan 2.
    const spacing = `<w:spacing w:before="0" w:after="0" w:line="${opts.line}" w:lineRule="${opts.exact ? "exact" : "auto"}"/>`;
    const rep: LegendReport = { blocks: 0, lines: 0, hypotheses: 0, ids: new Set(), regions: [] };
    const okPara = (k: number): ParaNode | null => {
        const n = nodes[k];
        const it = outline.items[k];
        if (!n || n.kind !== "paragraph" || !it || it.zone !== "main" || opts.skip.has(n.id) || n.hasNumbering || n.hasDrawing) return null;
        if (isHeadingRole(it.role as Role)) return null;
        return n;
    };
    const legendLine = (p: ParaNode): Line | null => {
        const before = paragraphLine(p.xml);
        const r = splitAtSeparator(p.xml, "=");
        if (!r || !symbolish(r.label)) return null;
        return { p, xml: r.xml, label: r.label, before };
    };

    for (let i = 0; i < nodes.length; i++) {
        const eq = okPara(i);
        if (!eq || !isEquation(eq)) continue;
        // Lewati baris kosong & baris rumus lanjutan.
        let j = i + 1;
        while (j < nodes.length) {
            const p = okPara(j);
            if (p && (isEmptyParagraph(p) || isEquation(p))) { j++; continue; }
            break;
        }
        let head: ParaNode | null = null;
        const hp = okPara(j);
        if (hp && LEGEND_HEAD_RE.test(hp.text.trim())) { head = hp; j++; }
        const lines: Line[] = [];
        for (; j < nodes.length; j++) {
            const p = okPara(j);
            if (!p || rep.ids.has(p.id)) break;
            const l = legendLine(p);
            if (!l) break;
            lines.push(l);
        }
        if (!lines.length || (!head && lines.length < 2)) continue;
        if (head) {
            const before = paragraphLine(head.xml);
            head.xml = withProps(head.xml, "w:pPr", {
                "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
                "w:jc": `<w:jc w:val="left"/>`,
                "w:keepNext": "<w:keepNext/>",
                "w:spacing": spacing,
            }, PPR_ORDER);
            rep.ids.add(head.id);
            void before;
        }
        layBlock(lines, opts.halfPts, opts.indent, rep.regions, { "w:jc": `<w:jc w:val="both"/>`, "w:spacing": spacing });
        for (const l of lines) rep.ids.add(l.p.id);
        rep.blocks++;
        rep.lines += lines.length;
        i = j - 1;
    }

    // Daftar definisi sesudah kalimat pengantar ("Perlakuan yang digunakan …, yaitu:" lalu
    // "P1 = Tanaman hanjeli 40% …"): ≥ 2 baris "simbol = arti" berurutan, artinya berupa kata
    // (bukan rumus). Kalimat pengantarnya tidak diubah.
    for (let i = 0; i < nodes.length; i++) {
        const lead = okPara(i);
        if (!lead || rep.ids.has(lead.id) || !/:\s*$/.test(lead.text.trim()) || isEquation(lead)) continue;
        let j = i + 1;
        while (j < nodes.length) { const p = okPara(j); if (p && isEmptyParagraph(p)) { j++; continue; } break; }
        const lines: Line[] = [];
        for (; j < nodes.length; j++) {
            const p = okPara(j);
            if (!p || rep.ids.has(p.id) || isEquation(p)) break;
            const l = legendLine(p);
            if (!l) break;
            const meaning = legendText(l.xml).split(/\t=\t/)[1] ?? "";
            if (!/[A-Za-z\u00C0-\u024F]{4,}/.test(meaning)) break;
            lines.push(l);
        }
        if (lines.length < 2) continue;
        layBlock(lines, opts.halfPts, opts.indent, rep.regions, { "w:jc": `<w:jc w:val="both"/>`, "w:spacing": spacing });
        for (const l of lines) rep.ids.add(l.p.id);
        rep.blocks++;
        rep.lines += lines.length;
        i = j - 1;
    }

    // Hipotesis: "H0 : P1 = P2 = … : Semua perlakuan …" → teks sesudah titik dua kedua sejajar.
    for (let i = 0; i < nodes.length; i++) {
        const group: Line[] = [];
        for (let j = i; j < nodes.length; j++) {
            const p = okPara(j);
            if (!p || rep.ids.has(p.id) || !HYPO_RE.test(p.text)) break;
            const before = paragraphLine(p.xml);
            const r = splitAtSeparator(p.xml, ":", 2);
            if (!r || r.label.length > 40) break;
            group.push({ p, xml: r.xml, label: r.label, before });
        }
        if (!group.length) continue;
        layBlock(group, opts.halfPts, opts.indent, rep.regions, { "w:jc": `<w:jc w:val="both"/>`, "w:spacing": spacing });
        for (const l of group) rep.ids.add(l.p.id);
        rep.hypotheses += group.length;
        i += group.length - 1;
    }
    return rep;
}

/** Catatan tabel "simbol = arti" (≥ 2 baris berurutan): kolom "=" sejajar, ukuran & jarak catatan tetap. */
export function alignNoteLegends(model: DocModel, tableNotes: TableNotes, noteHalf: number): { lines: number; regions: VerifyRegion[] } {
    const out = { lines: 0, regions: [] as VerifyRegion[] };
    const run: Line[] = [];
    const flush = () => {
        if (run.length >= 2) {
            layBlock(run, noteHalf, 0, out.regions);
            out.lines += run.length;
        }
        run.length = 0;
    };
    for (const n of model.body) {
        if (n.kind === "raw") continue;
        if (n.kind !== "paragraph" || !tableNotes.has(n.id)) { flush(); continue; }
        const before = paragraphLine(n.xml);
        const r = splitAtSeparator(n.xml, "=");
        if (!r || !symbolish(r.label)) { flush(); continue; }
        run.push({ p: n, xml: r.xml, label: r.label, before });
    }
    flush();
    return out;
}
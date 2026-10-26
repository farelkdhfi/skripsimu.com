/**
 * Pengukuran tinggi halaman dengan style FINAL dokumen, dipakai untuk:
 *  1. sampul: logo diperkecil bila isi sampul meluber (fitCoverPages)
 *  2. tanda tangan/materai melayang: baris nama di bawahnya diberi jarak agar tidak tertimpa
 *     (clearFloatingSignatures)
 *  3. lembar pernyataan: spasi diturunkan bertahap bila tidak muat satu halaman (fitStatementPage)
 *
 * Halaman sampul harus muat satu halaman. Bila isi satu halaman sampul diperkirakan meluber
 * (mis. tahun terdorong ke halaman baru), gambar terbesar di halaman itu (biasanya logo)
 * diperkecil secukupnya — proporsional, tidak pernah diperbesar, paling kecil MIN_SCALE dari
 * ukuran aslinya. Teks sampul tidak diubah.
 *
 * Tinggi halaman diperkirakan dari style FINAL dokumen (setelah setelan dasar diubah):
 * ukuran & jenis huruf tiap run (termasuk tanda paragraf pada baris kosong), spasi baris
 * (auto/exactly/at least) dan spasi sebelum/sesudah dari rantai style, indentasi, pemotongan
 * baris berdasarkan perkiraan lebar huruf, ganti baris manual, gambar inline, gambar floating
 * yang mendorong teks (wrap atas-bawah/persegi), tabel, serta grid baris dokumen.
 * Formatter tidak bisa melihat pemotongan halaman Word secara pasti, jadi dipakai cadangan.
 */
import { cellParagraphs, flatten, type BlockNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { PageInfo } from "./page-openers";
import { startsNewPage, type Zone } from "../structure/roles";
import { STATEMENT_RE } from "./statement";

const MIN_SCALE = 0.6;
/** Cadangan agar perkiraan tidak meleset tipis (± setengah baris 12 pt). */
const SAFETY = 160;
const EMU_PER_TWIP = 635;

/* ------------------------------------------------------------------ */
/* Style final                                                          */
/* ------------------------------------------------------------------ */

type Spacing = { before?: number; after?: number; line?: number; rule?: string; autoBefore?: boolean; autoAfter?: boolean };
type Ind = { left?: number; right?: number; firstLine?: number; hanging?: number };
type StyleRec = { basedOn: string | null; sz?: number; font?: string; spacing: Spacing; ind: Ind; grid?: boolean };

const num = (tag: string, name: string) => {
    const v = new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(tag)?.[1];
    return v === undefined ? undefined : Number(v);
};
const onOff = (tag: string, name: string) => {
    const v = new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1];
    return v === undefined ? undefined : !/^(0|false|off)$/.test(v);
};
const spacingOf = (pPr: string): Spacing => {
    const t = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    return {
        before: num(t, "before"), after: num(t, "after"), line: num(t, "line"),
        rule: /\bw:lineRule="(\w+)"/.exec(t)?.[1], autoBefore: onOff(t, "beforeAutospacing"), autoAfter: onOff(t, "afterAutospacing"),
    };
};
const indOf = (pPr: string): Ind => {
    const t = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    return {
        left: num(t, "left") ?? num(t, "start"), right: num(t, "right") ?? num(t, "end"),
        firstLine: num(t, "firstLine"), hanging: num(t, "hanging"),
    };
};
const szOf = (rPr: string) => num(/<w:sz\b[^>]*\/>/.exec(rPr)?.[0] ?? "", "val");
const fontOf = (rPr: string) => /<w:rFonts\b[^>]*\bw:ascii="([^"]+)"/.exec(rPr)?.[1] ?? /<w:rFonts\b[^>]*\bw:hAnsi="([^"]+)"/.exec(rPr)?.[1];
const pPrOf = (xml: string) => (/<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(xml)?.[1] ?? "").replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, "");
const rPrIn = (xml: string) => /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(xml)?.[1] ?? "";
const defined = <T extends object>(o: T) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>;

class Styles {
    private para = new Map<string, StyleRec>();
    private char = new Map<string, StyleRec>();
    private defaultPara: string | null = null;
    private base: StyleRec;

    constructor(xml: string) {
        const dd = /<w:docDefaults>([\s\S]*?)<\/w:docDefaults>/.exec(xml)?.[1] ?? "";
        const dr = /<w:rPrDefault>([\s\S]*?)<\/w:rPrDefault>/.exec(dd)?.[1] ?? "";
        const dp = /<w:pPrDefault>([\s\S]*?)<\/w:pPrDefault>/.exec(dd)?.[1] ?? "";
        this.base = { basedOn: null, sz: szOf(dr) ?? 20, font: fontOf(dr) ?? "Times New Roman", spacing: spacingOf(dp), ind: indOf(dp) };
        for (const m of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
            const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
            if (!id) continue;
            const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(m[2])?.[1] ?? "";
            const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[2])?.[1] ?? "";
            const rec: StyleRec = {
                basedOn: /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1] ?? null,
                sz: szOf(rPr), font: fontOf(rPr), spacing: spacingOf(pPr), ind: indOf(pPr),
                grid: onOff(/<w:snapToGrid\b[^>]*\/>/.exec(pPr)?.[0] ?? "", "val") ?? (/<w:snapToGrid\s*\/>/.test(pPr) ? true : undefined),
            };
            if (/w:type="paragraph"/.test(m[1])) {
                this.para.set(id, rec);
                if (/w:default="(?:1|true|on)"/.test(m[1])) this.defaultPara = id;
            } else if (/w:type="character"/.test(m[1])) this.char.set(id, rec);
        }
    }

    private chain(map: Map<string, StyleRec>, id: string | null | undefined): StyleRec[] {
        const out: StyleRec[] = [];
        for (let cur = id ?? null, n = 0; cur && n < 20; n++) {
            const s = map.get(cur);
            if (!s) break;
            out.unshift(s);
            cur = s.basedOn;
        }
        return out;
    }

    paragraph(styleId: string | undefined) {
        const recs = [this.base, ...this.chain(this.para, styleId ?? this.defaultPara)];
        let spacing: Spacing = {};
        let ind: Ind = {};
        let sz = this.base.sz!;
        let font = this.base.font!;
        let grid = true;
        for (const r of recs) {
            spacing = { ...spacing, ...defined(r.spacing) };
            ind = { ...ind, ...defined(r.ind) };
            if (r.sz !== undefined) sz = r.sz;
            if (r.font) font = r.font;
            if (r.grid !== undefined) grid = r.grid;
        }
        return { spacing, ind, sz, font, grid };
    }

    charStyle(id: string | undefined): { sz?: number; font?: string } {
        let sz: number | undefined;
        let font: string | undefined;
        for (const r of this.chain(this.char, id)) { if (r.sz !== undefined) sz = r.sz; if (r.font) font = r.font; }
        return { sz, font };
    }
}

/* ------------------------------------------------------------------ */
/* Ukuran huruf                                                         */
/* ------------------------------------------------------------------ */

type FontMetric = { line: number; scale: number };
/** line: tinggi baris tunggal ÷ ukuran huruf; scale: lebar relatif terhadap Times New Roman. */
function metric(font: string): FontMetric {
    const f = font.toLowerCase();
    if (/calibri|carlito/.test(f)) return { line: 1.22, scale: 0.98 };
    if (/arial|helvetica|liberation sans|arimo/.test(f)) return { line: 1.15, scale: 1.12 };
    if (/cambria|caladea/.test(f)) return { line: 1.17, scale: 1.08 };
    if (/georgia/.test(f)) return { line: 1.14, scale: 1.15 };
    if (/verdana|tahoma|segoe/.test(f)) return { line: 1.22, scale: 1.2 };
    if (/book antiqua|palatino/.test(f)) return { line: 1.17, scale: 1.08 };
    return { line: 1.15, scale: 1 };
}

/** Lebar karakter (em) Times New Roman; huruf lain diskalakan. */
function charEm(ch: string): number {
    if (ch === " ") return 0.25;
    if (/[A-Z]/.test(ch)) return /[MW]/.test(ch) ? 0.9 : /[IJ]/.test(ch) ? 0.36 : 0.68;
    if (/[a-z]/.test(ch)) return /[mw]/.test(ch) ? 0.76 : /[ijlft]/.test(ch) ? 0.29 : 0.47;
    if (/\d/.test(ch)) return 0.5;
    if (/[.,:;'!|]/.test(ch)) return 0.27;
    return 0.42;
}

/* ------------------------------------------------------------------ */
/* Pengukuran                                                           */
/* ------------------------------------------------------------------ */

type Img = { kind: "inline" | "anchor"; holder: ParaNode; index: number; cx: number; cy: number; band?: Band; area: number };
type Band = { rel: string; offset: number | null; align: string | null; distT: number; distB: number; wraps: boolean };
type Geometry = { width: number; height: number; pageH: number; marginTop: number; gridPitch: number | null };
type Token = { w: number; h: number; text: boolean; br?: boolean };

const DRAWING_RE = /<w:drawing>[\s\S]*?<\/w:drawing>/g;

function imagesOf(p: ParaNode): Img[] {
    const out: Img[] = [];
    let index = 0;
    for (const m of p.xml.matchAll(DRAWING_RE)) {
        const d = m[0];
        const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(d);
        if (!ext) { index++; continue; }
        const cx = Number(ext[1]) / EMU_PER_TWIP;
        const cy = Number(ext[2]) / EMU_PER_TWIP;
        if (d.includes("<wp:anchor")) {
            const head = /<wp:anchor\b[^>]*>/.exec(d)?.[0] ?? "";
            const pv = /<wp:positionV\b[^>]*relativeFrom="(\w+)"[^>]*>([\s\S]*?)<\/wp:positionV>/.exec(d);
            const off = pv ? /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(pv[2])?.[1] : undefined;
            out.push({
                kind: "anchor", holder: p, index, cx, cy, area: cx * cy,
                band: {
                    rel: pv?.[1] ?? "paragraph",
                    offset: off === undefined ? null : Number(off) / EMU_PER_TWIP,
                    align: pv ? /<wp:align>(\w+)<\/wp:align>/.exec(pv[2])?.[1] ?? null : null,
                    distT: Number(/\bdistT="(\d+)"/.exec(head)?.[1] ?? 0) / EMU_PER_TWIP,
                    distB: Number(/\bdistB="(\d+)"/.exec(head)?.[1] ?? 0) / EMU_PER_TWIP,
                    wraps: /<wp:wrap(?:TopAndBottom|Square|Tight|Through)\b/.test(d),
                },
            });
        } else out.push({ kind: "inline", holder: p, index, cx, cy, area: cx * cy });
        index++;
    }
    return out;
}

class Layout {
    y = 0;
    /** Batas bawah isi terakhir (teks/gambar/tabel); paragraf kosong penutup section boleh menggantung. */
    end = 0;
    bands: [number, number][] = [];
    constructor(private g: Geometry, private styles: Styles, private scaleOf: (img: Img) => number) {}

    private place(h: number) {
        for (let guard = 0; guard < 10; guard++) {
            const hit = this.bands.find(([a, b]) => this.y < b && this.y + h > a);
            if (!hit) break;
            this.y = hit[1];
        }
        this.y += h;
    }

    private lineHeight(maxSz: number, font: string, sp: Spacing, grid: boolean): number {
        const single = (maxSz / 2) * 20 * metric(font).line;
        const rule = sp.rule ?? "auto";
        let h = rule === "exact" ? (sp.line ?? 240)
            : rule === "atLeast" ? Math.max(sp.line ?? 0, single)
            : single * ((sp.line ?? 240) / 240);
        if (grid && rule === "auto" && this.g.gridPitch) h = Math.ceil(h / this.g.gridPitch - 0.01) * this.g.gridPitch;
        return h;
    }

    paragraph(p: ParaNode, width: number) {
        const pPr = pPrOf(p.xml);
        const st = this.styles.paragraph(/<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1]);
        const sp: Spacing = { ...st.spacing, ...defined(spacingOf(pPr)) };
        const ind: Ind = { ...st.ind, ...defined(indOf(pPr)) };
        const gridOn = onOff(/<w:snapToGrid\b[^>]*\/>/.exec(pPr)?.[0] ?? "", "val") ?? st.grid;
        const markSz = szOf(rPrIn(pPr)) ?? st.sz;
        const exact = sp.rule === "exact";

        const before = sp.autoBefore ? 280 : sp.before ?? 0;
        const after = sp.autoAfter ? 280 : sp.after ?? 0;
        this.y += before;
        const top = this.y;

        // Gambar floating: daerah yang tidak boleh dilewati teks.
        for (const img of imagesOf(p)) {
            if (img.kind !== "anchor" || !img.band?.wraps) continue;
            const b = img.band;
            const h = img.cy * this.scaleOf(img);
            const frame = b.rel === "page" ? this.g.pageH : this.g.height;
            const base = b.rel === "page" || b.rel === "topMargin" ? -this.g.marginTop
                : b.rel === "bottomMargin" ? this.g.height
                : b.rel === "margin" ? 0 : top;
            const y0 = b.offset !== null ? base + b.offset
                : b.align === "bottom" ? base + frame - h
                : b.align === "center" ? base + (frame - h) / 2
                : base;
            this.bands.push([y0 - b.distT, y0 + h + b.distB]);
        }

        // Token per run: kata (lebar, ukuran), gambar inline, ganti baris.
        const tokens: Token[] = [];
        let maxSz = 0;
        const body = p.xml.slice(Math.max(0, p.xml.indexOf("</w:pPr>")));
        let drawingIdx = 0;
        const imgs = imagesOf(p);
        for (const r of body.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)) {
            const run = r[0];
            const rPr = rPrIn(run);
            const cs = this.styles.charStyle(/<w:rStyle\s+w:val="([^"]+)"/.exec(rPr)?.[1]);
            const sz = szOf(rPr) ?? cs.sz ?? st.sz;
            const font = fontOf(rPr) ?? cs.font ?? st.font;
            const em = (sz / 2) * 20 * metric(font).scale * (/<w:b(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(rPr) ? 1.05 : 1);
            const caps = /<w:caps(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(rPr);
            for (const part of run.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:br\b[^>]*\/>|<w:cr\s*\/>|<w:drawing>[\s\S]*?<\/w:drawing>/g)) {
                const x = part[0];
                if (part[1] !== undefined) {
                    const text = caps ? part[1].toUpperCase() : part[1];
                    for (const w of text.split(/(\s+)/)) {
                        if (!w) continue;
                        tokens.push({ w: [...w].reduce((a, ch) => a + charEm(ch === " " ? " " : ch), 0) * em, h: sz, text: /\S/.test(w) });
                    }
                    if (/\S/.test(text)) maxSz = Math.max(maxSz, sz);
                } else if (x.startsWith("<w:tab")) tokens.push({ w: 720, h: sz, text: false });
                else if (x.startsWith("<w:br") || x.startsWith("<w:cr")) {
                    if (/w:type="page"/.test(x)) continue; // pemisah halaman ditangani di tingkat halaman
                    tokens.push({ w: 0, h: sz, text: false, br: true });
                } else {
                    const img = imgs[drawingIdx++];
                    if (img?.kind === "inline") tokens.push({ w: img.cx * this.scaleOf(img), h: -img.cy * this.scaleOf(img), text: true });
                }
            }
        }
        if (!maxSz) maxSz = markSz;

        const left = ind.left ?? 0;
        const right = ind.right ?? 0;
        const first = (ind.firstLine ?? 0) - (ind.hanging ?? 0);
        const textH = this.lineHeight(maxSz, st.font, sp, gridOn);
        const lines: number[] = [];
        let used = 0;
        let lineH = textH;
        let avail = Math.max(720, width - left - right - first);
        let any = false;
        const flush = () => { lines.push(lineH); used = 0; lineH = textH; any = false; avail = Math.max(720, width - left - right); };
        for (const t of tokens) {
            if (t.br) { flush(); continue; }
            if (!t.text && !any) continue; // spasi di awal baris tidak memakan tempat
            if (any && t.text && used + t.w > avail) flush();
            used += t.w;
            any = any || t.text;
            if (t.h < 0 && !exact) lineH = Math.max(lineH, -t.h);
        }
        if (any || !lines.length) lines.push(lineH);
        for (const h of lines) this.place(h);
        if (p.text.trim() || p.hasDrawing) this.end = this.y;
        this.y += after;
    }

    table(t: TableNode, width: number) {
        const grid = t.grid ? [...t.grid.matchAll(/<w:gridCol\s+w:w="(\d+)"/g)].map((m) => Number(m[1])) : [];
        const total = grid.reduce((a, b) => a + b, 0) || width;
        for (const row of t.rows) {
            const startY = this.y;
            let end = startY;
            let col = 0;
            for (const cell of row.cells) {
                const span = Number(/<w:gridSpan\s+w:val="(\d+)"/.exec(cell.props ?? "")?.[1] ?? 1);
                const w = grid.length ? grid.slice(col, col + span).reduce((a, b) => a + b, 0) * (width / total) : width / row.cells.length;
                col += span;
                const sub = new Layout(this.g, this.styles, this.scaleOf);
                for (const p of cellParagraphs(cell)) sub.paragraph(p, Math.max(720, w - 216));
                end = Math.max(end, startY + sub.y + 80);
            }
            const minH = Number(/<w:trHeight\b[^>]*w:val="(\d+)"/.exec(row.props ?? "")?.[1] ?? 0);
            this.y = Math.max(end, startY + minH);
        }
        this.end = this.y;
    }
}

/* ------------------------------------------------------------------ */
/* Halaman sampul                                                       */
/* ------------------------------------------------------------------ */

type CoverPage = { nodes: (ParaNode | TableNode)[]; sect: string | null };

/** Halaman-halaman sampul: dari awal dokumen sampai isi pertama di luar zona sampul. */
function coverPages(model: DocModel, info: PageInfo): CoverPage[] {
    const all = flatten(model.body);
    const pages: CoverPage[] = [];
    let cur: (ParaNode | TableNode)[] = [];
    const hasContent = (list: (ParaNode | TableNode)[]) =>
        list.some((n) => n.kind === "table" || n.text.trim() || n.hasDrawing);
    const close = (sect: string | null) => {
        if (hasContent(cur)) pages.push({ nodes: cur, sect });
        cur = [];
    };
    let i = 0;
    for (; i < all.length; i++) {
        const n = all[i];
        const z = info.get(n)?.zone;
        const content = n.kind === "table" || Boolean(n.text.trim()) || n.hasDrawing;
        if (z && z !== "cover" && content) break;
        if (n.kind === "paragraph" && /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPrOf(n.xml))) close(null);
        if (n.kind === "paragraph" && /<w:br\b[^>]*w:type="page"/.test(n.xml) && !n.text.trim() && !n.hasDrawing) {
            close(null);
            continue;
        }
        cur.push(n);
        if (n.kind === "paragraph" && n.hasSectPr) close(/<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? null);
    }
    close(null); // halaman sampul terakhir yang diakhiri pindah halaman judul berikutnya
    // Halaman tanpa sectPr sendiri memakai ukuran kertas section yang memuatnya (sectPr berikutnya).
    let next: string | null = model.bodySectPr;
    for (let j = all.length - 1; j >= 0; j--) {
        const n = all[j];
        if (n.kind === "paragraph" && n.hasSectPr) next = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? next;
        for (const p of pages) if (p.sect === null && p.nodes[p.nodes.length - 1] === n) p.sect = next;
    }
    return pages;
}

function geometry(sect: string, fallback: string | null): Geometry | null {
    const s = sect || fallback || "";
    const pgSz = /<w:pgSz\b[^>]*\/>/.exec(s)?.[0] ?? "";
    const pgMar = /<w:pgMar\b[^>]*\/>/.exec(s)?.[0] ?? "";
    const w = num(pgSz, "w");
    const h = num(pgSz, "h");
    if (!w || !h) return null;
    const top = Math.abs(num(pgMar, "top") ?? 1440);
    const bottom = Math.abs(num(pgMar, "bottom") ?? 1440);
    const left = num(pgMar, "left") ?? 1440;
    const right = num(pgMar, "right") ?? 1440;
    const grid = /<w:docGrid\b[^>]*\/>/.exec(s)?.[0] ?? "";
    const pitch = /w:type="(?:lines|linesAndChars|snapToChars)"/.test(grid) ? num(grid, "linePitch") ?? null : null;
    return { width: w - left - right, height: h - top - bottom, pageH: h, marginTop: top, gridPitch: pitch };
}

function measure(page: CoverPage, g: Geometry, styles: Styles, scaleOf: (img: Img) => number): number {
    const lay = new Layout(g, styles, scaleOf);
    for (const n of page.nodes) {
        if (n.kind === "paragraph") lay.paragraph(n, g.width);
        else lay.table(n, g.width);
    }
    return lay.end;
}

const scaleAttr = (s: string, k: number) => String(Math.max(1, Math.round(Number(s) * k)));

/** Perkecil gambar ke-index di paragraf (inline/floating, termasuk ukuran di pic:spPr). */
function scaleDrawing(p: ParaNode, index: number, k: number) {
    let i = 0;
    p.xml = p.xml.replace(DRAWING_RE, (d) => {
        if (i++ !== index) return d;
        return d
            .replace(/(<wp:extent\s+cx=")(\d+)("\s+cy=")(\d+)(")/, (_, a, cx, b, cy, c) => `${a}${scaleAttr(cx, k)}${b}${scaleAttr(cy, k)}${c}`)
            .replace(/(<a:ext\s+cx=")(\d+)("\s+cy=")(\d+)(")/g, (_, a, cx, b, cy, c) => `${a}${scaleAttr(cx, k)}${b}${scaleAttr(cy, k)}${c}`);
    });
}

export type CoverFitReport = { scaled: number; stillOverflow: number; percent: number[] };

export function fitCoverPages(model: DocModel, info: PageInfo, stylesXml: string): CoverFitReport {
    const styles = new Styles(stylesXml);
    const out: CoverFitReport = { scaled: 0, stillOverflow: 0, percent: [] };
    for (const page of coverPages(model, info)) {
        const g = geometry(page.sect ?? "", model.bodySectPr);
        if (!g) continue;
        const height = (scaleOf: (img: Img) => number) => measure(page, g, styles, scaleOf);
        if (height(() => 1) + SAFETY <= g.height) continue;

        // Gambar terbesar di halaman ini yang memakan tempat (inline, atau floating yang mendorong teks).
        const imgs = page.nodes.flatMap((n) => (n.kind === "paragraph" ? imagesOf(n) : []))
            .filter((im) => im.kind === "inline" || im.band?.wraps);
        if (!imgs.length) { out.stillOverflow++; continue; }
        const logo = imgs.reduce((a, b) => (b.area > a.area ? b : a));
        const with_ = (k: number) => height((im) => (im.holder === logo.holder && im.index === logo.index ? k : 1));
        if (with_(MIN_SCALE) + SAFETY > g.height) { out.stillOverflow++; continue; }

        let lo = MIN_SCALE;
        let hi = 1;
        for (let step = 0; step < 14; step++) {
            const mid = (lo + hi) / 2;
            if (with_(mid) + SAFETY <= g.height) lo = mid;
            else hi = mid;
        }
        const k = Math.floor(lo * 100) / 100;
        scaleDrawing(logo.holder, logo.index, k);
        out.scaled++;
        out.percent.push(Math.round(k * 100));
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Tanda tangan & materai melayang                                      */
/* ------------------------------------------------------------------ */

/** Spasi efektif paragraf (style + langsung), twips. */
function spacingEff(styles: Styles, p: ParaNode) {
    const pPr = pPrOf(p.xml);
    const st = styles.paragraph(/<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1]);
    const sp: Spacing = { ...st.spacing, ...defined(spacingOf(pPr)) };
    return {
        sp,
        sz: szOf(rPrIn(pPr)) ?? st.sz,
        before: sp.autoBefore ? 280 : sp.before ?? 0,
        after: sp.autoAfter ? 280 : sp.after ?? 0,
    };
}

/** Ganti/isi atribut w:spacing (dan buang autospacing/lines yang menimpanya). */
function setSpacing(xml: string, attrs: Record<string, number | string>): string {
    const end = xml.indexOf("</w:pPr>");
    const head = end === -1 ? "" : xml.slice(0, end);
    let tag = /<w:spacing\b[^>]*\/>/.exec(head)?.[0] ?? "<w:spacing/>";
    const original = /<w:spacing\b[^>]*\/>/.exec(head)?.[0];
    for (const [k, v] of Object.entries(attrs)) {
        tag = tag.replace(new RegExp(`\\s*w:${k}(?:Autospacing|Lines)?="[^"]*"`, "g"), "");
        tag = tag.replace(/\s*\/>$/, ` w:${k}="${v}"/>`);
    }
    if (original) return xml.replace(original, tag);
    if (end === -1) return xml.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr>${tag}</w:pPr>`);
    // w:spacing sebelum w:ind/w:jc/w:rPr agar urutan pPr sah.
    const at = head.search(/<w:(?:ind|contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b/);
    return at === -1 ? xml.slice(0, end) + tag + xml.slice(end) : xml.slice(0, at) + tag + xml.slice(at);
}

const isSignatureFloat = (img: Img) =>
    img.kind === "anchor" && Math.max(img.cx, img.cy) <= 7 * 567 && img.band !== undefined && !img.band.wraps &&
    /paragraph|line/.test(img.band.rel) && img.band.offset !== null;

function* blockLists(list: BlockNode[], info: PageInfo, zone: Zone | undefined): Generator<{ list: BlockNode[]; zone: Zone | undefined }> {
    yield { list, zone };
    for (const n of list) {
        if (n.kind === "container") yield* blockLists(n.children, info, zone);
        else if (n.kind === "table") {
            const z = info.get(n)?.zone ?? zone;
            for (const r of n.rows) for (const c of r.cells) yield* blockLists(c.children, info, z);
        }
    }
}

/**
 * Tanda tangan/materai melayang diposisikan relatif terhadap paragrafnya. Bila spasi di bawahnya
 * menyusut (spasi baris diseragamkan, baris kosong dihapus), baris nama naik dan tertimpa gambar.
 * Baris berteks pertama di bawah gambar diberi jarak sebelum secukupnya agar mulai di bawah gambar
 * (gambar boleh sedikit menumpang ±15% tingginya, ruang kosong bawaan tanda tangan).
 */
export function clearFloatingSignatures(model: DocModel, info: PageInfo, stylesXml: string): number {
    const styles = new Styles(stylesXml);
    const g = geometry(model.bodySectPr ?? "", null);
    if (!g) return 0;
    const heightOf = (p: ParaNode) => {
        const lay = new Layout(g, styles, () => 1);
        lay.paragraph(p, g.width);
        return lay.y;
    };
    let changed = 0;
    for (const { list, zone } of blockLists(model.body, info, undefined)) {
        list.forEach((n, i) => {
            if (n.kind !== "paragraph" || !n.hasDrawing) return;
            const z = info.get(n)?.zone ?? zone;
            if (z === "cover") return;
            const floats = imagesOf(n).filter((im) => isSignatureFloat(im) && !/\bbehindDoc="(?:1|true)"/.test(
                [...n.xml.matchAll(DRAWING_RE)][im.index]?.[0] ?? ""));
            if (!floats.length) return;
            const bottom = Math.max(...floats.map((im) => (im.band!.offset ?? 0) + im.cy * 0.85));
            // Dari atas paragraf jangkar sampai baris berteks berikutnya.
            let y = heightOf(n) - spacingEff(styles, n).before;
            let target: ParaNode | null = null;
            for (let j = i + 1; j < list.length; j++) {
                const m = list[j];
                if (m.kind !== "paragraph" || m.hasSectPr || m.hasPageBreak) break;
                if (m.text.trim()) { target = m; break; }
                if (/<wp:inline\b/.test(m.xml)) break;
                y += heightOf(m);
            }
            if (!target) return;
            const cur = spacingEff(styles, target).before;
            const need = Math.ceil(bottom - y);
            if (need <= cur) return;
            target.xml = setSpacing(target.xml, { before: need });
            changed++;
        });
    }
    return changed;
}

/* ------------------------------------------------------------------ */
/* Lembar pernyataan satu halaman                                       */
/* ------------------------------------------------------------------ */

const STEPS = [0.9, 0.8, 0.7, 0.6];

export type StatementFitReport = { pages: number; tightened: number; stillOverflow: number };

/**
 * Lembar pernyataan harus satu halaman. Bila meluber, spasi baris & jarak antarparagraf di
 * halaman itu diturunkan bertahap (90% … 60%) sampai muat. Ukuran huruf dan teks tidak diubah.
 */
export function fitStatementPage(model: DocModel, info: PageInfo, stylesXml: string, extra = 0): StatementFitReport {
    const styles = new Styles(stylesXml);
    const out: StatementFitReport = { pages: 0, tightened: 0, stillOverflow: 0 };
    const all = flatten(model.body);
    all.forEach((n, i) => {
        const meta = info.get(n);
        if (n.kind !== "paragraph" || meta?.role !== "front_heading" || !STATEMENT_RE.test(n.text.trim())) return;
        let s = i;
        while (s > 0) {
            const m = all[s - 1];
            if (m.kind !== "paragraph" || m.text.trim() || m.hasSectPr || m.hasPageBreak) break;
            s--;
        }
        let e = i + 1;
        let sect: string | null = null;
        for (; e < all.length; e++) {
            const m = all[e];
            const r = info.get(m)?.role;
            if (r && startsNewPage(r)) break;
            if (m.kind === "paragraph" && m.hasSectPr) { sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(m.xml)?.[0] ?? null; e++; break; }
            if (m.kind === "paragraph" && m.hasPageBreak) { e++; break; }
        }
        if (!sect) {
            for (let k = e; k < all.length; k++) {
                const m = all[k];
                if (m.kind === "paragraph" && m.hasSectPr) { sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(m.xml)?.[0] ?? null; break; }
            }
        }
        const g = geometry(sect ?? "", model.bodySectPr);
        if (!g) return;
        const page: CoverPage = { nodes: all.slice(s, e), sect };
        out.pages++;
        // extra: ruang kotak pernyataan (garis & jarak dalam) yang ditambahkan sesudahnya.
        const fits = () => measure(page, g, styles, () => 1) + SAFETY + extra <= g.height;
        if (fits()) return;

        const paras = page.nodes.filter((x): x is ParaNode => x.kind === "paragraph");
        const original = paras.map((p) => p.xml);
        const base = paras.map((p) => spacingEff(styles, p));
        let ok = false;
        for (const k of STEPS) {
            paras.forEach((p, idx) => {
                const { sp, sz, before, after } = base[idx];
                const rule = sp.rule ?? "auto";
                const line = sp.line ?? 240;
                const minExact = Math.round(sz * 10);
                const nextLine = rule === "exact" ? Math.max(minExact, Math.round(line * k))
                    : rule === "atLeast" ? Math.round(line * k)
                    : Math.max(240, Math.round(line * k));
                p.xml = setSpacing(original[idx], {
                    before: Math.round(before * k), after: Math.round(after * k), line: nextLine, lineRule: rule,
                });
            });
            if (fits()) { ok = true; break; }
        }
        out.tightened++;
        if (!ok) out.stillOverflow++;
    });
    return out;
}
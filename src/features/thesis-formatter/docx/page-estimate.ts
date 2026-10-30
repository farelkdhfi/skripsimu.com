/**
 * Perkiraan nomor halaman untuk field PAGEREF (Daftar Isi/Tabel/Ilustrasi/Lampiran).
 *
 * Field tetap otomatis: Word memperbaruinya saat dibuka (updateFields). Tetapi aplikasi lain
 * (LibreOffice, pratinjau Google Drive/WhatsApp, Word yang menolak pembaruan) menampilkan hasil
 * tersimpan apa adanya — tanpa isi, kolom "Halaman" kosong. Di sini hasil tersimpan diisi dengan
 * PERKIRAAN halaman: tata letak dihitung kasar (tinggi baris, lebar teks, tabel, gambar inline,
 * pindah halaman, section & penomorannya). Pada naskah uji ±75% tepat, sisanya meleset 1 halaman; begitu field
 * diperbarui, angka sebenarnya menggantikannya.
 */
import { flatten, parseDocument, type BlockNode, type CellNode, type ParaNode, type TableNode } from "./document-model";
import { measureLines, type TextPiece } from "../front/metrics";

type Spacing = { before: number; after: number; line: number; rule: string; pageBreakBefore: boolean; sz: number };
type Section = { w: number; h: number; top: number; bottom: number; left: number; right: number; fmt: string; start: number | null; continuous: boolean };

const num = (re: RegExp, s: string, d = 0) => { const m = re.exec(s); return m ? Number(m[1]) : d; };

function readSection(x: string): Section {
    const pgSz = /<w:pgSz\b[^>]*\/>/.exec(x)?.[0] ?? "";
    const mar = /<w:pgMar\b[^>]*\/>/.exec(x)?.[0] ?? "";
    const nt = /<w:pgNumType\b[^>]*\/>/.exec(x)?.[0] ?? "";
    return {
        w: num(/w:w="(\d+)"/, pgSz, 11906), h: num(/w:h="(\d+)"/, pgSz, 16838),
        top: num(/w:top="(-?\d+)"/, mar, 1440), bottom: num(/w:bottom="(-?\d+)"/, mar, 1440),
        left: num(/w:left="(\d+)"/, mar, 1440), right: num(/w:right="(\d+)"/, mar, 1440),
        fmt: /w:fmt="(\w+)"/.exec(nt)?.[1] ?? "decimal",
        start: /w:start="(\d+)"/.test(nt) ? num(/w:start="(\d+)"/, nt) : null,
        continuous: /<w:type w:val="continuous"\/>/.test(x),
    };
}

/** Spasi & ukuran huruf efektif per style (rantai basedOn, lalu docDefaults). */
function styleResolver(stylesXml: string) {
    const styles = new Map<string, { pPr: string; rPr: string; basedOn: string | null }>();
    let def: string | null = null;
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        if (!/w:type="paragraph"/.test(m[1])) continue;
        const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
        if (!id) continue;
        if (/w:default="(?:1|true|on)"/.test(m[1])) def = id;
        styles.set(id, {
            pPr: /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(m[2])?.[1] ?? "",
            rPr: /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[2])?.[1] ?? "",
            basedOn: /<w:basedOn w:val="([^"]+)"/.exec(m[2])?.[1] ?? null,
        });
    }
    const ddP = /<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/.exec(stylesXml)?.[0] ?? "";
    const ddR = /<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/.exec(stylesXml)?.[0] ?? "";
    const cache = new Map<string, Spacing>();
    return (id: string | null): Spacing => {
        const key = id ?? "";
        const hit = cache.get(key);
        if (hit) return hit;
        const chain: { pPr: string; rPr: string }[] = [];
        for (let s = id ?? def, d = 0; s && d < 12; d++) {
            const st = styles.get(s);
            if (!st) break;
            chain.push(st);
            s = st.basedOn;
        }
        chain.push({ pPr: ddP, rPr: ddR });
        const first = (re: RegExp, part: "pPr" | "rPr") => { for (const c of chain) { const m = re.exec(c[part]); if (m) return m; } return null; };
        const sp = (attr: string) => { for (const c of chain) { const t = /<w:spacing\b[^>]*\/>/.exec(c.pPr)?.[0]; const m = t ? new RegExp(`w:${attr}="([^"]+)"`).exec(t) : null; if (m) return m[1]; } return null; };
        const out: Spacing = {
            before: Number(sp("before") ?? 0), after: Number(sp("after") ?? 0),
            line: Number(sp("line") ?? 240), rule: sp("lineRule") ?? "auto",
            pageBreakBefore: Boolean(first(/<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/, "pPr")),
            sz: Number(first(/<w:sz w:val="(\d+)"/, "rPr")?.[1] ?? 24),
        };
        cache.set(key, out);
        return out;
    };
}

const ROMAN: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
const roman = (n: number) => { let s = ""; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const label = (n: number, fmt: string) => fmt === "lowerRoman" ? roman(n) : fmt === "upperRoman" ? roman(n).toUpperCase() : String(n);

/** Halaman (label tampil) tiap bookmark di dokumen. */
export function estimateBookmarkPages(documentXml: string, stylesXml: string): Map<string, string> {
    const model = parseDocument(documentXml);
    const style = styleResolver(stylesXml);
    const nodes = flatten(model.body);

    // Section berurutan: sectPr di paragraf menutup section; sisanya sectPr badan dokumen.
    const sections: Section[] = [];
    for (const n of nodes) if (n.kind === "paragraph" && n.hasSectPr) sections.push(readSection(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(n.xml)![0]));
    sections.push(readSection(model.bodySectPr ?? ""));

    let si = 0;
    let sec = sections[0];
    let page = 1;              // nomor halaman fisik di dalam hitungan
    let pageNo = sec.start ?? 1;
    let y = 0;
    const out = new Map<string, string>();
    const avail = () => Math.max(2000, sec.h - Math.max(0, sec.top) - Math.max(0, sec.bottom));
    const usable = () => Math.max(2000, sec.w - sec.left - sec.right);
    const newPage = () => { page++; pageNo++; y = 0; };

    /** Potongan teks paragraf (tebal/miring/ukuran per run) untuk mengukur jumlah baris. */
    const piecesOf = (p: ParaNode, defSz: number): TextPiece[] => {
        const out: TextPiece[] = [];
        const body = p.xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, "").replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, "").replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, "");
        for (const r of body.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)) {
            const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(r[1])?.[1] ?? "";
            const on = (t: string) => new RegExp(`<w:${t}(?:\\s+w:val="(?:1|true|on)")?\\s*/>`).test(rPr);
            const halfPts = Number(/<w:sz w:val="(\d+)"/.exec(rPr)?.[1] ?? defSz);
            let t = "";
            for (const a of r[1].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>|<w:br w:type="textWrapping"\/>/g)) {
                t += a[1] !== undefined ? a[1].replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">") : a[0].startsWith("<w:tab") ? "    " : "\n";
            }
            if (t) out.push({ t, b: on("b"), i: on("i"), halfPts });
        }
        for (const m of p.xml.matchAll(/<m:t(?:\s[^>]*)?>([^<]*)<\/m:t>/g)) out.push({ t: m[1], b: false, i: true, halfPts: defSz });
        return out;
    };
    type Measure = { before: number; after: number; lineH: number; lines: number; extra: number; keepNext: boolean };
    const measure = (p: ParaNode, width: number): Measure => {
        const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
        const st = style(p.styleId);
        const spTag = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
        const get = (a: string, d: number) => { const m = new RegExp(`w:${a}="(-?\\d+)"`).exec(spTag); return m ? Number(m[1]) : d; };
        const rule = /w:lineRule="(\w+)"/.exec(spTag)?.[1] ?? st.rule;
        const line = get("line", st.line);
        const pieces = piecesOf(p, st.sz);
        const sz = Math.max(st.sz, ...pieces.map((x) => x.halfPts));
        const natural = sz * 11.5;          // tinggi baris tunggal Times New Roman ±1,15 × ukuran huruf
        const lineH = rule === "exact" ? line : rule === "atLeast" ? Math.max(line, natural) : (line / 240) * natural;
        const ind = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
        const w = Math.max(1000, width - num(/w:left="(\d+)"/, ind) - num(/w:right="(\d+)"/, ind));
        const lines = pieces.length ? measureLines(pieces, w) : 1;
        let extra = 0;
        for (const m of p.xml.matchAll(/<wp:inline\b[\s\S]*?<wp:extent cx="\d+" cy="(\d+)"/g)) extra += Number(m[1]) / 635;
        const keepNext = /<w:keepNext(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPr) || /^(heading|judul)\s*\d/i.test(p.styleId ?? "");
        return { before: get("before", st.before), after: get("after", st.after), lineH, lines, extra, keepNext };
    };
    const paraHeight = (p: ParaNode, width: number): number => {
        const m = measure(p, width);
        return m.before + m.lines * m.lineH + m.extra + m.after;
    };
    const cellHeight = (c: CellNode, width: number): number => {
        let h = 80;
        for (const b of c.children) {
            if (b.kind === "paragraph") h += paraHeight(b, width - 170);
            else if (b.kind === "table") h += tableHeight(b, width - 170);
        }
        return h;
    };
    const tableHeight = (t: TableNode, width: number): number => {
        let h = 0;
        for (const r of t.rows) {
            const n = Math.max(1, r.cells.length);
            const widths = r.cells.map((c) => {
                const tcW = /<w:tcW w:w="(\d+)" w:type="(\w+)"/.exec(c.props ?? "");
                if (tcW?.[2] === "pct") return (Number(tcW[1]) / 5000) * width;
                if (tcW?.[2] === "dxa" && Number(tcW[1]) > 0) return Number(tcW[1]);
                return width / n;
            });
            const exact = /<w:trHeight w:val="(\d+)"(?: w:hRule="(\w+)")?/.exec(r.props ?? "");
            let rh = Math.max(...r.cells.map((c, i) => cellHeight(c, widths[i])));
            if (exact) rh = exact[2] === "exact" ? Number(exact[1]) : Math.max(rh, Number(exact[1]));
            h += rh;
        }
        return h;
    };
    const mark = (xml: string) => {
        const lab = label(pageNo, sec.fmt);
        for (const m of xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="([^"]+)"/g)) if (!out.has(m[1])) out.set(m[1], lab);
    };
    const place = (h: number) => {
        const H = avail();
        if (y > 0 && y + h > H) {
            // Sisa yang tidak muat pindah ke halaman berikutnya (paragraf panjang terbelah).
            const rest = y + h - H;
            newPage();
            y = rest;
            while (y > H) { newPage(); y -= H; }
        } else y += h;
    };

    const visit = (b: BlockNode) => {
        if (b.kind === "raw") { mark(b.xml); return; }
        if (b.kind === "container") { b.children.forEach(visit); return; }
        if (b.kind === "table") {
            mark(b.rows.map((r) => r.cells.map((c) => c.children.map((x) => (x.kind === "paragraph" ? x.xml : "")).join("")).join("")).join(""));
            place(tableHeight(b, usable()));
            return;
        }
        const p = b;
        const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
        const pbb = /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPr) ||
            (style(p.styleId).pageBreakBefore && !/<w:pageBreakBefore\s+w:val="(?:0|false|off)"/.test(pPr));
        const brAt = p.xml.search(/<w:br w:type="page"\/>/);
        const textBeforeBr = brAt > -1 && /<w:t(?:\s[^>]*)?>[^<]+<\/w:t>/.test(p.xml.slice(0, brAt));
        if ((pbb || (brAt > -1 && !textBeforeBr)) && y > 0) newPage();
        const m = measure(p, usable());
        const H = avail();
        // Judul (keepNext) tidak tertinggal sendirian di bawah halaman: ia + 1 baris berikutnya harus muat
        // (dikalibrasi dengan tata letak LibreOffice pada naskah uji: ±75% tepat, sisanya meleset 1).
        if (m.keepNext && y > 0 && y + m.before + m.lines * m.lineH + m.lineH > H) newPage();
        // Kontrol baris janda/yatim: paragraf yang hanya menyisakan 1 baris di bawah halaman pindah utuh.
        else if (!m.extra && m.lines >= 2 && y > 0) {
            const fit = Math.floor((H - y - m.before) / m.lineH);
            if (fit < 2 && fit < m.lines) newPage();
        }
        if (m.extra && y > 0 && y + m.before + m.extra + m.lineH > H) newPage();   // gambar inline tidak terbelah
        mark(p.xml);
        place(m.before + m.lines * m.lineH + m.extra + m.after);
        if (brAt > -1 && textBeforeBr) newPage();
        if (p.hasSectPr) {
            // Section berikutnya: halaman baru (kecuali continuous), penomoran mulai ulang bila diminta.
            si++;
            sec = sections[Math.min(si, sections.length - 1)];
            if (!sec.continuous) { page++; y = 0; pageNo = sec.start ?? pageNo + 1; }
            else if (sec.start !== null) pageNo = sec.start;
        }
    };
    nodes.forEach(visit);
    return out;
}

/**
 * Isi hasil tersimpan field PAGEREF dengan perkiraan halaman. Field tetap field (bisa diperbarui);
 * yang diganti hanya teks hasilnya. Mengembalikan XML baru & jumlah field yang diisi.
 */
export function fillPageRefs(documentXml: string, stylesXml: string): { xml: string; filled: number } {
    if (!/PAGEREF/.test(documentXml)) return { xml: documentXml, filled: 0 };
    let pages: Map<string, string>;
    try {
        pages = estimateBookmarkPages(documentXml, stylesXml);
    } catch {
        return { xml: documentXml, filled: 0 };
    }
    let filled = 0;
    const re = /(<w:instrText[^>]*>\s*PAGEREF\s+(\S+)[^<]*<\/w:instrText><\/w:r>)(<w:r>(<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:fldChar w:fldCharType="separate"\/><\/w:r>)((?:(?!w:fldCharType=)[\s\S])*?)(<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:fldChar w:fldCharType="end"\/><\/w:r>)/g;
    const xml = documentXml.replace(re, (all, instr: string, name: string, sep: string, rPr: string | undefined, result: string, end: string) => {
        const lab0 = pages.get(name);
        const lab = lab0 && /\\\* roman/.test(instr) && /^\d+$/.test(lab0) ? roman(Number(lab0)) : lab0;
        if (!lab) return all;
        // Hasil lama yang bukan sekadar nomor (mis. teks lain) tidak disentuh.
        const old = [...result.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("").trim();
        if (old && !/^[0-9ivxlcdm]+$/i.test(old)) return all;
        const rest = result.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").replace(/<\/?w:r(?:\s[^>]*)?>|<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g, "");
        if (rest.trim()) return all;
        filled++;
        return `${instr}${sep}<w:r>${rPr ?? ""}<w:t>${lab}</w:t></w:r>${end}`;
    });
    return { xml, filled };
}
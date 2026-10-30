/**
 * Pembacaan bagian depan menjadi aliran item: baris teks (dengan format miring/sub/sup),
 * baris kosong, gambar, dan batas halaman.
 *
 * Pembacaan mengikuti struktur XML (bukan regex datar), jadi:
 *  - teks di dalam kotak teks (text box) ikut terbaca sebagai baris sendiri,
 *  - salinan cadangan bentuk Word (mc:Fallback) tidak terbaca dua kali,
 *  - teks yang dihapus dengan lacak perubahan (w:del) tidak ikut,
 *  - ganti baris manual (Shift+Enter) memecah baris.
 *
 * Objek yang tidak bisa dipakai ulang (gambar format lama, grafik, SmartArt, OLE) tidak
 * menggagalkan pembacaan: teksnya tetap dibaca, objeknya ditandai `unsupported` dan dilaporkan.
 */
import { childEl, elementEnd, innerXml, splitElements } from "../docx/xml";
import { parseBlocks, type BlockNode, type ParaNode, type TableNode } from "../docx/document-model";
import type { Item, Pic, Rich, Seg } from "./types";

const decode = (s: string) =>
    s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

const ON = (rPr: string, tag: string) => {
    const m = new RegExp(`<w:${tag}(?:\\s+w:val="([^"]*)")?\\s*/>`).exec(rPr);
    return Boolean(m) && !/^(0|false|off)$/i.test(m![1] ?? "");
};

type Flags = { it: boolean; sub: boolean; sup: boolean; b: boolean };

function flagsOf(rPr: string): Flags {
    const va = /<w:vertAlign\s+w:val="(\w+)"/.exec(rPr)?.[1];
    return { it: ON(rPr, "i"), b: ON(rPr, "b"), sub: va === "subscript", sup: va === "superscript" };
}

type Tok =
    | { k: "text"; t: string; f: Flags }
    | { k: "soft" }
    | { k: "page" }
    | { k: "pic"; pic: Pic }
    | { k: "inner"; xml: string };

const PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/** Posisi tegak objek DrawingML melayang (wp:anchor): acuan & jarak atas (twip). */
function anchorOfDrawing(drawing: string): Pic["anchor"] {
    if (!/<wp:anchor\b/.test(drawing)) return undefined;
    const pv = /<wp:positionV\b[^>]*relativeFrom="(\w+)"[^>]*>([\s\S]*?)<\/wp:positionV>/.exec(drawing);
    if (!pv) return undefined;
    const off = /<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(pv[2]);
    if (!off) return undefined;           // rata atas/tengah/bawah halaman: tidak dihitung
    return { rel: pv[1], top: Math.round(Number(off[1]) / 635) };
}

/** Posisi tegak bentuk VML melayang (style margin-top + mso-position-vertical-relative). */
function anchorOfVml(attrs: string): Pic["anchor"] {
    const style = /\bstyle="([^"]*)"/.exec(attrs)?.[1] ?? "";
    if (!/position:\s*absolute/.test(style)) return undefined;
    const top = /(?:^|;)\s*margin-top\s*:\s*([^;]+)/.exec(style)?.[1];
    if (!top) return undefined;
    const rel = /mso-position-vertical-relative\s*:\s*([\w-]+)/.exec(style)?.[1] ?? "text";
    return { rel: rel === "text" ? "paragraph" : rel, top: Math.round(vmlLength(top) / 635) };
}

function picOf(drawing: string): Pic {
    const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(drawing);
    const uri = /<a:graphicData\b[^>]*\buri="([^"]+)"/.exec(drawing)?.[1] ?? "";
    return {
        xml: drawing,
        picture: uri === PICTURE_URI,
        cx: Number(ext?.[1] ?? 0),
        cy: Number(ext?.[2] ?? 0),
        embeds: [...drawing.matchAll(/<a:blip\b[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]),
        anchor: anchorOfDrawing(drawing),
    };
}

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Gambar DrawingML (inline) dari satu gambar relasi `rId` berukuran cx × cy EMU. */
function pictureXml(rId: string, cx: number, cy: number): string {
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="0" name="Tanda tangan"/>` +
        `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PICTURE_URI}"><pic:pic xmlns:pic="${PICTURE_URI}">` +
        `<pic:nvPicPr><pic:cNvPr id="0" name="Tanda tangan"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

/** Panjang CSS VML ("93.75pt", "1.2in", "40px", "2cm") → EMU. */
function vmlLength(v: string | undefined): number {
    const m = /^(-?[\d.]+)\s*(pt|in|cm|mm|px|emu)?$/.exec((v ?? "").trim());
    if (!m) return 0;
    const n = Number(m[1]);
    switch (m[2]) {
        case "in": return Math.round(n * 914400);
        case "cm": return Math.round(n * 360000);
        case "mm": return Math.round(n * 36000);
        case "px": return Math.round(n * 9525);
        case "emu": return Math.round(n);
        default: return Math.round(n * 12700);
    }
}

/** Gambar VML lama (<v:shape><v:imagedata r:id>) → gambar DrawingML yang bisa dipakai ulang. */
function picsFromVml(xml: string): Pic[] {
    const out: Pic[] = [];
    for (const m of xml.matchAll(/<v:(shape|rect|image)\b([^>]*)>([\s\S]*?)<\/v:\1>/g)) {
        const rid = /<v:imagedata\b[^>]*\br:id="([^"]+)"/.exec(m[3])?.[1];
        if (!rid) continue;
        const style = /\bstyle="([^"]*)"/.exec(m[2])?.[1] ?? "";
        const prop = (k: string) => new RegExp(`(?:^|;)\\s*${k}\\s*:\\s*([^;]+)`).exec(style)?.[1];
        const cx = vmlLength(prop("width"));
        const cy = vmlLength(prop("height"));
        if (cx <= 0 || cy <= 0) continue;
        out.push({ xml: pictureXml(rid, cx, cy), picture: true, cx, cy, embeds: [rid], anchor: anchorOfVml(m[2]) });
    }
    return out;
}

/**
 * Grup gambar (wpg:wgp) berisi beberapa gambar — mis. dua tanda tangan + garis dalam satu objek:
 * tiap gambar dikeluarkan sendiri-sendiri (ukuran sesuai skala grup), urut atas→bawah, kiri→kanan.
 * Bentuk lain di dalam grup (garis, kotak) diabaikan.
 */
function picsFromGroup(drawing: string): Pic[] {
    const grp = /<wpg:grpSpPr>[\s\S]*?<\/wpg:grpSpPr>/.exec(drawing)?.[0] ?? "";
    const n = (re: RegExp, s: string) => Number(re.exec(s)?.[1] ?? 0);
    const ext = { cx: n(/<a:ext\s+cx="(\d+)"/, grp), cy: n(/<a:ext\s+cx="\d+"\s+cy="(\d+)"/, grp) };
    const chExt = { cx: n(/<a:chExt\s+cx="(\d+)"/, grp), cy: n(/<a:chExt\s+cx="\d+"\s+cy="(\d+)"/, grp) };
    const sx = chExt.cx ? ext.cx / chExt.cx : 1;
    const sy = chExt.cy ? ext.cy / chExt.cy : 1;
    const kids: { pic: Pic; x: number; y: number }[] = [];
    for (const m of drawing.matchAll(/<pic:pic\b[\s\S]*?<\/pic:pic>/g)) {
        const rid = /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(m[0])?.[1];
        const xf = /<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/.exec(m[0])?.[0] ?? "";
        if (!rid || !xf) continue;
        const cx = Math.round(n(/<a:ext\s+cx="(\d+)"/, xf) * sx);
        const cy = Math.round(n(/<a:ext\s+cx="\d+"\s+cy="(\d+)"/, xf) * sy);
        if (cx <= 0 || cy <= 0) continue;
        kids.push({ pic: { xml: pictureXml(rid, cx, cy), picture: true, cx, cy, embeds: [rid], fromGroup: true }, x: n(/<a:off\s+x="(-?\d+)"/, xf), y: n(/<a:off\s+x="-?\d+"\s+y="(-?\d+)"/, xf) });
    }
    // Posisi tiap gambar = posisi grup + letaknya di dalam grup.
    const g = anchorOfDrawing(drawing);
    const chOffY = n(/<a:chOff\s+x="-?\d+"\s+y="(-?\d+)"/, grp);
    if (g) for (const k of kids) k.pic.anchor = { rel: g.rel, top: g.top + Math.round(((k.y - chOffY) * sy) / 635) };
    // Baris yang sama (selisih tinggi < setengah gambar) dibaca kiri→kanan.
    kids.sort((a, b) => (Math.abs(a.y - b.y) < Math.min(a.pic.cy, b.pic.cy) / sy / 2 ? a.x - b.x : a.y - b.y));
    return kids.map((k) => k.pic);
}

/** Isi paragraf di dalam kotak teks sebuah objek. */
function textBoxes(xml: string): string[] {
    const out: string[] = [];
    const re = /<w:txbxContent(?=[\s>])/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
        const end = elementEnd(xml, m.index, "w:txbxContent");
        out.push(innerXml(xml.slice(m.index, end)));
        re.lastIndex = end;
    }
    return out;
}

function objectTokens(el: { tag: string; xml: string }, out: Tok[]) {
    if (el.tag === "mc:AlternateContent") {
        // Hanya versi utama (mc:Choice); mc:Fallback adalah salinan cadangan — KECUALI tinta digital
        // (w14:contentPart, tanda tangan ditulis dengan pena/jari di Word): versi utamanya bukan gambar,
        // jadi gambar cadangannya (VML) yang dipakai.
        const parts = splitElements(innerXml(el.xml));
        const choice = parts.find((c) => c.tag === "mc:Choice");
        const fallback = parts.find((c) => c.tag === "mc:Fallback");
        if (choice && /<w14:contentPart\b/.test(choice.xml) && !/<a:blip\b/.test(choice.xml) && fallback && /<v:imagedata\b/.test(fallback.xml)) {
            for (const c of splitElements(innerXml(fallback.xml))) objectTokens(c, out);
            return;
        }
        if (choice) for (const c of splitElements(innerXml(choice.xml))) objectTokens(c, out);
        return;
    }
    // Tinta digital di dalam w:drawing: gambar cadangannya (a:pic/a:blip) dipakai, seukuran objeknya.
    if (el.tag === "w:drawing" && /<w14:contentPart\b/.test(el.xml)) {
        const rid = /<a:blip\b[^>]*\br:embed="([^"]+)"/.exec(el.xml)?.[1];
        const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(el.xml);
        if (rid && ext && Number(ext[1]) > 0 && Number(ext[2]) > 0) {
            const cx = Number(ext[1]), cy = Number(ext[2]);
            out.push({ k: "pic", pic: { xml: pictureXml(rid, cx, cy), picture: true, cx, cy, embeds: [rid], anchor: anchorOfDrawing(el.xml) } });
            return;
        }
    }
    if (el.tag === "w:drawing" && /<wpg:wgp\b/.test(el.xml) && /<pic:pic\b/.test(el.xml)) {
        const kids = picsFromGroup(el.xml);
        if (kids.length) {
            for (const pic of kids) out.push({ k: "pic", pic });
            for (const b of textBoxes(el.xml)) out.push({ k: "inner", xml: b });
            return;
        }
    }
    if (el.tag === "w:drawing") {
        const pic = picOf(el.xml);
        if (!pic.picture && /<a:blip\b|<c:chart\b|<dgm:relIds\b/.test(el.xml)) {
            pic.unsupported = "objek gabungan (grup gambar, grafik, atau SmartArt)";
        }
        out.push({ k: "pic", pic });
        for (const b of textBoxes(el.xml)) out.push({ k: "inner", xml: b });
        return;
    }
    if (el.tag === "w:pict") {
        const old = /<v:imagedata\b/.test(el.xml);
        // Gambar VML (format lama / cadangan tinta digital): diubah ke gambar biasa bila ukurannya terbaca.
        const vml = old ? picsFromVml(el.xml) : [];
        if (vml.length) {
            for (const pic of vml) out.push({ k: "pic", pic: { ...pic, fromGroup: vml.length > 1 || undefined } });
            for (const b of textBoxes(el.xml)) out.push({ k: "inner", xml: b });
            return;
        }
        // Bentuk/kotak teks VML: teksnya dibaca, bentuknya dianggap bingkai.
        out.push({ k: "pic", pic: { xml: el.xml, picture: false, cx: 0, cy: 0, embeds: [], unsupported: old ? "gambar format lama (VML)" : undefined } });
        for (const b of textBoxes(el.xml)) out.push({ k: "inner", xml: b });
        return;
    }
    if (el.tag === "w:object") {
        out.push({ k: "pic", pic: { xml: el.xml, picture: false, cx: 0, cy: 0, embeds: [], unsupported: "objek tertanam (OLE/persamaan lama)" } });
    }
}

function runTokens(run: string, out: Tok[]) {
    const f = flagsOf(childEl(run, "w:rPr") ?? "");
    for (const c of splitElements(innerXml(run))) {
        switch (c.tag) {
            case "w:t":
                out.push({ k: "text", t: decode(innerXml(c.xml)), f });
                break;
            case "w:tab":
            case "w:ptab":
                out.push({ k: "text", t: "\t", f });
                break;
            case "w:br":
                if (/w:type="page"/.test(c.xml)) out.push({ k: "page" });
                else if (!/w:type="column"/.test(c.xml)) out.push({ k: "soft" });
                break;
            case "w:cr":
                out.push({ k: "soft" });
                break;
            case "w:noBreakHyphen":
                out.push({ k: "text", t: "-", f });
                break;
            case "w:sym":
                break;
            default:
                objectTokens(c, out);
        }
    }
}

/** Isi paragraf (tanpa pPr), menelusuri hyperlink, revisi sisipan, content control, dll. */
function contentTokens(xml: string, out: Tok[]) {
    for (const c of splitElements(xml)) {
        switch (c.tag) {
            case "w:r":
                runTokens(c.xml, out);
                break;
            case "w:hyperlink": case "w:ins": case "w:smartTag": case "w:customXml":
            case "w:fldSimple": case "w:dir": case "w:bdo": case "w:moveTo":
                contentTokens(innerXml(c.xml), out);
                break;
            case "w:sdt": {
                const content = splitElements(innerXml(c.xml)).find((x) => x.tag === "w:sdtContent");
                if (content) contentTokens(innerXml(content.xml), out);
                break;
            }
            case "m:oMath": case "m:oMathPara":
                for (const m of c.xml.matchAll(/<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g)) {
                    out.push({ k: "text", t: decode(m[1]), f: { it: false, sub: false, sup: false, b: false } });
                }
                break;
            default:
                break; // w:pPr, w:del, w:moveFrom, bookmark, proofErr, commentRange, …
        }
    }
}

function paraContent(p: ParaNode): string {
    const inner = innerXml(p.xml);
    const lead = inner.length - inner.trimStart().length;
    if (inner.trimStart().startsWith("<w:pPr")) return inner.slice(elementEnd(inner, lead, "w:pPr"));
    return inner;
}

function cleanSegs(segs: Seg[]): Seg[] {
    const out: Seg[] = [];
    for (const s of segs) {
        const t = s.t.replace(/[ \u00a0]+/g, " ");
        if (!t) continue;
        const last = out[out.length - 1];
        if (last && last.it === s.it && last.sub === s.sub && last.sup === s.sup && last.b === s.b) last.t += t;
        else out.push({ ...s, t });
    }
    if (out.length) {
        out[0].t = out[0].t.replace(/^[\s]+/, "");
        out[out.length - 1].t = out[out.length - 1].t.replace(/[\s]+$/, "");
    }
    return out.filter((s) => s.t);
}

export const richText = (r: Rich) => r.map((s) => s.t).join("");

function readParagraph(p: ParaNode, inTable: boolean, out: Item[]) {
    if (p.hadPageBreakBefore) out.push({ k: "break" });
    const toks: Tok[] = [];
    contentTokens(paraContent(p), toks);

    const inner: string[] = [];
    const pics: Pic[] = [];
    let cur: Seg[] = [];
    let any = false;
    const flush = () => {
        const segs = cleanSegs(cur);
        cur = [];
        const raw = segs.map((s) => s.t).join("");
        const text = raw.replace(/\t+/g, " ").replace(/\s+/g, " ").trim();
        if (text) { out.push({ k: "line", segs, text, raw, para: p, inTable }); any = true; }
    };
    for (const t of toks) {
        if (t.k === "text") cur.push({ t: t.t, ...t.f });
        else if (t.k === "soft") flush();
        else if (t.k === "page") {
            flush();
            for (const pic of pics.splice(0)) { out.push({ k: "pic", pic, para: p }); any = true; }
            // Isi kotak teks sebelum pindah halaman ikut halaman SEBELUMNYA (mis. seluruh lembar
            // pernyataan di dalam kotak teks, lalu page break di paragraf yang sama).
            for (const x of inner.splice(0)) { readBlocks(parseBlocks(x, -1), true, out); any = true; }
            out.push({ k: "break" });
        } else if (t.k === "pic") pics.push(t.pic);
        else inner.push(t.xml);
    }
    flush();
    for (const pic of pics) { out.push({ k: "pic", pic, para: p }); any = true; }
    const hadInner = inner.length > 0;
    for (const x of inner) readBlocks(parseBlocks(x, -1), true, out);
    if (!any && !hadInner) out.push({ k: "blank", para: p });

    const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(p.xml)?.[0];
    if (sect && !/<w:type\s+w:val="continuous"/.test(sect)) out.push({ k: "break" });
}

function readTable(t: TableNode, out: Item[]) {
    for (const r of t.rows) for (const c of r.cells) readBlocks(c.children, true, out);
}

export function readBlocks(blocks: BlockNode[], inTable: boolean, out: Item[] = []): Item[] {
    for (const b of blocks) {
        if (b.kind === "paragraph") readParagraph(b, inTable, out);
        else if (b.kind === "table") readTable(b, out);
        else if (b.kind === "container") readBlocks(b.children, inTable, out);
        // Pembungkus blok lain (mis. w:customXml) yang berisi paragraf: isinya tetap dibaca.
        else if (/<w:p[\s>]/.test(b.xml)) readBlocks(parseBlocks(innerXml(b.xml), -1), inTable, out);
    }
    return out;
}
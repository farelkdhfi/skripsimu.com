/**
 * Penyusun bagian depan dari BERKAS TEMPLATE DOCX.
 *
 * Template dibaca apa adanya (paragraf, tabel, jarak, huruf, logo, section, footer) lalu:
 *  - teks contoh (judul, nama, NPM, pembimbing, tanggal, …) diganti data mahasiswa — format run
 *    template dipakai, dari naskah hanya teks dan cetak miring (nama ilmiah, istilah asing);
 *  - isi kata pengantar & abstrak diganti paragraf naskah dengan format paragraf template;
 *  - tanda tangan/materai contoh di template dibuang, diganti tanda tangan dari naskah (bila ada)
 *    pada bingkai/posisi yang sama;
 *  - style template disalin ke dokumen dengan nama sendiri (FmtTpl…), jadi tampilan tidak
 *    bergantung pada style naskah;
 *  - judul yang lebih panjang/pendek dari contoh diimbangi baris kosong sesudahnya, jadi isi
 *    halaman di bawahnya tetap di posisi template.
 */
import JSZip from "jszip";
import { parseBlocks, parseDocument, serializeBlock, type BlockNode, type ParaNode } from "../docx/document-model";
import { elementEnd, escapeXml, innerXml, mergeProps, openTag, splitElements, PPR_ORDER, RPR_ORDER } from "../docx/xml";
import { addContentTypeOverride, findOrAddRelationship } from "../docx/package-parts";
import { ABSTRACT_RE, ABSTRAK_RE, KP_RE, bare } from "./segment";
import { measureLines, type TextPiece } from "./metrics";
import { FIELD_LABEL, type FrontFieldKey, type FrontFields, type FrontPageKind, type FrontSpec, type LineItem, type Pic, type Rich, type Seg } from "./types";

const CM_EMU = 360000;
/** Lokasi berkas template di proyek (sumber) dan URL publiknya (untuk pemrosesan di browser). */
export const TEMPLATE_DIR = ["src", "features", "thesis-formatter", "presets", "templates"];
export const TEMPLATE_URL = "/templates";

/**
 * Pemuat berkas template. Bawaan: diunduh dari `/templates/<file>` (folder `public/templates`),
 * jadi bisa jalan di browser / Web Worker. Di Node (tes, server) diganti pemuat berkas lokal lewat
 * `setTemplateLoader` (lihat front/node.ts).
 */
export type TemplateLoader = (file: string) => Promise<Uint8Array>;
let loader: TemplateLoader = async (file) => {
    const res = await fetch(`${TEMPLATE_URL}/${file}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return new Uint8Array(await res.arrayBuffer());
};
export function setTemplateLoader(fn: TemplateLoader): void {
    loader = fn;
    cache.clear();
}

/* ------------------------------------------------------------------ */
/* Pemuatan template                                                    */
/* ------------------------------------------------------------------ */

export class TemplateError extends Error {}

type TplRel = { type: string; target: string; mode: string | null };

export type LoadedTemplate = {
    zip: JSZip;
    /** Tag pembuka w:document template (deklarasi namespace). */
    rootTag: string;
    /** XML bersih tiap halaman (section) sesuai `spec.pages`, TANPA sectPr. */
    pages: { kind: FrontPageKind; xml: string; sectPr: string }[];
    styles: string;
    rels: Map<string, TplRel>;
};

const cache = new Map<string, Promise<LoadedTemplate>>();


/** Atribut & penanda yang tidak boleh ikut disalin (id revisi, bookmark contoh, penanda ejaan). */
function clean(xml: string): string {
    return xml
        .replace(/\s(?:w:rsid\w*|w14:paraId|w14:textId|wp14:anchorId|wp14:editId)="[^"]*"/g, "")
        .replace(/<w:(?:bookmarkStart|bookmarkEnd|proofErr|permStart|permEnd|commentRangeStart|commentRangeEnd)\b[^>]*\/>/g, "")
        .replace(/<w:lastRenderedPageBreak\/>/g, "")
        .replace(/<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:commentReference\b[^>]*\/><\/w:r>/g, "");
}

async function readTemplate(spec: FrontSpec): Promise<LoadedTemplate> {
    let buf: Uint8Array;
    try {
        buf = await loader(spec.file);
    } catch (e) {
        throw new TemplateError(`berkas template "${spec.file}" tidak bisa dimuat (${e instanceof Error ? e.message : e}). Pastikan berkasnya ada di public${TEMPLATE_URL}/ dan ${TEMPLATE_DIR.join("/")}/.`);
    }
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file("word/document.xml")?.async("string");
    if (!docXml) throw new TemplateError(`berkas template "${spec.file}" bukan DOCX yang sah.`);
    const model = parseDocument(docXml);
    const rootTag = /<w:document\b[^>]*>/.exec(model.head)?.[0] ?? "";

    // Satu section = satu halaman template.
    const sections: { blocks: string[]; sectPr: string }[] = [];
    let cur: string[] = [];
    for (const b of model.body) {
        if (b.kind === "raw") continue;
        const xml = serializeBlock(b);
        if (b.kind === "paragraph" && b.hasSectPr) {
            const sectPr = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(b.xml)![0];
            cur.push(xml.replace(sectPr, ""));
            sections.push({ blocks: cur, sectPr });
            cur = [];
        } else cur.push(xml);
        if (sections.length === spec.pages.length) break;
    }
    if (sections.length < spec.pages.length) {
        throw new TemplateError(`berkas template hanya berisi ${sections.length} section; spesifikasi meminta ${spec.pages.length} halaman.`);
    }
    const pages = sections.map((s, i) => ({ kind: spec.pages[i], xml: clean(s.blocks.join("")), sectPr: clean(s.sectPr) }));

    // Pemeriksaan ringan: halaman prosa harus punya judulnya.
    const headRe: Partial<Record<FrontPageKind, RegExp>> = { kataPengantar: KP_RE, abstrak: ABSTRAK_RE, abstract: ABSTRACT_RE };
    for (const p of pages) {
        const re = headRe[p.kind];
        if (re && !paragraphsOf(parseBlocks(p.xml, 0)).some((q) => re.test(bare(q.text)))) {
            throw new TemplateError(`halaman template "${p.kind}" tidak memiliki judul halamannya; periksa urutan halaman di spesifikasi.`);
        }
    }

    const relXml = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "";
    const rels = new Map<string, TplRel>();
    for (const m of relXml.matchAll(/<Relationship\b[^>]*>/g)) {
        const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
        const type = /\bType="([^"]+)"/.exec(m[0])?.[1];
        const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
        if (id && type && target) rels.set(id, { type, target, mode: /\bTargetMode="([^"]+)"/.exec(m[0])?.[1] ?? null });
    }
    const styles = (await zip.file("word/styles.xml")?.async("string")) ?? "";
    return { zip, rootTag, pages, styles, rels };
}

export function loadTemplate(spec: FrontSpec): Promise<LoadedTemplate> {
    const key = spec.file;
    let p = cache.get(key);
    if (!p) {
        p = readTemplate(spec);
        cache.set(key, p);
        p.catch(() => cache.delete(key));
    }
    return p;
}

/* ------------------------------------------------------------------ */
/* Paragraf: run & teks                                                 */
/* ------------------------------------------------------------------ */

function paragraphsOf(blocks: BlockNode[], out: ParaNode[] = []): ParaNode[] {
    for (const b of blocks) {
        if (b.kind === "paragraph") out.push(b);
        else if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) paragraphsOf(c.children, out);
        else if (b.kind === "container") paragraphsOf(b.children, out);
    }
    return out;
}

type Atom = { k: "t"; text: string } | { k: "tab" } | { k: "br" } | { k: "x"; xml: string; img: boolean };
type Run = { rPr: string; atoms: Atom[] };
type Part = { k: "run"; run: Run } | { k: "other"; xml: string };
type PModel = { open: string; pPr: string; parts: Part[] };

const decode = (s: string) =>
    s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function parseRun(xml: string): Run {
    let rPr = "";
    const atoms: Atom[] = [];
    for (const c of splitElements(innerXml(xml))) {
        if (c.tag === "w:rPr") rPr = c.xml;
        else if (c.tag === "w:t") atoms.push({ k: "t", text: decode(innerXml(c.xml)) });
        else if (c.tag === "w:tab") atoms.push({ k: "tab" });
        else if ((c.tag === "w:br" && !/w:type="(?:page|column)"/.test(c.xml)) || c.tag === "w:cr") atoms.push({ k: "br" });
        else atoms.push({ k: "x", xml: c.xml, img: c.tag === "w:drawing" || c.tag === "w:pict" || c.tag === "mc:AlternateContent" });
    }
    return { rPr, atoms };
}

function parsePara(xml: string): PModel {
    const open = openTag(xml);
    let pPr = "";
    const parts: Part[] = [];
    for (const c of splitElements(innerXml(xml))) {
        if (c.tag === "w:pPr") pPr = c.xml;
        else if (c.tag === "w:r") parts.push({ k: "run", run: parseRun(c.xml) });
        else parts.push({ k: "other", xml: c.xml });
    }
    return { open: open.endsWith("/>") ? "<w:p>" : open, pPr, parts };
}

function serializeRun(r: Run): string {
    if (!r.atoms.length) return "";
    const body = r.atoms.map((a) =>
        a.k === "t" ? (a.text ? `<w:t xml:space="preserve">${escapeXml(a.text)}</w:t>` : "")
            : a.k === "tab" ? "<w:tab/>" : a.k === "br" ? "<w:br/>" : a.xml).join("");
    return body ? `<w:r>${r.rPr}${body}</w:r>` : "";
}

function serializePara(m: PModel): string {
    return `${m.open}${m.pPr}${m.parts.map((p) => (p.k === "run" ? serializeRun(p.run) : p.xml)).join("")}</w:p>`;
}

const atomLen = (a: Atom) => (a.k === "t" ? a.text.length : a.k === "tab" || a.k === "br" ? 1 : 0);

function textOfModel(m: PModel): string {
    let s = "";
    for (const p of m.parts) if (p.k === "run") for (const a of p.run.atoms) s += a.k === "t" ? a.text : a.k === "tab" ? "\t" : a.k === "br" ? "\n" : "";
    return s;
}

/** Run & sifat huruf di setiap posisi karakter. */
function charRuns(m: PModel): Run[] {
    const out: Run[] = [];
    for (const p of m.parts) if (p.k === "run") for (const a of p.run.atoms) for (let i = 0; i < atomLen(a); i++) out.push(p.run);
    return out;
}

/** Ganti karakter [s, e) dengan run baru (format dasar: run di posisi `s`). */
function replaceSpan(m: PModel, s: number, e: number, make: (baseRPr: string) => Run[]): void {
    let pos = 0;
    let inserted = false;
    const parts: Part[] = [];
    for (const part of m.parts) {
        if (part.k === "other") { parts.push(part); continue; }
        const run = part.run;
        let cur: Run = { rPr: run.rPr, atoms: [] };
        const flush = () => { if (cur.atoms.length) parts.push({ k: "run", run: cur }); cur = { rPr: run.rPr, atoms: [] }; };
        for (const a of run.atoms) {
            const len = atomLen(a);
            if (!len) { cur.atoms.push(a); continue; }
            const from = pos;
            const to = pos + len;
            pos = to;
            if (from < s) {
                if (a.k === "t") cur.atoms.push({ k: "t", text: a.text.slice(0, Math.min(len, s - from)) });
                else cur.atoms.push(a);
            }
            if (!inserted && s >= from && s < to) {
                flush();
                for (const r of make(run.rPr)) parts.push({ k: "run", run: r });
                inserted = true;
            }
            if (to > e) {
                if (a.k === "t") cur.atoms.push({ k: "t", text: a.text.slice(Math.max(0, e - from)) });
                else if (from >= e) cur.atoms.push(a);
            }
        }
        flush();
    }
    if (!inserted) {
        const last = [...m.parts].reverse().find((p) => p.k === "run") as { k: "run"; run: Run } | undefined;
        for (const r of make(last?.run.rPr ?? "")) parts.push({ k: "run", run: r });
    }
    m.parts = parts;
}

const isOn = (rPr: string, tag: string) => {
    const x = new RegExp(`<w:${tag}(?:\\s+w:val="([^"]*)")?\\s*/>`).exec(rPr);
    return Boolean(x) && !/^(0|false|off)$/i.test(x![1] ?? "");
};

/** rPr dasar dengan cetak miring/sub/sup dari naskah (tebal, ukuran, huruf tetap dari template). */
function styledRPr(base: string, s: Partial<Seg>): string {
    const kids = mergeProps(base || null, {
        "w:i": s.it ? "<w:i/>" : null,
        "w:iCs": s.it ? "<w:iCs/>" : null,
        "w:vertAlign": s.sub ? `<w:vertAlign w:val="subscript"/>` : s.sup ? `<w:vertAlign w:val="superscript"/>` : null,
    }, RPR_ORDER);
    return kids.length ? `<w:rPr>${kids.join("")}</w:rPr>` : "";
}

function richRuns(base: string, value: Rich): Run[] {
    return value.filter((s) => s.t).map((s) => {
        const atoms: Atom[] = [];
        s.t.split(/(\t|\n)/).forEach((piece) => {
            if (piece === "\t") atoms.push({ k: "tab" });
            else if (piece === "\n") atoms.push({ k: "br" });
            else if (piece) atoms.push({ k: "t", text: piece });
        });
        return { rPr: styledRPr(base, s), atoms };
    });
}

/* ------------------------------------------------------------------ */
/* Kapital & perkiraan baris                                            */
/* ------------------------------------------------------------------ */

const BINOMIAL = /^[A-Z][a-z]+(?:\s+[a-z]+){1,2}(?:\s+(?:L|var\.?\s*[a-z]+|subsp\.?\s*[a-z]+))?$|^[A-Z][a-z]+\s+spp?\.?$/;
const LATIN = /^(in\s+vitro|in\s+vivo|in\s+situ|in\s+silico|et\s+al\.?|ad\s+libitum)$/i;

/** Huruf kapital semua, kecuali nama ilmiah miring (Ananas comosus) yang tetap seperti aslinya. */
export function upperRich(r: Rich): Rich {
    return r.map((s) => {
        const core = s.t.replace(/^[\s(]+|[\s.,;:)]+$/g, "");
        const keep = s.it && BINOMIAL.test(core) && !LATIN.test(core);
        return keep ? s : { ...s, t: s.t.toUpperCase() };
    });
}

/** Teks paragraf beserta gaya hurufnya (untuk menghitung jumlah baris). */
function piecesOf(m: PModel): TextPiece[] {
    const out: TextPiece[] = [];
    for (const p of m.parts) {
        if (p.k !== "run") continue;
        const r = p.run.rPr;
        const caps = isOn(r, "caps");
        const style = { b: isOn(r, "b"), i: isOn(r, "i"), halfPts: szOf(r) };
        for (const a of p.run.atoms) {
            if (a.k === "t") out.push({ ...style, t: caps ? a.text.toUpperCase() : a.text });
            else if (a.k === "tab") out.push({ ...style, t: " " });
            else if (a.k === "br") out.push({ ...style, t: "\n" });
        }
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Pencocokan teks contoh                                               */
/* ------------------------------------------------------------------ */

/** Huruf & angka saja (huruf kecil, tanpa diakritik) beserta posisinya di teks asli. */
function normIndex(text: string): { norm: string; at: number[] } {
    let norm = "";
    const at: number[] = [];
    [...text].forEach((ch, i) => {
        const base = ch.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
        for (const c of base) if (/[a-z0-9]/.test(c)) { norm += c; at.push(i); }
    });
    return { norm, at };
}

const normOf = (s: string) => normIndex(s).norm;

function findSpans(text: string, sample: string): [number, number][] {
    const ns = normOf(sample);
    if (!ns) return [];
    const { norm, at } = normIndex(text);
    // Tanda baca penutup contoh ("IPM.") ikut menjadi bagian contoh.
    const tail = (/[^\p{L}\p{N}]*$/u.exec(sample.trim())?.[0] ?? "").replace(/\s+/g, "");
    const out: [number, number][] = [];
    for (let k = norm.indexOf(ns); k !== -1; k = norm.indexOf(ns, k + ns.length)) {
        const s = at[k];
        let e = at[k + ns.length - 1] + 1;
        if (tail && text.startsWith(tail, e)) e += tail.length;
        // Tanda baca/spasi penutup paragraf ikut terganti (titik di akhir judul contoh).
        if (/^[\s.]*$/.test(text.slice(e))) e = text.length;
        out.push([s, e]);
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Gambar                                                               */
/* ------------------------------------------------------------------ */

type ImgSlot = { page: FrontPageKind; para: number; cx: number; cy: number; xml: string };

/** Bingkai anchor (posisi, wrap) dari gambar template; mc:AlternateContent → versi utamanya. */
function anchorOf(xml: string): string | null {
    return /<wp:anchor\b[\s\S]*?<\/wp:anchor>/.exec(xml)?.[0] ?? null;
}

/** Elemen a:graphic gambar naskah, dengan deklarasi namespace dari bungkusnya. */
function graphicOf(pic: Pic): { graphic: string; ns: string } | null {
    const graphic = /<a:graphic\b[\s\S]*?<\/a:graphic>/.exec(pic.xml)?.[0];
    if (!graphic) return null;
    const head = /^<w:drawing\b[^>]*>\s*<wp:(?:inline|anchor)\b[^>]*>/.exec(pic.xml)?.[0] ?? "";
    const ns = (head.match(/\sxmlns:[\w-]+="[^"]*"/g) ?? []).join("");
    return { graphic, ns };
}

function fitSize(pic: Pic, boxCx: number, boxCy: number): { cx: number; cy: number } {
    const k = Math.min(boxCx / Math.max(1, pic.cx), boxCy / Math.max(1, pic.cy));
    return { cx: Math.round(pic.cx * k), cy: Math.round(pic.cy * k) };
}

function sizedGraphic(g: string, pic: Pic, cx: number, cy: number): string {
    return g.replace(new RegExp(`<a:ext\\s+cx="${pic.cx}"\\s+cy="${pic.cy}"\\s*/>`, "g"), `<a:ext cx="${cx}" cy="${cy}"/>`);
}

/** Gambar naskah di bingkai gambar template (posisi sama; diperkecil agar muat, rata bawah-tengah). */
function inTemplateFrame(slot: ImgSlot, pic: Pic): string | null {
    const anchor = anchorOf(slot.xml);
    const g = graphicOf(pic);
    if (!anchor || !g) return null;
    const { cx, cy } = fitSize(pic, slot.cx, slot.cy);
    const dx = Math.round((slot.cx - cx) / 2);
    const dy = slot.cy - cy;
    let a = anchor
        .replace(/^<wp:anchor\b[^>]*>/, (t) => t.replace(/>$/, g.ns + ">").replace(/\sbehindDoc="1"/, ' behindDoc="0"'))
        .replace(/<wp:extent\b[^>]*\/>/, `<wp:extent cx="${cx}" cy="${cy}"/>`)
        .replace(/<wp:effectExtent\b[^>]*\/>/, `<wp:effectExtent l="0" t="0" r="0" b="0"/>`)
        .replace(/<a:graphic\b[\s\S]*?<\/a:graphic>/, sizedGraphic(g.graphic, pic, cx, cy))
        .replace(/<wp:docPr\b[^>]*?(\/>|>[\s\S]*?<\/wp:docPr>)/, `<wp:docPr id="0" name="Tanda tangan"/>`);
    a = a.replace(/(<wp:positionH\b[^>]*>\s*<wp:posOffset>)(-?\d+)/, (_m, p: string, v: string) => `${p}${Number(v) + dx}`);
    a = a.replace(/(<wp:positionV\b[^>]*>\s*<wp:posOffset>)(-?\d+)/, (_m, p: string, v: string) => `${p}${Number(v) + dy}`);
    return `<w:drawing>${a}</w:drawing>`;
}

/**
 * Tanda tangan melayang (tidak menggeser tata letak), ditambatkan pada `offsetEmu` di bawah
 * awal paragraf tambatannya.
 */
function floatingSignature(pic: Pic, maxCx: number, maxCy: number, offsetEmu: (cy: number) => number): string | null {
    const g = graphicOf(pic);
    if (!g) return null;
    const { cx, cy } = fitSize(pic, maxCx, maxCy);
    return `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251700000" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"${g.ns}>` +
        `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(offsetEmu(cy))}</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
        `<wp:docPr id="0" name="Tanda tangan"/><wp:cNvGraphicFramePr/>${sizedGraphic(g.graphic, pic, cx, cy)}</wp:anchor></w:drawing>`;
}

/**
 * Paragraf kosong tepat di atas `name` pada kolom yang sama (ruang tanda tangan): paragraf
 * teratasnya dan tinggi ruang itu (twip).
 */
function spaceAbove(blocks: BlockNode[], name: ParaNode): { para: ParaNode; height: number } | null {
    const loc = new Map<ParaNode, { t: BlockNode | null; col: number }>();
    for (const b of blocks) {
        if (b.kind === "paragraph") loc.set(b, { t: null, col: 0 });
        else if (b.kind === "table") {
            for (const r of b.rows) {
                let col = 0;
                for (const c of r.cells) {
                    for (const p of paragraphsOf(c.children)) loc.set(p, { t: b, col });
                    col += Number(/<w:gridSpan w:val="(\d+)"/.exec(c.props ?? "")?.[1] ?? 1);
                }
            }
        }
    }
    const order = paragraphsOf(blocks);
    const at = loc.get(name);
    const idx = order.indexOf(name);
    if (!at || idx === -1) return null;
    const block: ParaNode[] = [];
    for (let i = idx - 1; i >= 0; i--) {
        const p = order[i];
        const l = loc.get(p)!;
        if (l.t !== at.t || l.col !== at.col) continue;
        if (p.text.trim() || /<w:drawing|<w:pict|mc:AlternateContent/.test(p.xml)) break;
        block.push(p);
    }
    if (!block.length) return null;
    const height = block.reduce((h, p) => {
        const m = parsePara(p.xml);
        const i = pPrInfo(m.pPr);
        return h + lineHeight(m.pPr, szOf(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m.pPr)?.[0] ?? "")) + i.before + i.after;
    }, 0);
    return { para: block[block.length - 1], height };
}

/* ------------------------------------------------------------------ */
/* Style, gambar template, footer, namespace                            */
/* ------------------------------------------------------------------ */

const PREFIX = "FmtTpl";
const newId = (id: string) => `${PREFIX}${id.replace(/[^A-Za-z0-9]/g, "")}`;

function styleMap(styles: string): Map<string, string> {
    const out = new Map<string, string>();
    for (const m of styles.matchAll(/<w:style\b[^>]*>/g)) {
        const id = /w:styleId="([^"]+)"/.exec(m[0])?.[1];
        if (id) out.set(id, styles.slice(m.index!, elementEnd(styles, m.index!, "w:style")));
    }
    return out;
}

/** Salin style yang dipakai template (berikut induknya) dengan id baru; Normal → FmtTplNormal mandiri. */
async function importStyles(target: JSZip, tpl: LoadedTemplate, used: Set<string>): Promise<void> {
    const all = styleMap(tpl.styles);
    const need = new Set<string>();
    const add = (id: string) => {
        if (need.has(id) || id === "Normal") return;
        const s = all.get(id);
        if (!s) return;
        need.add(id);
        const based = /<w:basedOn w:val="([^"]+)"/.exec(s)?.[1];
        if (based) add(based);
    };
    used.forEach(add);

    const dd = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(tpl.styles)?.[0] ?? "";
    const ddR = /<w:rPrDefault>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(dd)?.[1] ?? null;
    const ddP = /<w:pPrDefault>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)/.exec(dd)?.[1] ?? null;
    const normal = all.get("Normal") ?? "";
    const kidsOf = (el: string | null) => (el ? splitElements(innerXml(el)).map((c) => c.xml) : []);
    const tagOf = (x: string) => /^<([\w:]+)/.exec(x)![1];
    // Nilai dasar eksplisit supaya docDefaults naskah (spasi 1,08, Calibri, …) tidak ikut terbawa.
    const pBase = new Map<string, string>([
        ["w:widowControl", `<w:widowControl w:val="0"/>`],
        ["w:spacing", `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`],
        ["w:ind", `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`],
        ["w:contextualSpacing", `<w:contextualSpacing w:val="0"/>`],
        ["w:jc", `<w:jc w:val="left"/>`],
    ]);
    for (const x of [...kidsOf(ddP), ...kidsOf(childOf(normal, "w:pPr"))]) pBase.set(tagOf(x), x);
    const rBase = new Map<string, string>();
    for (const x of [...kidsOf(ddR), ...kidsOf(childOf(normal, "w:rPr"))]) rBase.set(tagOf(x), x);
    const pKids = mergeProps(null, Object.fromEntries(pBase), PPR_ORDER);
    const rKids = mergeProps(null, Object.fromEntries(rBase), RPR_ORDER);
    const out: string[] = [
        `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${PREFIX}Normal"><w:name w:val="${PREFIX} Normal"/><w:qFormat/>` +
        `<w:pPr>${pKids.join("")}</w:pPr><w:rPr>${rKids.join("")}</w:rPr></w:style>`,
    ];
    for (const id of need) {
        let s = all.get(id)!;
        const type = /w:type="(\w+)"/.exec(s)?.[1] ?? "paragraph";
        s = s.replace(/\sw:default="1"/, "")
            .replace(/w:styleId="[^"]+"/, `w:styleId="${newId(id)}"`)
            .replace(/<w:name w:val="([^"]*)"\/>/, (_m, n: string) => `<w:name w:val="${PREFIX} ${n}"/>`)
            .replace(/<w:basedOn w:val="([^"]+)"\/>/, (_m, b: string) => `<w:basedOn w:val="${b === "Normal" ? `${PREFIX}Normal` : newId(b)}"/>`)
            .replace(/<w:link w:val="[^"]*"\/>|<w:rsid w:val="[^"]*"\/>|<w:autoRedefine\/>/g, "")
            .replace(/<w:next w:val="([^"]+)"\/>/, (_m, n: string) => (type === "paragraph" ? `<w:next w:val="${n === "Normal" || !need.has(n) ? `${PREFIX}Normal` : newId(n)}"/>` : ""));
        // Style paragraf tanpa induk: berdiri di atas FmtTplNormal.
        if (type === "paragraph" && !/<w:basedOn\b/.test(s)) s = s.replace(/(<w:name\b[^>]*\/>)/, `$1<w:basedOn w:val="${PREFIX}Normal"/>`);
        out.push(s);
    }

    const file = target.file("word/styles.xml");
    if (!file) return;
    let xml = await file.async("string");
    // Format ulang dokumen yang sama: style template lama diganti.
    xml = xml.replace(new RegExp(`<w:style\\b[^>]*w:styleId="${PREFIX}[^"]*"[^>]*>[\\s\\S]*?<\\/w:style>`, "g"), "");
    target.file("word/styles.xml", xml.replace(/<\/w:styles>\s*$/, `${out.join("")}</w:styles>`));
}

function childOf(xml: string, tag: string): string | null {
    const i = xml.search(new RegExp(`<${tag}[\\s>]`));
    return i === -1 ? null : xml.slice(i, elementEnd(xml, i, tag));
}

/** Pakai style template: id diganti, paragraf/tabel tanpa style diberi FmtTplNormal/FmtTplTableNormal. */
function applyStyles(xml: string, used: Set<string>): string {
    xml = xml.replace(/<w:(pStyle|rStyle|tblStyle) w:val="([^"]+)"\/>/g, (_m, t: string, id: string) => {
        used.add(id);
        return `<w:${t} w:val="${id === "Normal" ? `${PREFIX}Normal` : newId(id)}"/>`;
    });
    xml = xml.replace(/<w:p>(?!<w:pPr>)/g, `<w:p><w:pPr><w:pStyle w:val="${PREFIX}Normal"/></w:pPr>`)
        .replace(/<w:p\/>/g, `<w:p><w:pPr><w:pStyle w:val="${PREFIX}Normal"/></w:pPr></w:p>`)
        .replace(/<w:pPr>(?!<w:pStyle\b)/g, `<w:pPr><w:pStyle w:val="${PREFIX}Normal"/>`)
        .replace(/<w:tblPr>(?!<w:tblStyle\b)/g, () => { used.add("TableNormal"); return `<w:tblPr><w:tblStyle w:val="${newId("TableNormal")}"/>`; });
    return xml;
}

const IMAGE_TYPES: Record<string, string> = {
    png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
    tif: "image/tiff", tiff: "image/tiff", emf: "image/x-emf", wmf: "image/x-wmf", svg: "image/svg+xml",
};
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

async function ensureDefaultType(zip: JSZip, ext: string, type: string) {
    const p = "[Content_Types].xml";
    const xml = await zip.file(p)!.async("string");
    if (new RegExp(`<Default\\s+Extension="${ext}"`, "i").test(xml)) return;
    zip.file(p, xml.replace(/<Types\b[^>]*>/, (t) => `${t}<Default Extension="${ext}" ContentType="${type}"/>`));
}

/** Gabung & normalkan path di dalam paket ("word" + "../media/x.png" → "media/x.png"); tanpa modul path Node. */
function posixJoin(base: string, rel: string): string {
    const parts = (rel.startsWith("/") ? rel.slice(1) : `${base}/${rel}`).split("/");
    const out: string[] = [];
    for (const p of parts) {
        if (!p || p === ".") continue;
        if (p === "..") out.pop(); else out.push(p);
    }
    return out.join("/");
}

/** Gambar template (logo) disalin ke paket naskah; r:embed diarahkan ke relasi baru. */
async function importMedia(target: JSZip, tpl: LoadedTemplate, xml: string, warn: (m: string) => void): Promise<string> {
    const ids = [...new Set([...xml.matchAll(/\br:(?:embed|id|link)="([^"]+)"/g)].map((m) => m[1]))];
    const map = new Map<string, string>();
    for (const id of ids) {
        const rel = tpl.rels.get(id);
        if (!rel || rel.type !== IMAGE_REL || rel.mode === "External") { warn(`Objek template dengan relasi ${id} tidak didukung dan dilewati.`); continue; }
        const src = posixJoin("word", rel.target);
        const data = await tpl.zip.file(src)?.async("uint8array");
        if (!data) continue;
        const ext = (/\.(\w+)$/.exec(src)?.[1] ?? "png").toLowerCase();
        const name = `media/fmttpl_${src.split("/").pop()}`;
        target.file(`word/${name}`, data);
        await ensureDefaultType(target, ext, IMAGE_TYPES[ext] ?? "application/octet-stream");
        map.set(id, await findOrAddRelationship(target, name, IMAGE_REL));
    }
    return xml.replace(/\b(r:(?:embed|id|link))="([^"]+)"/g, (m, a: string, id: string) => (map.has(id) ? `${a}="${map.get(id)}"` : m));
}

/** Footer template disalin apa adanya (nomor halaman romawi di tengah, atau kosong). */
async function importFooter(target: JSZip, tpl: LoadedTemplate, rId: string, n: number): Promise<string | null> {
    const rel = tpl.rels.get(rId);
    if (!rel) return null;
    let xml = await tpl.zip.file(posixJoin("word", rel.target))?.async("string");
    if (!xml) return null;
    xml = applyStyles(clean(xml), new Set());
    const name = `footerFmtTpl${n}.xml`;
    target.file(`word/${name}`, xml);
    await addContentTypeOverride(target, `/word/${name}`, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
    return findOrAddRelationship(target, name, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer");
}

async function blankHeader(target: JSZip): Promise<string> {
    const name = "headerFmtTplBlank.xml";
    target.file(`word/${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p/></w:hdr>`);
    await addContentTypeOverride(target, `/word/${name}`, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
    return findOrAddRelationship(target, name, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header");
}

/** Deklarasi namespace yang dipakai XML sisipan tetapi belum ada di elemen root naskah. */
export function missingNamespaces(rootTag: string, xml: string, sources: string[]): string {
    const declared = new Set([...rootTag.matchAll(/\sxmlns:([\w.-]+)=/g)].map((m) => m[1]));
    const used = new Set<string>();
    for (const m of xml.matchAll(/<\/?([A-Za-z][\w.-]*):[\w.-]+/g)) used.add(m[1]);
    for (const m of xml.matchAll(/\s([A-Za-z][\w.-]*):[\w.-]+="/g)) used.add(m[1]);
    const add: string[] = [];
    for (const p of used) {
        if (p === "xml" || p === "xmlns" || declared.has(p)) continue;
        if (new RegExp(`\\sxmlns:${p}="`).test(xml)) continue; // dideklarasikan di elemennya sendiri
        const uri = sources.map((s) => new RegExp(`\\sxmlns:${p}="([^"]+)"`).exec(s)?.[1]).find(Boolean);
        if (uri) add.push(` xmlns:${p}="${uri}"`);
    }
    return add.join("");
}

/* ------------------------------------------------------------------ */
/* Penyusunan                                                           */
/* ------------------------------------------------------------------ */

export type FillInput = {
    fields: FrontFields;
    /** Baris berformat setiap paragraf naskah (paragraf bisa beberapa baris Shift+Enter). */
    linesOf: (p: ParaNode) => LineItem[];
    kataPengantar: ParaNode[] | null;
    abstrak: { paras: ParaNode[]; keywords: ParaNode | null } | null;
    abstract: { paras: ParaNode[]; keywords: ParaNode | null } | null;
    signatures: Partial<Record<"pembimbingUtama" | "pembimbingAnggota" | "ketuaSidang" | "dekan", Pic>>;
    pernyataanPics: Pic[];
    kataPengantarPics: Pic[];
};

export type FillResult = {
    xml: string;
    tocRows: { text: string; bookmark: string }[];
    /** Data yang tidak ada di template (contohnya tidak ditemukan) atau diisi titik-titik. */
    placeholders: FrontFieldKey[];
    /** Tanda tangan/materai naskah yang dipakai & yang tidak mendapat tempat. */
    usedPics: Pic[];
    unusedPics: Pic[];
    warnings: string[];
    pageTexts: { kind: FrontPageKind; lines: string[] }[];
};

export type FillContext = {
    zip: JSZip;
    /** Tag root w:document naskah (diperbarui bila perlu namespace tambahan). */
    rootTag: string;
    /** Id docPr & bookmark terbesar yang sudah dipakai naskah. */
    maxDocPr: number;
    maxBookmark: number;
    tocField: boolean;
};

const SIGNER_KEYS = ["pembimbingUtama", "pembimbingAnggota", "ketuaSidang", "dekan"] as const;
const PROSA: FrontPageKind[] = ["kataPengantar", "abstrak", "abstract"];

function pPrInfo(pPr: string) {
    const sp = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    const ind = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    const num = (tag: string, s: string) => Number(new RegExp(`w:${tag}="(-?\\d+)"`).exec(s)?.[1] ?? 0);
    return {
        line: /w:line="(\d+)"/.test(sp) ? num("line", sp) : 240,
        rule: /w:lineRule="(\w+)"/.exec(sp)?.[1] ?? "auto",
        before: num("before", sp),
        after: num("after", sp),
        left: num("left", ind) || num("start", ind),
        right: num("right", ind) || num("end", ind),
        hasStyle: /<w:pStyle w:val="(?!FmtTplNormal)/.test(pPr),
    };
}

function lineHeight(pPr: string, halfPts: number): number {
    const i = pPrInfo(pPr);
    const natural = (halfPts / 2) * 20 * 1.15;
    if (i.rule === "exact") return i.line;
    if (i.rule === "atLeast") return Math.max(i.line, natural);
    return natural * (i.line / 240);
}

const szOf = (rPr: string) => Number(/<w:sz w:val="(\d+)"/.exec(rPr)?.[1] ?? 24);

export async function fillTemplate(tpl: LoadedTemplate, spec: FrontSpec, input: FillInput, ctx: FillContext): Promise<FillResult> {
    const warnings: string[] = [];
    const warn = (m: string) => { if (!warnings.includes(m)) warnings.push(m); };
    const f = input.fields;
    const placeholders = new Set<FrontFieldKey>();
    const usedPics: Pic[] = [];
    const tocRows: FillResult["tocRows"] = [];
    let bm = ctx.maxBookmark + 1;

    // Nilai setiap data (null → titik-titik).
    const valueOf = (k: FrontFieldKey): Rich => {
        const v = f[k];
        if (v === null || (typeof v === "string" && !v.trim())) {
            placeholders.add(k);
            return [{ t: spec.placeholder, it: false, sub: false, sup: false, b: false }];
        }
        return typeof v === "string" ? [{ t: v, it: false, sub: false, sup: false, b: false }] : v;
    };
    const samples = (Object.entries(spec.sample) as [FrontFieldKey, string][]).filter(([, v]) => v && v.trim());
    const foundSample = new Set<FrontFieldKey>();

    const outPages: string[] = [];
    const pageTexts: FillResult["pageTexts"] = [];
    const usedStyles = new Set<string>();

    for (const page of tpl.pages) {
        const fixedImages = spec.fixedImagePages.includes(page.kind);

        /* --- 1. gambar contoh (tanda tangan/materai) dikeluarkan; gambar template (logo) disalin --- */
        const imgSlots: ImgSlot[] = [];
        let pageSrc = page.xml;
        if (!fixedImages) {
            const b0 = parseBlocks(pageSrc, 0);
            paragraphsOf(b0).forEach((p, i) => {
                if (!/<w:drawing|<w:pict|mc:AlternateContent/.test(p.xml)) return;
                const m = parsePara(p.xml);
                for (const part of m.parts) {
                    if (part.k !== "run") continue;
                    part.run.atoms = part.run.atoms.filter((a) => {
                        if (a.k !== "x" || !a.img) return true;
                        const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(a.xml);
                        imgSlots.push({ page: page.kind, para: i, cx: Number(ext?.[1] ?? 0), cy: Number(ext?.[2] ?? 0), xml: a.xml });
                        return false;
                    });
                }
                p.xml = serializePara(m);
            });
            pageSrc = b0.map(serializeBlock).join("");
        }
        // Relasi gambar template → relasi baru di naskah (sebelum gambar naskah disisipkan).
        pageSrc = await importMedia(ctx.zip, tpl, pageSrc, warn);

        const blocks = parseBlocks(pageSrc, 0);
        const paras = paragraphsOf(blocks);
        const topLevel = new Set(blocks.filter((b): b is ParaNode => b.kind === "paragraph"));
        const models = new Map<ParaNode, PModel>(paras.map((p) => [p, parsePara(p.xml)]));
        const slotParas = new Map<FrontFieldKey, ParaNode>();
        const delta = new Map<ParaNode, number>();

        /* --- 2. isi prosa diganti paragraf naskah --- */
        let bodyRange: [number, number] | null = null;
        let keywordsPara: ParaNode | null = null;
        if (PROSA.includes(page.kind)) {
            const headRe = page.kind === "kataPengantar" ? KP_RE : page.kind === "abstrak" ? ABSTRAK_RE : ABSTRACT_RE;
            const label = page.kind === "abstrak" ? spec.keywordLabels.abstrak : page.kind === "abstract" ? spec.keywordLabels.abstract : null;
            const list = blocks.filter((b): b is ParaNode => b.kind === "paragraph");
            const h = list.findIndex((p) => headRe.test(bare(p.text)));
            let e = h + 1;
            const isKw = (p: ParaNode) => Boolean(label && normOf(p.text).startsWith(normOf(label)));
            while (e < list.length && list[e].text.trim() && !isKw(list[e])) e++;
            if (label) keywordsPara = list.slice(e).find(isKw) ?? null;
            const bodyTpl = list.slice(h + 1, e);
            if (!bodyTpl.length) throw new TemplateError(`halaman template "${page.kind}" tidak memiliki paragraf isi contoh.`);
            bodyRange = [blocks.indexOf(bodyTpl[0]), blocks.indexOf(bodyTpl[bodyTpl.length - 1])];

            // Judul halaman: bookmark (daftar isi berkolom) + penanda TC (daftar isi field Word).
            const head = list[h];
            const id = bm++;
            const name = `_FmtFront${id}`;
            const hm = models.get(head)!;
            const headText = textOfModel(hm).replace(/\s+/g, " ").trim().toUpperCase();
            const tc = ctx.tocField
                ? `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TC "${escapeXml(headText)}" \\f c \\l 1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
                : "";
            // Daftar isi field Word (TOC dengan sakelar o dan f): judul cukup masuk lewat TC, bukan lewat tingkat
            // outline style judul template — supaya tidak tercantum dua kali.
            if (ctx.tocField) hm.pPr = `<w:pPr>${mergeProps(hm.pPr || null, { "w:outlineLvl": `<w:outlineLvl w:val="9"/>` }, PPR_ORDER).join("")}</w:pPr>`;
            hm.parts.unshift({ k: "other", xml: `<w:bookmarkStart w:id="${id}" w:name="${name}"/>` });
            hm.parts.push({ k: "other", xml: `<w:bookmarkEnd w:id="${id}"/>${tc}` });
            tocRows.push({ text: headText, bookmark: name });
        }

        /* --- 3. teks contoh diganti data --- */
        const bodySet = new Set<BlockNode>(bodyRange ? blocks.slice(bodyRange[0], bodyRange[1] + 1) : []);
        for (const p of paras) {
            if (bodySet.has(p)) continue;
            const m = models.get(p)!;
            const text = textOfModel(m);
            if (!text.trim()) continue;
            const cands: { k: FrontFieldKey; s: number; e: number }[] = [];
            for (const [k, sample] of samples) for (const [s, e] of findSpans(text, sample)) cands.push({ k, s, e });
            cands.sort((a, b) => (b.e - b.s) - (a.e - a.s));
            const chosen: typeof cands = [];
            for (const c of cands) if (!chosen.some((x) => c.s < x.e && x.s < c.e)) chosen.push(c);
            if (!chosen.length) continue;
            const before = piecesOf(m);
            const runsAt = charRuns(m);
            const halfPts = szOf(runsAt[chosen[0].s]?.rPr ?? "");
            for (const c of chosen.sort((a, b) => b.s - a.s)) {
                foundSample.add(c.k);
                slotParas.set(c.k, p);
                // Bentuk huruf mengikuti contoh: kapital semua / miring semua.
                const span = [...text.slice(c.s, c.e)];
                const its = span.map((ch, i) => ({ ch, it: isOn(runsAt[c.s + i]?.rPr ?? "", "i") })).filter((x) => /\p{L}/u.test(x.ch));
                const upper = its.some((x) => !x.it && /\p{Lu}/u.test(x.ch)) && !its.some((x) => !x.it && /\p{Ll}/u.test(x.ch));
                const allItalic = its.length > 0 && its.every((x) => x.it);
                const noItalic = its.length > 0 && its.every((x) => !x.it);
                let value = valueOf(c.k);
                // Naskah yang memiringkan seluruh judul: itu gaya, bukan penanda nama ilmiah.
                const letters = value.reduce((a, s) => a + s.t.replace(/[^\p{L}]/gu, "").length, 0);
                const itLetters = value.reduce((a, s) => a + (s.it ? s.t.replace(/[^\p{L}]/gu, "").length : 0), 0);
                if (letters && itLetters / letters > 0.8) value = value.map((s) => ({ ...s, it: false }));
                if (allItalic) value = value.map((s) => ({ ...s, it: true }));
                else if (noItalic) value = value.map((s) => ({ ...s, it: false }));
                if (upper) value = upperRich(value);
                let end = c.e;
                // Nama penanda tangan yang di template disambung jabatannya ("…, IPM. Pembimbing Utama"
                // terbungkus ke baris kedua): jabatan selalu di baris sendiri, berapa pun panjang nama.
                if ((SIGNER_KEYS as readonly string[]).includes(c.k) && text.slice(c.e).trim() && !/^\s*\n/.test(text.slice(c.e))) {
                    while (end < text.length && text[end] === " ") end++;
                    value = [...value, { t: "\n", it: false, sub: false, sup: false, b: false }];
                }
                replaceSpan(m, c.s, end, (base) => richRuns(base, value));
            }
            // Perubahan tinggi paragraf (judul lebih panjang/pendek dari contoh).
            if (topLevel.has(p) && !PROSA.includes(page.kind)) {
                const info = pPrInfo(m.pPr);
                const pgW = Number(/<w:pgSz\b[^>]*w:w="(\d+)"/.exec(page.sectPr)?.[1] ?? 11906);
                const ml = Number(/<w:pgMar\b[^>]*w:left="(\d+)"/.exec(page.sectPr)?.[1] ?? 1440);
                const mr = Number(/<w:pgMar\b[^>]*w:right="(\d+)"/.exec(page.sectPr)?.[1] ?? 1440);
                const width = pgW - ml - mr - info.left - info.right;
                const d = (measureLines(piecesOf(m), width) - measureLines(before, width)) * lineHeight(m.pPr, halfPts);
                if (d) delta.set(p, d);
            }
        }

        /* --- 4. tanda tangan/materai naskah --- */
        if (!fixedImages) {
            const put = (para: ParaNode, drawing: string) => {
                const m = models.get(para)!;
                m.parts.unshift({ k: "run", run: { rPr: "<w:rPr><w:noProof/></w:rPr>", atoms: [{ k: "x", xml: drawing, img: true }] } });
            };
            if (page.kind === "pengesahan") {
                const maxCx = spec.signature.maxWidthCm * CM_EMU;
                const maxCy = spec.signature.maxHeightCm * CM_EMU;
                for (const k of SIGNER_KEYS) {
                    const pic = input.signatures[k];
                    const para = slotParas.get(k);
                    if (!pic || !para) continue;
                    // Ditambatkan di ruang kosong tanda tangan di atas nama (kolom yang sama), rata bawah.
                    const space = spaceAbove(blocks, para);
                    const d = space
                        ? floatingSignature(pic, maxCx, Math.min(maxCy, Math.max(0, space.height - 60) * 635), (cy) => Math.max(0, space.height * 635 - cy - 60 * 635))
                        : floatingSignature(pic, maxCx, maxCy, (cy) => -(cy + 36000));
                    if (d) { put(space?.para ?? para, d); usedPics.push(pic); }
                }
            } else {
                const pics = page.kind === "pernyataan" ? [...input.pernyataanPics] : page.kind === "kataPengantar" ? [...input.kataPengantarPics] : [];
                // Pasangkan menurut bentuk: persegi (materai) ↔ persegi, lebar (tanda tangan) ↔ lebar.
                const ratio = (cx: number, cy: number) => Math.log(Math.max(1, cx) / Math.max(1, cy));
                const slots = [...imgSlots];
                while (pics.length && slots.length) {
                    let best: [number, number, number] | null = null;
                    pics.forEach((pc, i) => slots.forEach((sl, j) => {
                        const cost = Math.abs(ratio(pc.cx, pc.cy) - ratio(sl.cx, sl.cy));
                        if (!best || cost < best[0]) best = [cost, i, j];
                    }));
                    const [, i, j] = best!;
                    const d = inTemplateFrame(slots[j], pics[i]);
                    if (d) { put(paras[slots[j].para], d); usedPics.push(pics[i]); }
                    pics.splice(i, 1);
                    slots.splice(j, 1);
                }
            }
        }

        // Model → XML.
        for (const p of paras) p.xml = serializePara(models.get(p)!);

        /* --- 5. tinggi diimbangi baris kosong sesudahnya --- */
        const list = blocks.filter((b): b is ParaNode => b.kind === "paragraph" && topLevel.has(b));
        let leftover = 0;
        for (const [p, d0] of delta) {
            let d = d0;
            const at = list.indexOf(p);
            const next = list[at + 1];
            if (d < 0) {
                if (next) next.xml = bumpBefore(next.xml, -d);
                continue;
            }
            for (const q of list.slice(at + 1)) {
                if (d <= 0) break;
                if (q.text.trim() || /<w:drawing|<w:pict|mc:AlternateContent/.test(q.xml)) continue;
                const m = parsePara(q.xml);
                const lh = lineHeight(m.pPr, szOf(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m.pPr)?.[0] ?? ""));
                const info = pPrInfo(m.pPr);
                const have = lh + info.before + info.after;
                const take = Math.min(d, have - 20);
                if (take <= 0) continue;
                const left = have - take;
                q.xml = setSpacing(q.xml, { before: 0, after: 0, line: Math.max(20, Math.round(left)), rule: "exact" });
                d -= take;
            }
            leftover += Math.max(0, d);
        }
        // Sisa tambahan tinggi masih tertampung ruang kosong di bawah halaman template?
        if (leftover > 40 && leftover > pageSlack(page.xml, page.sectPr) - 120) {
            warn(`${pageLabel(page.kind)}: judul jauh lebih panjang dari contoh template; isi halaman ini mungkin meluber ke halaman berikutnya. Periksa di Word (persingkat judul bila perlu).`);
        }

        /* --- 6. isi prosa --- */
        let pageXml: string;
        if (bodyRange) {
            const bodyTpl = blocks[bodyRange[0]] as ParaNode;
            const tplModel = parsePara(bodyTpl.xml);
            const baseRPr = tplModel.parts.find((x) => x.k === "run")?.k === "run"
                ? (tplModel.parts.find((x) => x.k === "run") as { k: "run"; run: Run }).run.rPr : "";
            const src = page.kind === "kataPengantar" ? input.kataPengantar : page.kind === "abstrak" ? input.abstrak?.paras ?? null : input.abstract?.paras ?? null;
            const made: string[] = [];
            for (const up of src ?? []) {
                const lines = input.linesOf(up);
                if (!lines.length) continue;
                const value: Rich = [];
                lines.forEach((l, i) => {
                    if (i) value.push({ t: "\n", it: false, sub: false, sup: false, b: false });
                    value.push(...l.segs.map((s) => ({ ...s, t: i === 0 ? s.t : s.t })));
                });
                // Tab/spasi ketikan di awal paragraf dibuang; indentasi dari template.
                if (value[0]) value[0] = { ...value[0], t: value[0].t.replace(/^[\s\t]+/, "") };
                const num = /<w:numPr>[\s\S]*?<\/w:numPr>/.exec(up.xml.slice(0, Math.max(0, up.xml.indexOf("</w:pPr>"))))?.[0];
                let pPr = tplModel.pPr || "<w:pPr></w:pPr>";
                if (num) {
                    const first = pPrInfo(pPr);
                    const fl = Number(/w:firstLine="(\d+)"/.exec(pPr)?.[1] ?? 0);
                    pPr = `<w:pPr>${mergeProps(pPr, { "w:numPr": num, "w:ind": `<w:ind w:left="${first.left + fl + 360}" w:hanging="360"/>` }, PPR_ORDER).join("")}</w:pPr>`;
                }
                made.push(serializePara({ open: "<w:p>", pPr, parts: richRuns(baseRPr, value.filter((s) => s.t)).map((run) => ({ k: "run" as const, run })) }));
            }
            if (!made.length) made.push(serializePara({ open: "<w:p>", pPr: tplModel.pPr, parts: [] }));

            // Kata kunci: label template tetap, isinya dari naskah.
            if (keywordsPara) {
                const label = page.kind === "abstrak" ? spec.keywordLabels.abstrak : spec.keywordLabels.abstract;
                const m = parsePara(keywordsPara.xml);
                const text = textOfModel(m);
                const { norm, at } = normIndex(text);
                const nl = normOf(label);
                let cut = norm.startsWith(nl) ? at[nl.length - 1] + 1 : 0;
                while (cut < text.length && /[\s:\uFF1A]/.test(text[cut])) cut++;
                const kwSrc = page.kind === "abstrak" ? input.abstrak?.keywords : input.abstract?.keywords;
                let value: Rich = [];
                if (kwSrc) {
                    const lines = input.linesOf(kwSrc);
                    const all: Rich = lines.flatMap((l, i) => (i ? [{ t: " ", it: false, sub: false, sup: false, b: false }, ...l.segs] : l.segs));
                    const full = all.map((s) => s.t).join("");
                    const lab = /^\s*(?:kata\s*kunci|key\s*words?)\s*[:\uFF1A]\s*/i.exec(full)?.[0].length ?? 0;
                    value = cutRich(all, lab);
                }
                const base = charRuns(m)[cut]?.rPr ?? "";
                replaceSpan(m, cut, text.length, () => richRuns(base.replace(/<w:b\/>|<w:bCs\/>/g, ""), value));
                keywordsPara.xml = serializePara(m);
            }
            pageXml = [
                ...blocks.slice(0, bodyRange[0]).map(serializeBlock),
                ...made,
                ...blocks.slice(bodyRange[1] + 1).map(serializeBlock),
            ].join("");
        } else {
            pageXml = blocks.map(serializeBlock).join("");
        }
        outPages.push(pageXml);
        pageTexts.push({ kind: page.kind, lines: paragraphsOf(parseBlocks(pageXml, 0)).map((q) => q.text.replace(/\s+/g, " ").trim()).filter(Boolean) });
    }

    for (const [k] of samples) if (!foundSample.has(k)) warn(`Teks contoh ${FIELD_LABEL[k].toLowerCase()} ("${spec.sample[k]}") tidak ditemukan di berkas template, jadi data ini tidak tercantum. Periksa spesifikasi template.`);

    /* --- section, style, gambar template, footer --- */
    const zip = ctx.zip;
    const hdr = await blankHeader(zip);
    const footerIds = new Map<string, string>();
    let fn = 0;
    const footerFor = async (rId: string) => {
        if (!footerIds.has(rId)) {
            const id = await importFooter(zip, tpl, rId, ++fn);
            if (id) footerIds.set(rId, id);
        }
        return footerIds.get(rId) ?? null;
    };
    let xml = "";
    for (let i = 0; i < tpl.pages.length; i++) {
        const page = tpl.pages[i];
        let sect = page.sectPr.replace(/<w:(?:headerReference|footerReference)\b[^>]*\/>/g, "");
        const refs: string[] = [`<w:headerReference w:type="default" r:id="${hdr}"/>`, `<w:headerReference w:type="even" r:id="${hdr}"/>`];
        if (/<w:titlePg\b/.test(sect)) refs.push(`<w:headerReference w:type="first" r:id="${hdr}"/>`);
        for (const m of page.sectPr.matchAll(/<w:footerReference\b[^>]*w:type="(\w+)"[^>]*r:id="([^"]+)"[^>]*\/>|<w:footerReference\b[^>]*r:id="([^"]+)"[^>]*w:type="(\w+)"[^>]*\/>/g)) {
            const type = m[1] ?? m[4];
            const id = await footerFor(m[2] ?? m[3]);
            if (id) refs.push(`<w:footerReference w:type="${type}" r:id="${id}"/>`);
        }
        sect = sect.replace(/^<w:sectPr\b[^>]*>/, (t) => `${t}${refs.join("")}`);
        xml += attachSect(outPages[i], sect);
    }
    xml = applyStyles(xml, usedStyles);
    await importStyles(zip, tpl, usedStyles);

    // Id gambar unik di seluruh dokumen.
    let docPr = ctx.maxDocPr;
    xml = xml.replace(/(<wp:docPr\b[^>]*?\bid=")(\d+)"/g, (_m, a: string) => `${a}${++docPr}"`);

    return {
        xml, tocRows,
        placeholders: [...placeholders],
        usedPics,
        unusedPics: [...input.pernyataanPics, ...input.kataPengantarPics, ...Object.values(input.signatures)].filter((p): p is Pic => Boolean(p) && !usedPics.includes(p!)),
        warnings,
        pageTexts,
    };
}

/* ------------------------------------------------------------------ */
/* Perkiraan tinggi halaman template                                    */
/* ------------------------------------------------------------------ */

const twipOf = (sect: string, re: RegExp, dflt: number) => Number(re.exec(sect)?.[1] ?? dflt);

/** Tinggi paragraf (twip) pada lebar `width`. */
function paraHeight(p: ParaNode, width: number): number {
    const m = parsePara(p.xml);
    const info = pPrInfo(m.pPr);
    const firstRun = m.parts.find((x) => x.k === "run" && x.run.atoms.some((a) => a.k === "t")) as { k: "run"; run: Run } | undefined;
    const halfPts = szOf(firstRun?.run.rPr ?? /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m.pPr)?.[0] ?? "");
    const pieces = piecesOf(m);
    const lines = pieces.some((x) => x.t.trim()) ? measureLines(pieces, Math.max(1000, width - info.left - info.right)) : 1;
    let h = lines * lineHeight(m.pPr, halfPts);
    // Gambar sebaris memperbesar tinggi baris (spasi otomatis).
    for (const g of p.xml.matchAll(/<wp:inline\b[\s\S]*?<wp:extent\s+cx="\d+"\s+cy="(\d+)"/g)) {
        if (info.rule !== "exact") h = Math.max(h, Number(g[1]) / 635 + lineHeight(m.pPr, halfPts) * 0.25);
    }
    // Gambar melayang yang mendorong teks (atas-bawah) menambah tinggi.
    for (const a of p.xml.matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)) {
        if (!/<wp:wrapTopAndBottom\b/.test(a[0])) continue;
        const cy = Number(/<wp:extent\s+cx="\d+"\s+cy="(\d+)"/.exec(a[0])?.[1] ?? 0);
        const off = Number(/<wp:positionV\b[^>]*>\s*<wp:posOffset>(-?\d+)/.exec(a[0])?.[1] ?? 0);
        h += Math.max(0, (cy + off) / 635 - h);
    }
    return h + info.before + info.after;
}

function blocksHeight(blocks: BlockNode[], width: number): number {
    let h = 0;
    for (const b of blocks) {
        if (b.kind === "paragraph") h += paraHeight(b, width);
        else if (b.kind === "table") {
            const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(b.open + (b.props ?? ""))?.[0] ?? "";
            const pad = (side: string) => Number(new RegExp(`<w:tblCellMar>[\\s\\S]*?<w:${side} w:w="(\\d+)"`).exec(tblPr)?.[1] ?? (side === "left" || side === "right" ? 108 : 0));
            for (const r of b.rows) {
                let rowH = 0;
                for (const c of r.cells) {
                    const cw = Number(/<w:tcW w:w="(\d+)"/.exec(c.props ?? "")?.[1] ?? width);
                    rowH = Math.max(rowH, blocksHeight(c.children, cw - pad("left") - pad("right")));
                }
                h += rowH + pad("top") + pad("bottom");
            }
        } else if (b.kind === "container") h += blocksHeight(b.children, width);
    }
    return h;
}

/** Ruang kosong (twip) di bawah isi halaman template — perkiraan. */
function pageSlack(pageXml: string, sectPr: string): number {
    const w = twipOf(sectPr, /<w:pgSz\b[^>]*w:w="(\d+)"/, 11906);
    const hgt = twipOf(sectPr, /<w:pgSz\b[^>]*w:h="(\d+)"/, 16838);
    const avail = hgt - twipOf(sectPr, /<w:pgMar\b[^>]*w:top="(\d+)"/, 1440) - twipOf(sectPr, /<w:pgMar\b[^>]*w:bottom="(\d+)"/, 1440);
    const width = w - twipOf(sectPr, /<w:pgMar\b[^>]*w:left="(\d+)"/, 1440) - twipOf(sectPr, /<w:pgMar\b[^>]*w:right="(\d+)"/, 1440);
    return avail - blocksHeight(parseBlocks(pageXml, 0), width);
}

function pageLabel(k: FrontPageKind): string {
    return { sampulMuka: "Sampul muka", sampulKedua: "Sampul kedua", pengesahan: "Lembar pengesahan", pernyataan: "Lembar pernyataan", kataPengantar: "Kata pengantar", abstrak: "Abstrak", abstract: "Abstract" }[k];
}

/** Potong `n` karakter pertama dari teks berformat. */
function cutRich(r: Rich, n: number): Rich {
    const out: Rich = [];
    let left = n;
    for (const s of r) {
        if (left >= s.t.length) { left -= s.t.length; continue; }
        out.push({ ...s, t: s.t.slice(left) });
        left = 0;
    }
    if (out[0]) out[0] = { ...out[0], t: out[0].t.replace(/^\s+/, "") };
    return out.filter((s) => s.t);
}

function setSpacing(pXml: string, o: { before: number; after: number; line: number; rule: string }): string {
    const m = parsePara(pXml);
    const kids = mergeProps(m.pPr || null, { "w:spacing": `<w:spacing w:before="${o.before}" w:after="${o.after}" w:line="${o.line}" w:lineRule="${o.rule}"/>` }, PPR_ORDER);
    m.pPr = `<w:pPr>${kids.join("")}</w:pPr>`;
    return serializePara(m);
}

function bumpBefore(pXml: string, add: number): string {
    const m = parsePara(pXml);
    const sp = /<w:spacing\b[^>]*\/>/.exec(m.pPr)?.[0];
    const cur = Number(/w:before="(\d+)"/.exec(sp ?? "")?.[1] ?? 0);
    const next = sp
        ? (/w:before="/.test(sp) ? sp.replace(/w:before="\d+"/, `w:before="${Math.round(cur + add)}"`) : sp.replace("<w:spacing", `<w:spacing w:before="${Math.round(add)}"`))
        : `<w:spacing w:before="${Math.round(add)}"/>`;
    const kids = mergeProps(m.pPr || null, { "w:spacing": next }, PPR_ORDER);
    m.pPr = `<w:pPr>${kids.join("")}</w:pPr>`;
    return serializePara(m);
}

/** sectPr ke paragraf terakhir halaman. */
function attachSect(pageXml: string, sect: string): string {
    const els = splitElements(pageXml);
    const idx = els.map((e) => e.tag).lastIndexOf("w:p");
    if (idx === -1 || idx !== els.length - 1) return pageXml + `<w:p><w:pPr>${sect}</w:pPr></w:p>`;
    const m = parsePara(els[idx].xml);
    // Paragraf kosong pembawa section di ujung halaman dibuat setipis mungkin (tidak terlihat),
    // supaya halaman yang penuh tidak meluberkan paragraf kosong ini ke halaman baru.
    const blank = !textOfModel(m).trim() && !m.parts.some((p) => p.k === "run" && p.run.atoms.some((a) => a.k === "x"));
    if (blank) m.parts = m.parts.filter((p) => p.k !== "run");
    const tiny = blank
        ? { "w:spacing": `<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>`, "w:rPr": `<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>` }
        : {};
    m.pPr = `<w:pPr>${mergeProps(m.pPr || null, { ...tiny, "w:sectPr": sect }, PPR_ORDER).join("")}</w:pPr>`;
    els[idx] = { ...els[idx], xml: serializePara(m) };
    return els.map((e) => e.xml).join("");
}
/**
 * Garis & tampilan tabel data, dijalankan SETELAH normalisasi tabel.
 *
 * 1. Garis (preset tableBorders "minimal"). Semua aturan diturunkan dari STRUKTUR grid tabel,
 *    bukan dari kebiasaan satu penulis:
 *      ─ garis atas tabel
 *      ─ di bawah judul kolom induk yang membawahi beberapa kolom (hanya selebar anak kolomnya)
 *      ─ di bawah blok header (header bertingkat dikenali dari sel gabung vertikal ATAU dari
 *        baris yang membagi judul induk, baik pakai merge maupun sel kosong)
 *      ─ di bawah baris satuan ("……%……", "– – mg – –", "(%)", "(1) (2) (3)")
 *      ─ di atas baris ringkasan di bagian bawah (Jumlah/Total/Rata-rata/… + SD/KK di bawahnya)
 *      ─ garis bawah tabel; selain itu tanpa garis (termasuk garis tegak)
 *    Garis dihitung per TEPI di tiap kolom grid, lalu dipasang di kedua sel yang berbagi tepi
 *    (bawah sel atas + atas sel bawah). Kalau hanya satu sisi, Word membatalkannya bila sisi
 *    lain bernilai "nil" (dan style tabel seperti Plain Table/Table Grid tidak bisa ditimpa
 *    tblBorders, jadi semua dipasang per sel).
 *    Tabel tata letak (rumus bernomor, susunan gambar, kotak 1 baris/1 kolom) tidak diberi
 *    garis: border aslinya dipulihkan.
 * 2. Perataan & tebal mengikuti naskah asli: perataan yang diatur penulis dipertahankan; sel
 *    yang belum diatur memakai aturan preset (dihitung per kolom grid, jadi sel gabungan tidak
 *    salah kolom). Tebal judul kolom persis seperti aslinya (termasuk tebal dari style tabel).
 */
import { mergeProps, withProps, PPR_ORDER, RPR_ORDER, TBLPR_ORDER, TCPR_ORDER } from "./xml";
import { cellParagraphs, type CellNode, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { styleBoldOf } from "./heading-weight";

type Weight = "all" | "none" | "mixed" | "empty";
type Orig = { jc: string | null; bold: Weight };
type TableOrig = {
    cells: Map<string, Orig>;
    tblBorders: string | null;
    tcBorders: Map<string, string | null>;
    /** Teks sel seperti TAMPAK di Word: Tab bertitik penuntun ditulis "…" (mis. "[tab]mM[tab]" → "…mM…"). */
    shown: Map<string, string>;
};
export type TableLook = Map<TableNode, TableOrig>;

const B_RE = /<w:b(?:\s+w:val="([^"]*)")?\s*\/>/;
const isOn = (m: RegExpExecArray | null) => Boolean(m) && !/^(0|false|off)$/.test(m![1] ?? "");
const ownPPr = (xml: string) => {
    const e = xml.indexOf("</w:pPr>");
    return e === -1 ? "" : xml.slice(0, e).replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").replace(/<w:pPrChange\b[\s\S]*$/, "");
};
const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const textRuns = (xml: string) =>
    [...xml.slice(Math.max(0, xml.indexOf("</w:pPr>"))).matchAll(RUN_RE)].map((m) => m[0]).filter((r) => /<w:t[\s>]/.test(r));
const rPrOf = (r: string) => /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(r)?.[1] ?? "";
const BORDERS_OF = (props: string | null, tag: string) =>
    new RegExp(`<${tag}\\s*/>|<${tag}>[\\s\\S]*?</${tag}>`).exec(props ?? "")?.[0] ?? null;

/* ------------------------------------------------------------------ */
/* Teks tampak: Tab dengan titik penuntun                               */
/* ------------------------------------------------------------------ */

const TABS_OF = (xml: string) => /<w:tabs>[\s\S]*?<\/w:tabs>/.exec(xml)?.[0] ?? "";
/** Tab stop dengan titik/strip/garis penuntun (bukan "none", bukan tab yang dihapus). */
const LEADER_TAB = /<w:tab\b(?=[^>]*w:leader="(?!none")[^"]+")(?![^>]*w:val="clear")[^>]*\/>/;

/** Style paragraf (beserta rantai basedOn & style default) yang punya tab bertitik penuntun. */
function leaderStyles(stylesXml: string): (id: string | undefined) => boolean {
    const own = new Map<string, boolean>();
    const base = new Map<string, string>();
    let def: string | undefined;
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        if (!/w:type="paragraph"/.test(m[1])) continue;
        const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
        if (!id) continue;
        if (/w:default="(?:1|true|on)"/.test(m[1])) def = id;
        own.set(id, LEADER_TAB.test(TABS_OF(m[2])));
        const b = /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1];
        if (b) base.set(id, b);
    }
    return (id) => {
        for (let cur = id ?? def, n = 0; cur && n < 20; cur = base.get(cur), n++) if (own.get(cur)) return true;
        return false;
    };
}

const decodeXml = (s: string) =>
    s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function shownText(xml: string, leader: boolean): string {
    const end = xml.indexOf("</w:pPr>");
    const body = end === -1 ? xml : xml.slice(end);
    let s = "";
    for (const m of body.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\s*\/>|<w:ptab\b([^>]*)\/>|<w:(?:br|cr)\b[^>]*\/>/g)) {
        if (m[1] !== undefined) s += decodeXml(m[1]);
        else if (m[0].startsWith("<w:tab")) s += leader ? "…" : " ";
        else if (m[0].startsWith("<w:ptab")) s += /w:leader="(?!none")[^"]+"/.test(m[2] ?? "") ? "…" : " ";
        else s += " ";
    }
    return s;
}

/* ------------------------------------------------------------------ */
/* Tebal dari style tabel (format bersyarat)                            */
/* ------------------------------------------------------------------ */

type CondBold = { whole: boolean; firstRow: boolean; lastRow: boolean; firstCol: boolean; lastCol: boolean };

function tableStyleBold(stylesXml: string): Map<string, CondBold> {
    const out = new Map<string, CondBold>();
    for (const m of stylesXml.matchAll(/<w:style\b[^>]*w:type="table"[^>]*>([\s\S]*?)<\/w:style>/g)) {
        const id = /w:styleId="([^"]+)"/.exec(m[0].slice(0, m[0].indexOf(">")))?.[1];
        if (!id) continue;
        const body = m[1];
        const own = body.replace(/<w:tblStylePr\b[\s\S]*?<\/w:tblStylePr>/g, "");
        const cond = (type: string) => {
            const x = new RegExp(`<w:tblStylePr\\s+w:type="${type}"\\s*>([\\s\\S]*?)</w:tblStylePr>`).exec(body)?.[1];
            return x ? isOn(B_RE.exec(rPrOf(x))) : false;
        };
        out.set(id, {
            whole: isOn(B_RE.exec(rPrOf(own))),
            firstRow: cond("firstRow"), lastRow: cond("lastRow"), firstCol: cond("firstCol"), lastCol: cond("lastCol"),
        });
    }
    return out;
}

function lookFlags(tblPr: string) {
    const look = /<w:tblLook\b[^>]*\/>/.exec(tblPr)?.[0] ?? "";
    const v = parseInt(/w:val="([0-9A-Fa-f]{4})"/.exec(look)?.[1] ?? "0", 16);
    const flag = (attr: string, bit: number) => {
        const a = new RegExp(`w:${attr}="([^"]*)"`).exec(look)?.[1];
        return a !== undefined ? !/^(0|false|off)$/.test(a) : (v & bit) !== 0;
    };
    return { firstRow: flag("firstRow", 0x20), lastRow: flag("lastRow", 0x40), firstCol: flag("firstColumn", 0x80), lastCol: flag("lastColumn", 0x100) };
}

/** Dicatat sebelum normalisasi: perataan langsung, tebal tampilan, dan border asli. */
export function captureTableLook(nodes: (ParaNode | TableNode)[], stylesXml: string): TableLook {
    const styleBold = styleBoldOf(stylesXml);
    const styleLeader = leaderStyles(stylesXml);
    const tStyles = tableStyleBold(stylesXml);
    const out: TableLook = new Map();
    for (const t of nodes) {
        if (t.kind !== "table") continue;
        const tStyle = tStyles.get(/<w:tblStyle\s+w:val="([^"]+)"/.exec(t.props ?? "")?.[1] ?? "");
        const look = lookFlags(t.props ?? "");
        const lastRow = t.rows.length - 1;
        const cells = new Map<string, Orig>();
        const tcBorders = new Map<string, string | null>();
        const shown = new Map<string, string>();
        t.rows.forEach((row, r) => row.cells.forEach((cell, c) => {
            tcBorders.set(`${r}:${c}`, BORDERS_OF(cell.props, "w:tcBorders"));
            shown.set(`${r}:${c}`, cellParagraphs(cell).map((p) => {
                const pPr = ownPPr(p.xml);
                const leader = LEADER_TAB.test(TABS_OF(pPr)) ||
                    (!/<w:tabs>/.test(pPr) && styleLeader(/<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1]));
                return shownText(p.xml, leader);
            }).join(" ").replace(/\s+/g, " ").trim());
            const cond = Boolean(tStyle) && (
                tStyle!.whole ||
                (r === 0 && look.firstRow && tStyle!.firstRow) ||
                (r === lastRow && look.lastRow && tStyle!.lastRow) ||
                (c === 0 && look.firstCol && tStyle!.firstCol) ||
                (c === row.cells.length - 1 && look.lastCol && tStyle!.lastCol));
            cellParagraphs(cell).forEach((p, k) => {
                const pPr = ownPPr(p.xml);
                const paraBold = styleBold(/<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1]);
                const flags = textRuns(p.xml).map((run) => {
                    const rp = rPrOf(run);
                    const d = B_RE.exec(rp);
                    if (d) return isOn(d);
                    return styleBold(/<w:rStyle\s+w:val="([^"]+)"/.exec(rp)?.[1]) || paraBold || cond;
                });
                const bold: Weight = !flags.length ? "empty" : flags.every(Boolean) ? "all" : flags.some(Boolean) ? "mixed" : "none";
                cells.set(`${r}:${c}:${k}`, { jc: /<w:jc\s+w:val="(\w+)"/.exec(pPr)?.[1] ?? null, bold });
            });
        }));
        out.set(t, { cells, tblBorders: BORDERS_OF(t.props, "w:tblBorders"), tcBorders, shown });
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Grid tabel                                                           */
/* ------------------------------------------------------------------ */

type GCell = {
    cell: CellNode; r: number; c: number; start: number; span: number;
    vm: "restart" | "continue" | null;
    /** Baris teratas & terbawah dari sel gabung vertikal tempat sel ini berada. */
    top: number; bottom: number;
    text: string;
    /** Teks seperti tampak di Word (Tab bertitik penuntun → "…"), untuk mengenali baris satuan. */
    shown: string;
};

const cellText = (cell: CellNode) => cellParagraphs(cell).map((p) => p.text).join(" ").replace(/\s+/g, " ").trim();
const hasObject = (cell: CellNode) => cellParagraphs(cell).some((p) => /<w:drawing\b|<w:pict\b|<m:oMath|<w:object\b/.test(p.xml));
const overlaps = (a: GCell, b: GCell) => a.start < b.start + b.span && b.start < a.start + a.span;

function gridOf(t: TableNode, orig: TableOrig | undefined): GCell[][] {
    const grid = t.rows.map((row, r) => {
        let col = Number(/<w:gridBefore\s+w:val="(\d+)"/.exec(row.props ?? "")?.[1] ?? 0);
        return row.cells.map((cell, c): GCell => {
            const span = Math.max(1, Number(/<w:gridSpan\s+w:val="(\d+)"/.exec(cell.props ?? "")?.[1] ?? 1));
            const v = /<w:vMerge(?:\s+w:val="(\w+)")?\s*\/>/.exec(cell.props ?? "");
            const text = cellText(cell);
            const g: GCell = {
                cell, r, c, start: col, span, vm: v ? (v[1] === "restart" ? "restart" : "continue") : null,
                top: r, bottom: r, text, shown: orig?.shown.get(`${r}:${c}`) ?? text,
            };
            col += span;
            return g;
        });
    });
    for (let r = 1; r < grid.length; r++) for (const g of grid[r]) {
        if (g.vm !== "continue") continue;
        const up = grid[r - 1].find((a) => a.start === g.start && a.vm !== null);
        if (up) g.top = up.top;
    }
    const bottom = new Map<string, number>();
    for (const g of grid.flat()) bottom.set(`${g.top}:${g.start}`, Math.max(bottom.get(`${g.top}:${g.start}`) ?? g.r, g.r));
    for (const g of grid.flat()) g.bottom = bottom.get(`${g.top}:${g.start}`)!;
    return grid;
}

/* ------------------------------------------------------------------ */
/* Pengenal baris                                                       */
/* ------------------------------------------------------------------ */

const NUMERIC_RE = /^[\s(]*[-+−±<>≤≥~]?\s*\d[\d.,]*\s*(?:±\s*\d[\d.,]*)?\s*%?\s*[)*a-z]{0,3}$/i;
const DECIMAL = /\d[.,]\d/;
const FILLER = /\.{3,}|…|_{3,}|-{3,}|(?:-\s+){2,}|[–—]{2,}|(?:[–—]\s+){2,}/;

/** Sel satuan: "……%……", "– – mg/kg – –", "(%)", "%". */
function unitCell(t: string): boolean {
    if (FILLER.test(t)) {
        const core = t.replace(/[.…_\-–—·\s]+/g, " ").trim();
        return core.length <= 25 && !DECIMAL.test(core) && !/\d{3,}/.test(core);
    }
    if (/^\(\s*[^()]{1,15}\)$/.test(t)) return !/^\(\s*[\d.,\s]+\)$/.test(t);
    return /^(?:%|‰|°\s?c|ppm|ppb)$/i.test(t);
}

/**
 * Baris satuan atau baris nomor kolom "(1) (2) (3)" tepat di bawah judul kolom. Dikenali dari
 * teks yang TAMPAK (titik ketikan maupun Tab bertitik penuntun), atau dari satu satuan pendek
 * yang dibentangkan di atas beberapa kolom data (mis. "mM", "g/ekor/hari" tanpa titik).
 */
function isUnitRow(grid: GCell[][], r: number): boolean {
    const row = grid[r];
    const cells = row.filter((g) => g.shown);
    if (!cells.length) return false;
    if (cells.every((g) => /^\(\d{1,2}\)$/.test(g.shown))) return cells.length >= 2;
    if (cells.every((g) => unitCell(g.shown)) && row.map((g) => g.shown).join(" ").length <= 120) return true;
    // Satu satuan polos selebar beberapa kolom; bukan judul induk yang dibagi lagi di baris bawahnya.
    const [only] = cells;
    return cells.length === 1 && only.span >= 2 && only.text.length <= 20 &&
        /[A-Za-z%‰°µμ]/.test(only.text) && !/\d[.,]\d/.test(only.text) && !NUMERIC_RE.test(only.text) &&
        !PRIMARY.test(only.text) && !(r + 1 < grid.length && subdivides(grid, r + 1, r));
}

/** Baris r membagi judul kolom induk di atasnya (header bertingkat, dengan atau tanpa merge). */
function subdivides(grid: GCell[][], r: number, end: number): boolean {
    const row = grid[r];
    const head = grid.slice(0, end + 1).flat();
    const parents = head.filter((p) => p.span > 1 && p.text && p.bottom === r - 1);
    if (!parents.length || row.some((g) => DECIMAL.test(g.text))) return false;
    const within = (g: GCell, p: GCell) => g.start >= p.start && g.start + g.span <= p.start + p.span;
    if (!parents.some((p) => { const kids = row.filter((g) => within(g, p)); return kids.length >= 2 && kids.some((k) => k.text); })) return false;
    const outside = row.filter((g) => !parents.some((p) => within(g, p)));
    if (!outside.length) return row.every((g) => !g.text || !NUMERIC_RE.test(g.text));
    // Kolom di luar judul induk: diteruskan merge, kosong, atau judulnya memang ditulis di baris ini.
    return outside.every((g) =>
        (g.vm === "continue" && g.top <= end) || !g.text || head.filter((h) => overlaps(h, g)).every((h) => !h.text));
}

const REPEAT_HEADER = /<w:tblHeader\b(?![^>]*w:val="(?:0|false|off)")[^>]*\/>/;

/** Indeks baris header terakhir (-1 bila tabel tidak punya baris judul kolom). */
function headerEnd(t: TableNode, grid: GCell[][]): number {
    const last = grid.length - 1;
    if (grid[0].every((g) => !g.text || NUMERIC_RE.test(g.text))) return -1;
    let end = 0;
    for (let r = 1; r < last; r++) {
        const row = grid[r];
        if (isUnitRow(grid, r)) break;
        const noData = !row.some((g) => DECIMAL.test(g.text));
        const merged = noData && row.some((g) => g.vm === "continue" && g.top <= end);
        const repeated = noData && REPEAT_HEADER.test(t.rows[r].props ?? "") &&
            !row.some((g) => g.text && NUMERIC_RE.test(g.text));
        if (!merged && !repeated && !subdivides(grid, r, end)) break;
        end = r;
    }
    return end;
}

const PRIMARY = /^(?:jumlah|total|sub\s*-?\s*total|grand\s+total|rata[-\s]?rata|rerata|rataan|mean|average|sum)\b/i;
const SECONDARY = /^(?:sd|se|sem|sb|simpangan\s+baku|standar\s+deviasi|standard\s+(?:deviation|error)|galat\s+baku|kk|cv|koefisien\s+keragaman|bnt|bnj|lsd|hsd|p[\s-]?value|nilai\s+p)\b/i;

/** "Jumlah", "Total (%)", "Rata-rata ± SD" ya; "Total fenolik" (nama data) tidak. */
const isLabel = (re: RegExp, s: string) => {
    const m = re.exec(s);
    if (!m) return false;
    const rest = s.slice(m[0].length).trim();
    return !rest || /^[:±(\-–=]/.test(rest) || /^(?:total|keseluruhan|umum|akhir)\b/i.test(rest);
};

/** Baris teratas dari blok ringkasan di dasar tabel (-1 bila tidak ada). */
function summaryTop(grid: GCell[][], firstBody: number): number {
    let top = -1;
    let primary = false;
    for (let r = grid.length - 1; r > firstBody; r--) {
        const first = grid[r].map((g) => g.text).find(Boolean) ?? "";
        if (!first) { if (top === -1) continue; break; }
        if (isLabel(PRIMARY, first)) { primary = true; top = r; continue; }
        if (isLabel(SECONDARY, first)) { top = r; continue; }
        break;
    }
    return primary ? top : -1;
}

function captionBefore(nodes: (ParaNode | TableNode)[], outline: Outline, i: number): boolean {
    for (let j = i - 1, seen = 0; j >= 0 && seen < 3; j--) {
        const n = nodes[j];
        if (n.kind === "table") return false;
        if (!n.text.trim() && !n.hasDrawing) continue;
        if (outline.items[j]?.role === "table_caption") return true;
        seen++;
    }
    return false;
}

/* ------------------------------------------------------------------ */

const LINE = (sz: number) => `w:val="single" w:sz="${sz}" w:space="0" w:color="000000"`;
const NIL = `w:val="nil"`;

function setRunsBold(xml: string, on: boolean): string {
    const b = on ? "<w:b/>" : `<w:b w:val="0"/>`;
    const bCs = on ? "<w:bCs/>" : `<w:bCs w:val="0"/>`;
    const end = xml.indexOf("</w:pPr>");
    const head = end === -1 ? "" : xml.slice(0, end);
    const body = end === -1 ? xml : xml.slice(end);
    return head + body.replace(RUN_RE, (r) => (/<w:t[\s>]/.test(r) ? withProps(r, "w:rPr", { "w:b": b, "w:bCs": bCs }, RPR_ORDER) : r));
}

const setTcPr = (g: GCell, patch: Record<string, string | null>) => {
    g.cell.props = `<w:tcPr>${mergeProps(g.cell.props, patch, TCPR_ORDER).join("")}</w:tcPr>`;
};
const setJc = (p: ParaNode, jc: string) => {
    p.xml = withProps(p.xml, "w:pPr", { "w:jc": `<w:jc w:val="${jc}"/>` }, PPR_ORDER);
};
const normJc = (jc: string) => jc.replace(/^start$/, "left").replace(/^end$/, "right");

export function applyTableLines(
    nodes: (ParaNode | TableNode)[],
    outline: Outline,
    look: TableLook,
    opts: { lines: boolean; numericAlign: "center" | "right"; skip: Set<number>; borderSize?: number },
): { tables: number } {
    const sz = opts.borderSize ?? 4;
    let tables = 0;
    nodes.forEach((t, i) => {
        if (t.kind !== "table" || opts.skip.has(t.id)) return;
        const zone = outline.items[i]?.zone;
        if (zone !== "main" && zone !== "back") return;
        const orig = look.get(t);
        const grid = gridOf(t, orig);
        const all = grid.flat();
        const cols = Math.max(0, ...all.map((g) => g.start + g.span));

        // Tabel tata letak: kembalikan border & perataan asli, tanpa aturan tabel data.
        const layout = grid.length < 2 || cols < 2 ||
            (!captionBefore(nodes, outline, i) && all.some((g) => hasObject(g.cell)));
        if (layout) {
            if (opts.lines && orig) {
                t.props = `<w:tblPr>${mergeProps(t.props, { "w:tblBorders": orig.tblBorders }, TBLPR_ORDER).join("")}</w:tblPr>`;
                for (const g of all) setTcPr(g, { "w:tcBorders": orig.tcBorders.get(`${g.r}:${g.c}`) ?? null });
            }
            for (const g of all) cellParagraphs(g.cell).forEach((p, k) => {
                const o = orig?.cells.get(`${g.r}:${g.c}:${k}`);
                if (o?.jc) setJc(p, normJc(o.jc));
            });
            return;
        }

        const last = grid.length - 1;
        const hEnd = headerEnd(t, grid);
        const unit = hEnd >= 0 && hEnd + 1 < last && isUnitRow(grid, hEnd + 1) ? hEnd + 1 : -1;
        const headLast = unit >= 0 ? unit : hEnd;
        const firstBody = headLast + 1;
        const summary = summaryTop(grid, firstBody);

        if (opts.lines) {
            // line[b][c]: ada garis di tepi atas baris b (b = jumlah baris → tepi bawah tabel) pada kolom grid c.
            const line = Array.from({ length: grid.length + 1 }, () => new Array<boolean>(cols).fill(false));
            const full = (b: number) => line[b].fill(true);
            full(0);
            full(grid.length);
            if (hEnd >= 0) {
                for (const p of grid.slice(0, hEnd + 1).flat()) {
                    if (p.span > 1 && p.text && p.r === p.top && p.bottom < hEnd) {
                        for (let c = p.start; c < p.start + p.span; c++) line[p.bottom + 1][c] = true;
                    }
                }
                full(hEnd + 1);
            }
            if (unit >= 0) full(unit + 1);
            if (summary >= 0) full(summary);
            // Tidak ada garis yang memotong sel gabung vertikal.
            for (const g of all) {
                if (g.r !== g.top || g.bottom === g.r) continue;
                for (let b = g.r + 1; b <= g.bottom; b++) for (let c = g.start; c < g.start + g.span; c++) line[b][c] = false;
            }
            const edge = (b: number, g: GCell) => line[b].slice(g.start, g.start + g.span).some(Boolean);
            for (const g of all) {
                setTcPr(g, {
                    "w:tcBorders": `<w:tcBorders><w:top ${edge(g.r, g) ? LINE(sz) : NIL}/><w:left ${NIL}/>` +
                        `<w:bottom ${edge(g.r + 1, g) ? LINE(sz) : NIL}/><w:right ${NIL}/>` +
                        `<w:insideH ${NIL}/><w:insideV ${NIL}/></w:tcBorders>`,
                });
            }
            const tb = `<w:tblBorders><w:top ${LINE(sz)}/><w:left ${NIL}/><w:bottom ${LINE(sz)}/><w:right ${NIL}/>` +
                `<w:insideH ${NIL}/><w:insideV ${NIL}/></w:tblBorders>`;
            t.props = `<w:tblPr>${mergeProps(t.props, { "w:tblBorders": tb }, TBLPR_ORDER).join("")}</w:tblPr>`;
        }

        // Kolom angka dihitung per kolom grid (bukan urutan sel), dari baris isi.
        const colNumeric = Array.from({ length: cols }, (_, c) => {
            const vals = grid.slice(firstBody).flat().filter((g) => g.start === c && g.span === 1 && g.text).map((g) => g.text);
            return vals.length >= 2 && vals.filter((v) => NUMERIC_RE.test(v)).length / vals.length >= 0.6;
        });

        for (const g of all) {
            const isHead = g.r <= headLast;
            cellParagraphs(g.cell).forEach((p, k) => {
                const o = orig?.cells.get(`${g.r}:${g.c}:${k}`);
                let jc: string | null = o?.jc ? normJc(o.jc) : null;
                if (!jc && g.text) {
                    if (g.r === unit) jc = "center";
                    else if (!isHead) {
                        jc = NUMERIC_RE.test(g.text) || (g.span === 1 && colNumeric[g.start] && g.text.length <= 4)
                            ? opts.numericAlign : "left";
                    }
                }
                if (jc) setJc(p, jc);
                if (o?.bold === "all") p.xml = setRunsBold(p.xml, true);
                else if (o?.bold === "none" && isHead) p.xml = setRunsBold(p.xml, false);
            });
        }
        tables++;
    });
    return { tables };
}
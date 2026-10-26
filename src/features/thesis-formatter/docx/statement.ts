/**
 * Lembar pernyataan (dan kotak/bingkai berbentuk tabel 1 sel).
 *
 *  - Tabel 1 baris × 1 sel yang membungkus beberapa paragraf adalah BINGKAI, bukan tabel data:
 *    tidak boleh dinormalkan (garis atas-bawah, isi spasi 1 rata kiri).
 *  - Lembar pernyataan yang ditulis di dalam bingkai seperti itu dikeluarkan dari tabelnya
 *    (teks tidak berubah) agar judul, baris "Nama : …", dan tanda tangannya diformat seperti
 *    halaman lain. Kotak buatan penulis (tabel bingkai, shape persegi) dibuang; preset "box"
 *    membungkus isi halaman yang SUDAH diformat dengan satu kotak setinggi isinya (garis bawah
 *    tepat di bawah nama penulis), preset "none" tanpa kotak.
 */
import { flatten, parseBlocks, type BlockNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { startsNewPage } from "../structure/roles";
import type { PageInfo } from "./page-openers";
import { maxId } from "./figure-split";

export const STATEMENT_RE = /^(?:(?:lembar|surat|halaman)\s+)?pernyataan\b/i;
const DRAWING_FLAG_RE = /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]|<v:shape[\s>]/;

const textsOf = (list: BlockNode[]) =>
    list.filter((c): c is ParaNode => c.kind === "paragraph").map((p) => p.text.trim()).filter(Boolean);

/** Tabel 1×1 berisi beberapa paragraf (atau satu paragraf panjang) = bingkai. */
export function isFrameTable(t: TableNode): boolean {
    if (t.rows.length !== 1 || t.rows[0].cells.length !== 1) return false;
    const texts = textsOf(t.rows[0].cells[0].children);
    return texts.length >= 3 || texts.some((x) => x.length >= 150);
}

export function findFrameTables(model: DocModel): Set<number> {
    const out = new Set<number>();
    for (const n of flatten(model.body)) if (n.kind === "table" && isFrameTable(n)) out.add(n.id);
    return out;
}

const isStatementHead = (t: string) => t.length <= 120 && STATEMENT_RE.test(t.replace(/\s+/g, " ").trim());

/**
 * Bingkai tabel yang memuat lembar pernyataan (judulnya di dalam bingkai, atau tepat di atasnya)
 * diganti isinya. Dipanggil sebelum outline dibangun.
 */
export function unwrapStatementFrames(model: DocModel): { count: number; tables: number; rows: number; cells: number } {
    const out = { count: 0, tables: 0, rows: 0, cells: 0 };
    const visit = (list: BlockNode[]) => {
        for (let i = 0; i < list.length; i++) {
            const n = list[i];
            if (n.kind === "container") { visit(n.children); continue; }
            if (n.kind !== "table" || !isFrameTable(n)) continue;
            const inner = n.rows[0].cells[0].children;
            const first = textsOf(inner)[0] ?? "";
            let prev = "";
            for (let k = i - 1; k >= 0; k--) {
                const m = list[k];
                if (m.kind !== "paragraph") break;
                if (m.text.trim()) { prev = m.text.trim(); break; }
                if (m.hasSectPr || m.hasPageBreak) break;
            }
            if (!isStatementHead(first) && !isStatementHead(prev)) continue;
            list.splice(i, 1, ...inner);
            i += inner.length - 1;
            out.count++;
            out.tables++;
            out.rows++;
            out.cells++;
        }
    };
    visit(model.body);
    return out;
}

export type StatementRange = { pageStart: number; heading: number; end: number };

/** Halaman pernyataan: dari awal halamannya sampai section break / judul halaman berikutnya. */
export function statementRanges(nodes: (ParaNode | TableNode)[], outline: Outline): StatementRange[] {
    const out: StatementRange[] = [];
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (n.kind !== "paragraph" || !it || it.role !== "front_heading" || !isStatementHead(n.text)) return;
        let s = i;
        while (s > 0) {
            const m = nodes[s - 1];
            if (m.kind !== "paragraph" || m.text.trim() || m.hasSectPr || m.hasPageBreak) break;
            s--;
        }
        let e = i + 1;
        for (; e < nodes.length; e++) {
            const m = nodes[e];
            const r = outline.items[e]?.role;
            if (r && startsNewPage(r)) break;
            if (m.kind === "paragraph" && (m.hasSectPr || m.hasPageBreak)) { e++; break; }
        }
        out.push({ pageStart: s, heading: i, end: e });
    });
    return out;
}

const RUN_RE = /<w:r\b(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;

/** Ukuran gambar/shape (twips) dari wp:extent atau gaya VML. */
function sizeOf(run: string): { w: number; h: number } | null {
    const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(run);
    if (ext) return { w: Number(ext[1]) / 635, h: Number(ext[2]) / 635 };
    const style = /<v:(?:rect|roundrect|shape)\b[^>]*\sstyle="([^"]*)"/.exec(run)?.[1] ?? "";
    const pt = (k: string) => Number(new RegExp(`${k}:\\s*([\\d.]+)pt`).exec(style)?.[1] ?? 0) * 20;
    return style ? { w: pt("width"), h: pt("height") } : null;
}

/**
 * Shape persegi panjang besar tanpa teks & tanpa foto di halaman pernyataan = bingkai buatan
 * penulis; dihapus (diganti bingkai seragam atau dihilangkan sesuai preset).
 */
export function removeFrameShapes(
    nodes: (ParaNode | TableNode)[],
    range: StatementRange,
    area: { width: number; height: number },
): number {
    let removed = 0;
    for (let k = range.pageStart; k < range.end; k++) {
        const p = nodes[k];
        if (p.kind !== "paragraph" || !p.hasDrawing) continue;
        const xml = p.xml.replace(RUN_RE, (run) => {
            const shape = /<wps:wsp\b|<v:rect\b|<v:roundrect\b/.test(run);
            if (!shape || /<pic:pic\b|<wps:txbx\b|<v:textbox\b|<w:txbxContent\b/.test(run)) return run;
            if (/<a:prstGeom\s+prst="(?!rect|roundRect)/.test(run)) return run;
            const s = sizeOf(run);
            if (!s || s.w < area.width * 0.6 || s.h < area.height * 0.4) return run;
            removed++;
            return "";
        });
        if (xml !== p.xml) {
            p.xml = xml;
            p.hasDrawing = DRAWING_FLAG_RE.test(xml);
        }
    }
    return removed;
}

const PBB_RE = /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/;
const BOX_LINE = `w:val="single" w:sz="12" w:space="0" w:color="000000"`;
/** Jarak teks ke garis kotak: kiri/kanan (lebar teks tetap; kotak menjorok ke margin) & atas/bawah. */
const PAD_X = 113;
const PAD_Y = 170;
/** Tinggi tambahan kotak (garis + jarak dalam atas & bawah), untuk pengukuran halaman. */
export const STATEMENT_BOX_EXTRA = PAD_Y * 2 + 60;

/**
 * Isi lembar pernyataan (judul sampai baris terakhir berisi, biasanya nama penulis) dibungkus satu
 * kotak bergaris di keempat sisi. Tinggi kotak mengikuti isi; lebar teks di dalamnya sama dengan
 * lebar teks halaman sehingga pemotongan baris tidak berubah. Dipanggil PALING AKHIR.
 */
export function boxStatementPages(model: DocModel, info: PageInfo, opts: { width: number }): number {
    let boxed = 0;
    let id = maxId(model.body) + 1;
    const visit = (list: BlockNode[]) => {
        for (let i = 0; i < list.length; i++) {
            const n = list[i];
            if (n.kind === "container") { visit(n.children); continue; }
            if (n.kind !== "paragraph" || info.get(n)?.role !== "front_heading" || !isStatementHead(n.text)) continue;
            let end = i + 1;
            for (; end < list.length; end++) {
                const m = list[end];
                if (m.kind === "container") break;
                if (m.kind === "raw") continue;          // bookmark dsb. ikut masuk kotak
                const r = info.get(m)?.role;
                if (r && startsNewPage(r)) break;
                if (m.kind === "paragraph" && m.hasSectPr) break;
                if (m.kind === "paragraph" && m.hasPageBreak) { if (m.text.trim()) end++; break; }
            }
            while (end > i + 1) {
                const m = list[end - 1];
                if (m.kind !== "paragraph" || m.text.trim() || m.hasDrawing) break;
                end--;
            }
            const moved = list.slice(i, end);
            // Pindah halaman sebelum judul tidak berlaku di dalam tabel: dipindah ke paragraf penanda.
            const pEnd = n.xml.indexOf("</w:pPr>");
            const pageBreak = pEnd !== -1 && PBB_RE.test(n.xml.slice(0, pEnd));
            if (pageBreak) n.xml = n.xml.replace(PBB_RE, "");
            if (moved[moved.length - 1].kind !== "paragraph") moved.push(parseBlocks("<w:p/>", id++)[0]);
            const w = opts.width + PAD_X * 2;
            const side = (k: string) => `<w:${k} ${BOX_LINE}/>`;
            const tbl = parseBlocks(
                `<w:tbl><w:tblPr><w:tblW w:w="${w}" w:type="dxa"/><w:tblInd w:w="${-PAD_X}" w:type="dxa"/>` +
                `<w:tblBorders>${["top", "left", "bottom", "right"].map(side).join("")}<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>` +
                `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="${PAD_Y}" w:type="dxa"/><w:left w:w="${PAD_X}" w:type="dxa"/>` +
                `<w:bottom w:w="${PAD_Y}" w:type="dxa"/><w:right w:w="${PAD_X}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>` +
                `<w:tblGrid><w:gridCol w:w="${w}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
                id,
            )[0] as TableNode;
            id += 10;
            tbl.rows[0].cells[0].children = moved;
            const lead = pageBreak
                ? [parseBlocks(`<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`, id++)[0]]
                : [];
            list.splice(i, end - i, ...lead, tbl);
            i += lead.length;
            boxed++;
        }
    };
    visit(model.body);
    return boxed;
}
/**
 * Kelompok tata letak yang saling bergantung — tidak boleh diubah sebagian.
 *
 * Banyak penulis menyusun gambar dengan objek MELAYANG yang posisinya dihitung relatif terhadap
 * paragraf di sekitarnya, misalnya tabel 2 kolom tanpa garis berisi: foto melayang → beberapa
 * paragraf kosong sebagai pengganjal → kotak teks melayang berisi keterangan "Ilustrasi 1. …"
 * (+1 cm di bawah paragraf tertentu) → baris "Sumber". Susunan ini hanya rapi selama SEMUA
 * jaraknya sama dengan aslinya. Mengubah sebagian (foto dijadikan inline, spasi pengganjal
 * diseragamkan, lebar sel diubah) membuat keterangan jatuh menimpa teks atau terpotong sel.
 *
 * Prinsip: susunan seperti ini dikenali dari CIRINYA (ada objek melayang / kotak teks / shape /
 * frame), lalu dibiarkan utuh dan dilaporkan.
 */
import { serializeBlock, type BlockNode, type DocModel, type ParaNode, type TableNode } from "./document-model";

/** Objek yang posisinya tidak ikut alur teks (atau kotak teks/shape yang biasanya ditata manual). */
const FLOATING_RE = /<wp:anchor\b|<w:txbxContent\b|<wps:wsp\b|<wpg:wgp\b|<w:framePr\b|<v:shape\b[^>]*position:\s*absolute|<v:textbox\b/;

export const hasFloating = (xml: string): boolean => FLOATING_RE.test(xml);

/** Kotak teks / shape melayang (bukan foto biasa). */
const SHAPE_FLOAT_RE = /<w:txbxContent\b|<wps:wsp\b|<wpg:wgp\b|<v:textbox\b|<v:rect\b|<v:line\b|<v:oval\b|<v:group\b/;
export const hasFloatingShape = (xml: string): boolean => SHAPE_FLOAT_RE.test(xml);

export type TableClass = "data" | "layout" | "complex";

/**
 * - complex: berisi objek melayang, kotak teks, shape, frame, atau tabel bersarang → tidak disentuh.
 * - layout : tanpa garis sama sekali (tabel penata posisi) → tidak diberi garis/format data.
 * - data   : selebihnya.
 * Ragu → complex (aman: tidak diubah).
 */
export function tableClass(t: TableNode): TableClass {
    const xml = serializeBlock(t);
    if (hasFloating(xml)) return "complex";
    for (const r of t.rows) for (const c of r.cells) if (c.children.some((x) => x.kind === "table" || (x.kind === "container" && /<w:tbl\b/.test(serializeBlock(x))))) return "complex";
    if (borderlessTable(t)) return "layout";
    // Tabel penata gambar (foto berdampingan + keterangan + sumber): sedikit baris, berisi gambar,
    // tidak ada sel berisi angka/data. Bukan tabel data → tidak diberi garis tabel.
    if (t.rows.length <= 4 && /<pic:pic\b/.test(xml) && !/<w:tbl\b[\s\S]*<w:tbl\b/.test(xml.slice(6))) return "layout";
    return "data";
}

/**
 * Lebar tabel tata letak yang melebihi area teks (margin baru lebih sempit) diperkecil proporsional:
 * gridCol & tcW (dxa) diskalakan bersama; struktur (gridSpan, vMerge) tidak diubah.
 */
export function fitTableWidth(t: TableNode, usable: number): boolean {
    const cols = [...(t.grid ?? "").matchAll(/<w:gridCol\s+w:w="(\d+)"\s*\/>/g)].map((m) => Number(m[1]));
    const total = cols.reduce((a, b) => a + b, 0);
    if (!total || total <= usable) return false;
    const k = usable / total;
    t.grid = (t.grid ?? "").replace(/<w:gridCol\s+w:w="(\d+)"\s*\/>/g, (_m, w: string) => `<w:gridCol w:w="${Math.floor(Number(w) * k)}"/>`);
    for (const r of t.rows) for (const c of r.cells) {
        if (c.props) c.props = c.props.replace(/<w:tcW\s+w:w="(\d+)"\s+w:type="dxa"\s*\/>/, (_m, w: string) => `<w:tcW w:w="${Math.floor(Number(w) * k)}" w:type="dxa"/>`);
    }
    if (t.props) t.props = t.props.replace(/<w:tblW\s+w:w="(\d+)"\s+w:type="dxa"\s*\/>/, (_m, w: string) => `<w:tblW w:w="${Math.floor(Number(w) * k)}" w:type="dxa"/>`);
    return true;
}

/** Tabel tanpa satu garis pun (tblBorders & tcBorders nil/none). */
export function borderlessTable(t: { props: string | null; rows: { cells: { props: string | null }[] }[] }): boolean {
    const tb = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/.exec(t.props ?? "")?.[1];
    if (tb === undefined) return false;
    const sides = [...tb.matchAll(/<w:(?:top|left|bottom|right|start|end|insideH|insideV)\b[^>]*>/g)].map((m) => m[0]);
    if (sides.length < 4 || sides.some((s) => !/w:val="(?:nil|none)"/.test(s))) return false;
    return t.rows.every((r) => r.cells.every((c) => {
        const cb = /<w:tcBorders>([\s\S]*?)<\/w:tcBorders>/.exec(c.props ?? "")?.[1] ?? "";
        return [...cb.matchAll(/w:val="([^"]+)"/g)].every((m) => m[1] === "nil" || m[1] === "none");
    }));
}

/** Semua tabel kelas complex (termasuk yang bersarang & di dalam container). */
export function complexTables(model: DocModel): Set<TableNode> {
    const out = new Set<TableNode>();
    const walk = (list: BlockNode[], inComplex: boolean) => {
        for (const n of list) {
            if (n.kind === "container") walk(n.children, inComplex);
            else if (n.kind === "table") {
                const c = inComplex || tableClass(n) === "complex";
                if (c) out.add(n);
                for (const r of n.rows) for (const cell of r.cells) walk(cell.children, c);
            }
        }
    };
    walk(model.body, false);
    return out;
}

/** Paragraf yang berada di dalam tabel complex (tidak boleh diubah formatnya). */
export function paragraphsInComplexTables(model: DocModel): Set<ParaNode> {
    const out = new Set<ParaNode>();
    const collect = (list: BlockNode[]) => {
        for (const n of list) {
            if (n.kind === "paragraph") out.add(n);
            else if (n.kind === "container") collect(n.children);
            else if (n.kind === "table") for (const r of n.rows) for (const c of r.cells) collect(c.children);
        }
    };
    for (const t of complexTables(model)) for (const r of t.rows) for (const c of r.cells) collect(c.children);
    return out;
}

/**
 * Di badan teks: paragraf ini satu susunan dengan kotak teks/shape melayang di dekatnya (≤ `span`
 * paragraf, di daftar yang sama), mis. foto melayang + kotak keterangan melayang di bawahnya.
 */
export function nearFloatingShape(list: BlockNode[], index: number, span = 3): boolean {
    for (let k = Math.max(0, index - span); k <= Math.min(list.length - 1, index + span); k++) {
        const n = list[k];
        if (n.kind === "paragraph" && hasFloatingShape(n.xml)) return true;
    }
    return false;
}
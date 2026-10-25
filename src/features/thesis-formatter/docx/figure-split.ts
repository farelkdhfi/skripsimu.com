/**
 * Paragraf yang memuat gambar DAN keterangannya sekaligus ("[grafik] Ilustrasi 2. …" atau
 * "Ilustrasi 2. … [grafik]") dipecah menjadi paragraf gambar + paragraf keterangan, supaya
 * keterangannya dikenali (nomor otomatis, Daftar Gambar/Ilustrasi). Gambar melayang di
 * paragraf seperti itu dijadikan inline: posisinya memang milik keterangan tersebut.
 * Teks tidak berubah; hanya batas paragrafnya.
 */
import { innerXml, openTag, splitElements, withProps, PPR_ORDER, type XmlEl } from "./xml";
import { parseBlocks, type BlockNode, type DocModel, type ParaNode } from "./document-model";
import { paragraphLine } from "./verify";

const FIG_CAPTION_RE = /^(?:gambar|gbr\.?|ilustrasi|grafik|bagan|diagram|figure|fig\.?)\s*\d+(?:[.\-]\d+)*\s*[.:]?\s*\S/i;
const DRAWING_RE = /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/;
const PAGE_BR_RE = /<w:br\b[^>]*w:type="page"/;

const isDrawingRun = (e: XmlEl) => e.tag === "w:r" && DRAWING_RE.test(e.xml);
/** Run tanpa teks yang hanya berisi ganti baris — sisa pemisah gambar & keterangan. */
const isBreakRun = (e: XmlEl) =>
    e.tag === "w:r" && !/<w:t[\s>]/.test(e.xml) && /<w:br\b|<w:cr\b/.test(e.xml) && !PAGE_BR_RE.test(e.xml);
const lineOf = (els: XmlEl[]) => paragraphLine(els.map((e) => e.xml).join(""));

function trimBreaks(els: XmlEl[]): XmlEl[] {
    let a = 0;
    let b = els.length;
    while (a < b && isBreakRun(els[a])) a++;
    while (b > a && isBreakRun(els[b - 1])) b--;
    return els.slice(a, b);
}

function maxId(list: BlockNode[]): number {
    let m = 0;
    const walk = (xs: BlockNode[]) => {
        for (const n of xs) {
            if (n.kind === "paragraph") m = Math.max(m, n.id);
            else if (n.kind === "table") {
                m = Math.max(m, n.id);
                for (const r of n.rows) for (const c of r.cells) walk(c.children);
            } else if (n.kind === "container") walk(n.children);
        }
    };
    walk(list);
    return m;
}

/** Gambar melayang → inline (ukuran & grafik sama). Bentuk mc:AlternateContent dibiarkan. */
function inlineOf(run: string): string {
    if (/<mc:AlternateContent\b/.test(run)) return run;
    return run.replace(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g, (anchor) => {
        const pick = (re: RegExp) => re.exec(anchor)?.[0] ?? "";
        const extent = pick(/<wp:extent\b[^>]*\/>/);
        const graphic = pick(/<a:graphic\b[\s\S]*?<\/a:graphic>/);
        if (!extent || !graphic) return anchor;
        return `<wp:inline distT="0" distB="0" distL="0" distR="0">${extent}` +
            pick(/<wp:effectExtent\b[^>]*\/>/) +
            pick(/<wp:docPr\b[^>]*\/>|<wp:docPr\b[^>]*>[\s\S]*?<\/wp:docPr>/) +
            pick(/<wp:cNvGraphicFramePr\s*\/>|<wp:cNvGraphicFramePr>[\s\S]*?<\/wp:cNvGraphicFramePr>/) +
            `${graphic}</wp:inline>`;
    });
}

export function splitFigureCaptions(model: DocModel): number {
    let id = maxId(model.body) + 1;
    let count = 0;

    const visit = (list: BlockNode[]) => {
        for (let i = 0; i < list.length; i++) {
            const n = list[i];
            if (n.kind === "container") { visit(n.children); continue; }
            if (n.kind !== "paragraph" || !n.hasDrawing || !n.text.trim()) continue;

            const els = splitElements(innerXml(n.xml));
            const pPr = els[0]?.tag === "w:pPr" ? els[0].xml : "";
            const rest = pPr ? els.slice(1) : els;
            const drawings = rest.filter(isDrawingRun);
            if (!drawings.length || drawings.some((e) => /<w:t[\s>]/.test(e.xml))) continue;

            const first = rest.indexOf(drawings[0]);
            const last = rest.indexOf(drawings[drawings.length - 1]);
            const before = lineOf(rest.slice(0, first));
            const between = lineOf(rest.slice(first, last + 1));
            const after = lineOf(rest.slice(last + 1));
            // Teks hanya di satu sisi gambar, dan teks itu keterangan gambar bernomor.
            if (between || (before && after)) continue;
            if (!FIG_CAPTION_RE.test(before || after)) continue;

            const others = trimBreaks(rest.filter((e) => !isDrawingRun(e)));
            const figXml = withProps(
                `<w:p>${pPr}${drawings.map((e) => inlineOf(e.xml)).join("")}</w:p>`,
                "w:pPr",
                {
                    "w:pStyle": null, "w:numPr": null, "w:outlineLvl": null, "w:sectPr": null,
                    "w:keepNext": "<w:keepNext/>", "w:jc": `<w:jc w:val="center"/>`,
                },
                PPR_ORDER,
            );
            // Gambar didahulukan; section break & keepNext tetap di paragraf terakhir (keterangan).
            const capXml = withProps(
                `${openTag(n.xml)}${pPr}${others.map((e) => e.xml).join("")}</w:p>`,
                "w:pPr",
                { "w:pageBreakBefore": null },
                PPR_ORDER,
            );

            const fig = parseBlocks(figXml, id++)[0] as ParaNode;
            const cap = parseBlocks(capXml, 0)[0] as ParaNode;
            Object.assign(n, { ...cap, id: n.id });
            list.splice(i, 0, fig);
            i++;
            count++;
        }
    };

    visit(model.body);
    return count;
}
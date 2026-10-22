/**
 * Jalur utama untuk input DOCX: dokumen asli DISUNTING DI TEMPAT, bukan
 * dibangun ulang.
 */

import JSZip from "jszip";
import { elementEnd, withProps, PPR_ORDER } from "./xml";
import {
    flatten, markStyleNumbering, parseDocument, serializeDocument, tableText,
    type BlockNode, type ParaNode, type TableNode,
} from "./document-model";
import { applyFormat, collapseEmptyPages, type ApplyReport } from "./apply-format";
import {
    createBlankPart, createPageNumberFooter, enableUpdateFields, inspectHeaders,
    isLandscape, patchCompat, patchSectPr, patchStyles, type SectionPlan,
} from "./package-parts";
import { fitDrawings, type ImageReport } from "./images";
import { arrangeTableCaptions, detectTableNotes, type CaptionReport } from "./tables";
import { convertTypedLists } from "./lists";
import { buildTocFields, type TocReport } from "./toc";
import { fingerprint, verify, type DocFingerprint, type VerifyIssue } from "./verify";
import { buildOutline, outlineStats, type Outline } from "../structure/outline";
import { createThesisRules, type FormatRules } from "../rules/format-rules";
import { isHeadingRole, startsNewPage, type Role, type Zone } from "../structure/roles";
import type { FormatConfig } from "../schemas/format-config";
import { layoutCover } from "./cover";
import { buildSignatureTables, type SigReport } from "./signature";
import { normalizeHeadings, type HeadingReport } from "./headings";
import { promoteUnnumberedCaptions } from "./tables";
import { refreshListCaches } from "./toc-cache";

export type TransformResult = {
    buffer: Buffer;
    outline: Outline;
    stats: ReturnType<typeof outlineStats>;
    report: ApplyReport;
    issues: VerifyIssue[];
    notes: string[];
    sections: number;
    landscapeSections: number;
    portraitConverted: number;
    images: ImageReport;
    captions: CaptionReport;
    lists: number;
    toc: TocReport;
    before: DocFingerprint;
    after: DocFingerprint;
    preview: { role: string; zone: string; text: string }[];
    signatures: SigReport;
    headings: HeadingReport;
};

export type TransformOptions = { aiRoles?: (Role | null)[] };

type SectionRef = { where: "para"; node: ParaNode } | { where: "body" };

type SectionInfo = {
    ref: SectionRef;
    firstZone: Zone | null;
    firstRole: Role | null;
    coverOnly: boolean;
    nodes: (ParaNode | TableNode)[];
};

const hasContent = (n: ParaNode | TableNode) =>
    n.kind === "table" || Boolean(n.text.trim()) || n.hasDrawing;

function ensureMainSectionBreak(nodes: (ParaNode | TableNode)[], outline: Outline, bodySectPr: string | null): boolean {
    const start = outline.mainStart;
    if (start <= 0 || start >= nodes.length) return false;
    for (let k = start - 1; k >= 0; k--) {
        const n = nodes[k];
        if (n.kind === "paragraph" && n.hasSectPr) return false;
        if (hasContent(n)) break;
    }
    let k = start - 1;
    while (k >= 0 && nodes[k].kind !== "paragraph") k--;
    if (k < 0) return false;
    const target = nodes[k] as ParaNode;
    let template = bodySectPr;
    for (let j = start; j < nodes.length; j++) {
        const n = nodes[j];
        if (n.kind === "paragraph" && n.hasSectPr) {
            template = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? template;
            break;
        }
    }
    if (!template) return false;
    target.xml = withProps(target.xml, "w:pPr", { "w:sectPr": template }, PPR_ORDER);
    target.hasSectPr = true;
    return true;
}

function collectSections(body: BlockNode[], zoneOf: Map<number, Zone>, roleOf: Map<number, Role>): SectionInfo[] {
    const out: SectionInfo[] = [];
    let firstZone: Zone | null = null;
    let firstRole: Role | null = null;
    let allCover = true;
    let any = false;
    let list: (ParaNode | TableNode)[] = [];

    const close = (ref: SectionRef) => {
        out.push({ ref, firstZone, firstRole, coverOnly: any && allCover, nodes: list });
        firstZone = null; firstRole = null; allCover = true; any = false; list = [];
    };

    for (const n of flatten(body)) {
        list.push(n);
        if (hasContent(n)) {
            const z = zoneOf.get(n.id) ?? "main";
            if (firstZone === null) { firstZone = z; firstRole = roleOf.get(n.id) ?? "body"; }
            if (z !== "cover") allCover = false;
            any = true;
        }
        if (n.kind === "paragraph" && n.hasSectPr) close({ where: "para", node: n });
    }
    close({ where: "body" });
    return out;
}

function planNumbering(sections: SectionInfo[], rules: FormatRules) {
    let romanStarted = false;
    let arabicStarted = false;
    return sections.map((s) => {
        if (s.coverOnly) return { fmt: rules.numbering("cover"), start: 1 as number | null, cover: true, hideFirst: false };
        const zone = s.firstZone ?? "main";
        const fmt = rules.numbering(zone);
        const front = zone === "cover" || zone === "front";
        let start: number | null = null;
        if (front && !romanStarted) { start = 1; romanStarted = true; }
        if (!front && !arabicStarted) { start = 1; arabicStarted = true; romanStarted = true; }
        return { fmt, start, cover: false, hideFirst: s.firstZone === "cover" };
    });
}

/**
 * Tabel di section landscape cukup ditampilkan portrait? Ya bila ≤ 8 kolom,
 * kata terpanjang tiap kolom muat, dan estimasi lebar nyaman tidak berlebihan.
 */
function fitsPortrait(nodes: (ParaNode | TableNode)[], width: number, halfPts: number): boolean {
    const charTw = halfPts * 5;
    const pad = 170;
    let anyTable = false;
    for (const n of nodes) {
        if (n.kind === "paragraph") {
            if (n.hasDrawing) {
                const cx = /<wp:extent\s+cx="(\d+)"/.exec(n.xml);
                if (cx && Number(cx[1]) / 635 > width * 1.05) return false;
            }
            continue;
        }
        anyTable = true;
        const rows = tableText(n);
        const cols = Math.max(0, ...rows.map((r) => r.length));
        if (cols > 8) return false;
        let minW = 0;
        let comfy = 0;
        for (let j = 0; j < cols; j++) {
            const cells = rows.map((r) => r[j] ?? "");
            const word = Math.max(1, ...cells.map((c) => Math.max(0, ...c.split(/\s+/).map((w) => w.length))));
            const len = Math.max(1, ...cells.map((c) => c.length));
            minW += word * charTw + pad;
            comfy += Math.min(len, 30) * charTw + pad;
        }
        if (minW > width || comfy > width * 2) return false;
    }
    return anyTable;
}

export async function transformDocx(
    input: Buffer | ArrayBuffer | Uint8Array,
    config: FormatConfig,
    opts: TransformOptions = {},
): Promise<TransformResult> {
    const zip = await JSZip.loadAsync(input);
    const docFile = zip.file("word/document.xml");
    if (!docFile) throw Object.assign(new Error("File bukan DOCX yang sah (word/document.xml tidak ditemukan)."), { status: 422 });

    const originalXml = await docFile.async("string");
    const before = fingerprint(originalXml);

    const model = parseDocument(originalXml);
    const nodes = flatten(model.body);
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    if (stylesXml) markStyleNumbering(nodes, stylesXml);
    const outline = buildOutline(nodes, opts.aiRoles ?? []);
    const promoted = promoteUnnumberedCaptions(nodes, outline);
    // Catatan tabel dideteksi dari dokumen asli, sebelum daftar & ukuran font diubah.
    const tableNotes = detectTableNotes(nodes, outline, Math.round(config.fontSize * 2));
    for (const id of promoted) tableNotes.delete(id);
    const rules = createThesisRules(config);
    const notes: string[] = [];

    const page = rules.page("main");
    const usableWidth = page.widthTwips - page.margin.left - page.margin.right;

    const zoneOfNode = new Map<number, Zone>();
    const roleOfNode = new Map<number, Role>();
    nodes.forEach((n, i) => {
        zoneOfNode.set(n.id, outline.items[i]?.zone ?? "main");
        roleOfNode.set(n.id, outline.items[i]?.role ?? "body");
    });

    const numbered = config.pageNumbers !== "keep";
    if (numbered && config.romanFrontMatter) ensureMainSectionBreak(nodes, outline, model.bodySectPr);

    /* -------- daftar ketikan → numbering Word -------- */

    const runs: ParaNode[][] = [];
    let cur: ParaNode[] = [];
    nodes.forEach((n, i) => {
        const it = outline.items[i];
        if (
            n.kind === "paragraph" &&
            it?.role === "body" &&
            it.hint.listMarker &&
            it.zone !== "cover" &&
            !n.hasNumbering &&
            !tableNotes.has(n.id)
        ) {
            cur.push(n);
            return;
        }

        if (n.kind === "paragraph" && tableNotes.has(n.id)) {
            if (cur.length) {
                runs.push(cur);
                cur = [];
            }
            return;
        }

        if (n.kind === "paragraph" && it?.role === "empty") return;
        if (cur.length) { runs.push(cur); cur = []; }
    });
    if (cur.length) runs.push(cur);
    const lists = await convertTypedLists(zip, runs, config.fontFamily, Math.round(config.indent * 567));

    const report = applyFormat(model, outline, rules, {
        cleanupBlankPages: config.cleanupBlankPages,
        normalizeTables: config.normalizeTables,
    });

    /* -------- judul bab/subbab: nomor & huruf -------- */

    const headings = normalizeHeadings(nodes, outline, rules, report, { titleCase: true });

    /* -------- keterangan & catatan tabel -------- */

    const captions = arrangeTableCaptions(model, nodes, outline, rules, tableNotes);
    if (captions.moved) notes.push(`${captions.moved} keterangan tabel dipindah ke atas tabelnya.`);
    if (captions.refsUnresolved) {
        notes.push(`${captions.refsUnresolved} rujukan tabel di teks menunjuk nomor yang tidak ada di dokumen dan dibiarkan apa adanya. Periksa manual.`);
    }
    if (captions.uncaptioned.length) {
        notes.push(`${captions.uncaptioned.length} tabel di bab belum memiliki judul: ${captions.uncaptioned.slice(0, 5).join("; ")}. Tambahkan judulnya di atas tabel. Formatter tidak membuat judul sendiri.`);
    }

    /* -------- daftar isi/tabel/gambar/lampiran → field TOC -------- */

    const toc = buildTocFields(model, nodes, outline, { upper: config.chapterUppercase, width: usableWidth });

    /* -------- blok tanda tangan → tabel tanpa border -------- */

    const signatures = buildSignatureTables(model, nodes, outline, { width: usableWidth });

    /* -------- header & footer -------- */

    let footerRId: string | null = null;
    let blankFooter: string | null = null;
    let blankHeader: string | null = null;
    const headerRemap = new Map<string, string>();

    if (numbered) {
        if (config.pageNumbers !== "none") {
            footerRId = await createPageNumberFooter(
                zip, config.pageNumbers === "bottom-right" ? "right" : "center",
                config.fontFamily, Math.round(config.fontSize * 2),
            );
        }
        blankFooter = await createBlankPart(zip, "footer");
        blankHeader = await createBlankPart(zip, "header");

        const headers = await inspectHeaders(zip);
        const norm = (s: string) => s.replace(/\s+/g, " ").trim().toLowerCase();
        const headingTexts = new Set(
            nodes
                .filter((n, i) => n.kind === "paragraph" && isHeadingRole(outline.items[i]?.role ?? "body"))
                .map((n) => norm((n as ParaNode).text)),
        );
        const duplicates = headers.filter((h) => h.text && headingTexts.has(norm(h.text)));
        const pageOnly = headers.filter((h) => h.hasPageField && !h.text);
        for (const h of [...pageOnly, ...duplicates]) headerRemap.set(h.rId, blankHeader);
        for (const h of duplicates) notes.push(`Header berisi "${h.text.slice(0, 60)}" dihapus karena menduplikasi judul halaman.`);

        const docXmlForRefs = serializeDocument(model);
        const running = headers.filter((h) => h.text && !headerRemap.has(h.rId) && docXmlForRefs.includes(`r:id="${h.rId}"`));
        for (const h of running) {
            notes.push(`Header berisi teks "${h.text.slice(0, 60)}" dipertahankan. Hapus manual bila tidak disengaja — pedoman skripsi umumnya tidak memakai header berisi teks.`);
        }
    }

    /* -------- sectPr: kertas, orientasi, margin, penomoran -------- */

    const sections = collectSections(model.body, zoneOfNode, roleOfNode);
    const numbering = planNumbering(sections, rules);
    const sectXml = (s: SectionInfo) => s.ref.where === "para"
        ? /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(s.ref.node.xml)?.[0] ?? null
        : model.bodySectPr;
    const tableSize = rules.table().fontSize;
    const toPortrait = sections.map((s) => {
        const x = sectXml(s);
        return Boolean(x) && isLandscape(x!) && fitsPortrait(s.nodes, usableWidth, tableSize);
    });

    let landscapeSections = 0;
    let portraitConverted = 0;

    sections.forEach((s, i) => {
        const num = numbering[i];
        const old0 = sectXml(s);
        const stillLandscape = Boolean(old0) && isLandscape(old0!) && !toPortrait[i];
        const opensPage = !s.firstRole || startsNewPage(s.firstRole);
        // Section tabel yang dijadikan portrait menyambung ke halaman sebelumnya,
        // begitu juga isi sesudahnya — kecuali diawali judul yang memang membuka halaman.
        const continuous = !opensPage && !stillLandscape && (toPortrait[i] || (i > 0 && toPortrait[i - 1]));

        const plan: SectionPlan = {
            page: rules.page(s.firstZone ?? "main"),
            numbering: numbered && num.fmt ? { fmt: num.fmt, start: num.start } : null,
            footer: !numbered ? "keep" : num.cover || !footerRId ? { blank: blankFooter! } : { rId: footerRId },
            headerRemap,
            firstPageBlank: numbered && num.hideFirst && blankFooter && blankHeader
                ? { footer: blankFooter, header: blankHeader } : undefined,
            orientation: toPortrait[i] ? "portrait" : undefined,
            continuous: continuous || undefined,
        };
        if (stillLandscape) landscapeSections++;
        if (toPortrait[i]) portraitConverted++;

        if (s.ref.where === "para") {
            const p = s.ref.node;
            const start = p.xml.search(/<w:sectPr(?=[\s>/])/);
            if (start === -1) return;
            const end = elementEnd(p.xml, start, "w:sectPr");
            p.xml = p.xml.slice(0, start) + patchSectPr(p.xml.slice(start, end), plan) + p.xml.slice(end);
        } else if (model.bodySectPr) {
            model.bodySectPr = patchSectPr(model.bodySectPr, plan);
        }
    });

    /* -------- gambar yang melebihi area teks -------- */

    const images: ImageReport = { scaled: 0, floatingOverflow: 0 };
    {
        const landscapeWidth = page.heightTwips - page.margin.top - page.margin.bottom;
        let pending: ParaNode[] = [];
        const flush = (sect: string | null) => {
            const width = sect && isLandscape(sect) ? landscapeWidth : usableWidth;
            for (const p of pending) p.xml = fitDrawings(p.xml, width, images);
            pending = [];
        };
        for (const n of flatten(model.body)) {
            if (n.kind === "paragraph" && n.hasDrawing) pending.push(n);
            if (n.kind === "paragraph" && n.hasSectPr) flush(/<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? null);
        }
        flush(model.bodySectPr);
        if (images.floatingOverflow) {
            notes.push(`${images.floatingOverflow} gambar melayang (floating) lebih lebar dari area teks dan tidak diubah otomatis. Periksa dan perkecil manual.`);
        }
    }

    /* -------- tata letak sampul -------- */

    const cover = (() => {
        const cp = rules.page("cover");
        return layoutCover(model, nodes, outline, {
            height: cp.heightTwips - cp.margin.top - cp.margin.bottom,
            width: usableWidth,
            style: rules.cover(),
        });
    })();
    if (cover.pages) {
        notes.push(`${cover.pages} halaman sampul disusun ke urutan baku: jenis dokumen & judul di atas, logo & penulis di tengah, identitas institusi di margin bawah.`);
    }

    if (config.cleanupBlankPages) {
        const headingIds = new Set(nodes.filter((_, i) => isHeadingRole(outline.items[i]?.role ?? "body")).map((n) => n.id));
        collapseEmptyPages(model, report, headingIds);
    }

    const listCache = refreshListCaches(model, {
        dropTc: captions.appendix,
    });

    /* -------- tulis kembali paket -------- */

    const newXml = serializeDocument(model);
    zip.file("word/document.xml", newXml);

    await patchStyles(zip, config.fontFamily, Math.round(config.fontSize * 2), usableWidth);
    await enableUpdateFields(zip);
    await patchCompat(zip);

    const after = fingerprint(newXml);
    const issues = verify(before, after, {
        unorderedPrefix: cover.reordered ? cover.coverLines : 0,
        regions: [
            ...headings.regions,
            ...captions.regions,
            ...signatures.regions,
            ...listCache.regions,
        ],
        added: signatures.added,
    });

    const preview = nodes.slice(0, 300).map((n, i) => ({
        role: outline.items[i]?.role ?? "body",
        zone: outline.items[i]?.zone ?? "main",
        text: n.kind === "table"
            ? `[Tabel ${n.rows.length} baris × ${n.rows[0]?.cells.length ?? 0} kolom]`
            : n.hasDrawing && !n.text.trim() ? "[Gambar]" : n.text.replace(/\t/g, " ").slice(0, 280),
    }));

    const buffer = await zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });


    return {
        buffer,
        outline,
        stats: outlineStats(nodes, outline),
        report,
        issues,
        notes,
        sections: sections.length,
        landscapeSections,
        portraitConverted,
        images,
        lists,
        captions,
        toc,
        signatures,
        before,
        after,
        preview,
        headings
    };

}
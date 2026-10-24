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
import { detectTableNotes, formatTableNotes, spaceAfterTables, keepTablesTogether } from "./tables";
import { formatLists, applyListLayout } from "./lists";
import { buildTocFields, type TocReport } from "./toc";
import { buildOutline, outlineStats, type Outline } from "../structure/outline";
import { createThesisRules, type FormatRules } from "../rules/format-rules";
import { isHeadingRole, startsNewPage, type Role, type Zone } from "../structure/roles";
import type { FormatConfig } from "../schemas/format-config";
import { layoutCover } from "./cover";
import { buildSignatureTables, findSignatureTables, type SigReport } from "./signature";
import { normalizeHeadings, type HeadingReport } from "./headings";
import { refreshListCaches, patchListEntryStyle, demoteFromToc } from "./toc-cache";
import { realizeCaps } from "./text-case";
import { centerCoverImages } from "./cover-logo";
import { normalizeEmphasis, patchHeadingStyles } from "./emphasis";
import { formatKeywords } from "./front-text";
import { arrangeReferences, type CitationStyle } from "./references";
import { cleanFlow, patchWidowControl } from "./flow";
import { arrangeCaptions, promoteUnnumberedCaptions, hangTableCaptions, type CaptionReport } from "./captions";
import { buildCoverFromTemplate, type CoverAI } from "./cover-template";
import { detectMinorHeadings } from "./minor-headings";
import { stripHeadingDirectBold, ensureBoldHeadingStyles } from "./heading-runs";
import { layoutEquations } from "./equations";
import { unifyChapterHeadings } from "./chapter-heads";
import { alignLabelLines } from "./label-lines";
import { unindentContinuations } from "./continuation";
import { fingerprint, verify, type DocFingerprint, type VerifyIssue, type VerifyRegion } from "./verify";
import { realignStrayProse } from "./stray-prose";


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
    captions: { table: CaptionReport; figure: CaptionReport; notes: number };
    lists: number;
    toc: TocReport;
    before: DocFingerprint;
    after: DocFingerprint;
    preview: { role: string; zone: string; text: string }[];
    signatures: SigReport;
    headings: HeadingReport;
    coverRemoved: { tables: number; rows: number; cells: number };
};

export type TransformOptions = { aiRoles?: (Role | null)[]; coverAI?: CoverAI };

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
    // Judul bab → satu paragraf dua baris ("BAB I" ↵ "PENDAHULUAN" / "I" ↵ "PENDAHULUAN"),
    // sebelum outline. Peran AI disesuaikan karena paragraf judul yang digabung hilang.
    const preNodes = flatten(model.body);
    const chapters = unifyChapterHeadings(model);
    const nodes = flatten(model.body);
    const stylesXml = await zip.file("word/styles.xml")?.async("string");
    if (stylesXml) markStyleNumbering(nodes, stylesXml);
    const aiRoles = (opts.aiRoles ?? []).filter((_, i) => !chapters.removed.has(preNodes[i] as ParaNode));
    const outline = buildOutline(nodes, aiRoles, chapters.chapterIds);
    const promoted = promoteUnnumberedCaptions(nodes, outline);
    // Catatan tabel dideteksi dari dokumen asli, sebelum daftar & ukuran font diubah.
    const tableNotes = detectTableNotes(nodes, outline, Math.round(config.fontSize * 2));
    for (const id of promoted) tableNotes.delete(id);
    const rules = createThesisRules(config);
    const notes: string[] = [];
    if (chapters.merged || chapters.split) notes.push(`${chapters.merged + chapters.split} judul bab diseragamkan menjadi satu judul dua baris (nomor bab di atas, judul bab di bawah).`);

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

    /* -------- daftar (ketikan & Word asli) -------- */

    // Judul tingkat 4+ (Heading 4, "a. Rancangan …") dilindungi dari modul daftar/alur/penekanan.
    const minor = await detectMinorHeadings(zip, nodes, outline);
    const keepOut = new Set<number>([...tableNotes.keys(), ...minor.ids]);

    const listRep = await formatLists(zip, nodes, outline, {
        font: config.fontFamily,
        exclude: keepOut,
    });

    const lists = listRep.items;

    // Tabel yang isinya tanda tangan tidak dinormalkan sebagai tabel data.
    const sigTables = findSignatureTables(nodes, outline);
    const report = applyFormat(model, outline, rules, {
        cleanupBlankPages: config.cleanupBlankPages,
        normalizeTables: config.normalizeTables,
        skipTables: sigTables,
    });

    applyListLayout(nodes, listRep.layout);
    const minorRegions = minor.apply({ rep: report, font: config.fontFamily, halfPts: Math.round(config.fontSize * 2) });
    if (minor.count) notes.push(`${minor.count} judul tingkat 4 ("a. …", "1) …", Heading 4) diseragamkan: penanda di margin kiri, tebal, menempel ke paragraf di bawahnya.`);

    const flow = cleanFlow(model, nodes, outline, keepOut);
    const equations = layoutEquations(nodes, usableWidth);
    if (equations.numbered) notes.push(`${equations.numbered} rumus bernomor dirapikan: rumus di tengah, nomor persamaan rata kanan.`);
    const strayProse = realignStrayProse(nodes, outline, {
        line: Math.round(config.lineSpacing * 240),
        firstLine: Math.round(config.indent * 567),
    });
    if (strayProse) notes.push(`${strayProse} kalimat isi yang rata tengah/kanan dikembalikan menjadi alinea biasa.`);
    const continuations = unindentContinuations(nodes, outline, minor.ids);
    if (continuations) notes.push(`${continuations} paragraf lanjutan (diawali huruf kecil, mis. "dengan … adalah …" setelah rumus) dibuat tanpa indentasi karena bukan alinea baru.`);
    if (flow.breaks) notes.push(`${flow.breaks} pindah halaman manual di tengah bab dihapus.`);
    if (flow.blanks) notes.push(`${flow.blanks} baris kosong (Enter tanpa isi) yang hanya menjadi jarak dihapus; jarak antarbagian diatur formatter.`);
    if (flow.softBreaks) notes.push(`${flow.softBreaks} paragraf berisi ganti baris manual (Shift+Enter) di tengah kalimat dirapikan.`);
    if (flow.spaces) notes.push(`${flow.spaces} spasi ganda / indentasi manual dengan spasi atau tab dirapikan.`);
    if (flow.formulas) notes.push(`${flow.formulas} baris rumus dibuat rata tengah.`);
    if (flow.merged) notes.push(`${flow.merged} paragraf yang terbelah disatukan kembali.`);

    /* -------- judul bab/subbab: nomor & huruf -------- */

    const headings = normalizeHeadings(nodes, outline, rules, report, { titleCase: true });

    if (headings.chaptersSplit) notes.push(`${headings.chaptersSplit} judul bab satu baris diseragamkan menjadi dua baris ("BAB II" lalu judulnya).`);

    /* -------- keterangan & catatan tabel -------- */
    const tableCaps = arrangeCaptions(model, nodes, outline, "table");
    const figureCaps = arrangeCaptions(model, nodes, outline, "figure");
    hangTableCaptions(model, config.fontFamily);
    const tableNotesFormatted = formatTableNotes(model, rules, tableNotes);
    spaceAfterTables(model, tableNotes);
    keepTablesTogether(model, {
        usableWidth,
        pageHeight: page.heightTwips - page.margin.top - page.margin.bottom,
        halfPts: rules.table().fontSize,
    });
    const captions = { table: tableCaps, figure: figureCaps, notes: tableNotesFormatted };

    for (const c of [tableCaps, figureCaps]) {
        const noun = c.kind === "table" ? "tabel" : "gambar";
        if (c.moved) notes.push(`${c.moved} keterangan ${noun} dipindah ke ${c.kind === "table" ? "atas tabel" : "bawah gambar"}nya.`);
        if (c.split) notes.push(`${c.split} keterangan ${noun} yang memuat penjelasan/sumber dipisah: judul tetap di keterangan (masuk Daftar ${c.kind === "table" ? "Tabel" : "Gambar"}), penjelasan & "Sumber:" di baris bawahnya.`);
        if (c.uncaptioned.length)
            notes.push(`${c.uncaptioned.length} ${noun} di bab belum memiliki keterangan: ${c.uncaptioned.slice(0, 5).join("; ")}. Tambahkan judulnya ${c.kind === "table" ? "di atas tabel" : "di bawah gambar"}. Formatter tidak membuat judul sendiri.`);
        if (c.unresolved.length)
            notes.push(`Rujukan ${c.unresolved.join(", ")} di teks tidak menemukan ${noun}nya di dokumen dan dibiarkan apa adanya.`);
        if (c.skipped.length)
            notes.push(`Penomoran ${noun} di Bab ${c.skipped.join(", ")} tidak diubah karena ada keterangan ${noun} yang formatnya tidak dikenali. Periksa keterangan ${noun} di bab tersebut.`);
    }

    const emphasis = normalizeEmphasis(nodes, outline, keepOut);
    if (emphasis.paragraphs || emphasis.listItems) {
        notes.push(`Penekanan yang tidak pada tempatnya dirapikan: ${emphasis.paragraphs} paragraf isi yang seluruhnya miring/bergaris bawah, ${emphasis.listItems} butir daftar diseragamkan. Penekanan pada istilah tertentu dipertahankan.`);
    }

    const keywords = formatKeywords(nodes, outline);
    const labels = alignLabelLines(nodes, outline, { halfPts: Math.round(config.fontSize * 2) });
    if (labels.lines) notes.push(`${labels.lines} baris "Label : isi" (identitas, kop surat, singkatan) disejajarkan titik duanya.`);
    const refs = arrangeReferences(model, nodes, outline, config.citationStyle as CitationStyle);
    if (refs.sorted) notes.push(`Daftar pustaka (${refs.entries} entri) diurutkan sesuai abjad nama penulis.`);
    if (refs.numbersRemoved) notes.push(`${refs.numbersRemoved} nomor di depan entri daftar pustaka dihapus (gaya ${config.citationStyle} tidak bernomor).`);
    if (refs.capsFixed) notes.push(`${refs.capsFixed} nama penulis yang ditulis kapital semua di daftar pustaka dirapikan.`);
    if (refs.mismatched.length) {
        notes.push(`${refs.mismatched.length} entri daftar pustaka tidak sesuai gaya ${config.citationStyle} dan perlu disesuaikan manual: ${refs.mismatched.slice(0, 5).map((t) => `"${t}…"`).join("; ")}${refs.mismatched.length > 5 ? "; …" : ""}`);
    }


    /* -------- daftar isi/tabel/gambar/lampiran → field TOC -------- */
    realizeCaps(model);

    const toc = buildTocFields(model, nodes, outline, { upper: config.chapterUppercase, width: usableWidth });

    const demoted = demoteFromToc(nodes, outline, model.body);
    if (demoted.appendix) notes.push(`${demoted.appendix} judul lampiran hanya dicantumkan di Daftar Lampiran; Daftar Isi cukup memuat LAMPIRAN.`);
    if (demoted.shaped) notes.push(`${demoted.shaped} judul lampiran diseragamkan menjadi "LAMPIRAN n" dan judulnya di baris berikutnya, dengan penomoran berurutan.`);
    if (demoted.front) notes.push(`${demoted.front} judul bagian awal (persetujuan, pengesahan, pernyataan, motto, persembahan) tidak dicantumkan di Daftar Isi. Daftar Isi dimulai dari Kata Pengantar.`);

    /* -------- blok tanda tangan → tabel tanpa border -------- */

    const signatures = buildSignatureTables(model, nodes, outline, {
        width: usableWidth,
        font: config.fontFamily,
        halfPts: Math.round(config.fontSize * 2),
        tables: sigTables,
    });
    if (signatures.blocks) notes.push(`${signatures.blocks} blok tanda tangan disusun menjadi tabel tanpa garis dengan ruang tanda tangan yang seragam.`);
    if (signatures.cleaned) notes.push(`${signatures.cleaned} tabel tanda tangan dirapikan: garis tabel dibuang, isi rata tengah, dan blok tidak terpotong ke halaman lain.`);

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
        // Header bernomor halaman ("II – 1", "1") dikosongkan: nomor halaman sudah di footer.
        const pageOnly = headers.filter((h) => h.hasPageField);
        const pagedText = pageOnly.filter((h) => h.text);
        if (pagedText.length) notes.push(`${pagedText.length} header berisi nomor halaman (mis. "${pagedText[0].text.slice(0, 20)}") dikosongkan karena nomor halaman sudah diatur di footer.`);
        for (const h of [...pageOnly, ...duplicates]) headerRemap.set(h.rId, blankHeader);
        for (const h of duplicates) notes.push(`Header berisi "${h.text.slice(0, 60)}" dihapus karena menduplikasi judul halaman.`);

        const docXmlForRefs = serializeDocument(model);
        // Header berisi teks (judul berjalan, nama penulis, "SKRIPSI — BAB IV …") dikosongkan:
        // pedoman skripsi tidak memakai header teks.
        const running = headers.filter((h) => h.text && !headerRemap.has(h.rId) && docXmlForRefs.includes(`r:id="${h.rId}"`));
        for (const h of running) headerRemap.set(h.rId, blankHeader);
        if (running.length) {
            const sample = running.slice(0, 3).map((h) => `"${h.text.slice(0, 50)}"`).join(", ");
            notes.push(`${running.length} header berisi teks dikosongkan (${sample}${running.length > 3 ? ", …" : ""}). Pedoman skripsi tidak memakai header teks.`);
        }
    }

    /* -------- bingkai halaman & kolom: format skripsi satu kolom tanpa bingkai -------- */
    {
        const PGB = /<w:pgBorders\b[^>]*\/>|<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/g;
        const COLS = /<w:cols\b[^>]*\/>|<w:cols\b[\s\S]*?<\/w:cols>/g;
        const multiCol = (s: string) => /<w:cols\b[^>]*w:num="([2-9])"/.test(s);
        let borders = 0;
        let columns = 0;
        const fix = (s: string) => {
            let x = s;
            if (x.includes("<w:pgBorders")) { x = x.replace(PGB, ""); borders++; }
            if (multiCol(x)) { x = x.replace(COLS, `<w:cols w:space="720"/>`); columns++; }
            return x;
        };
        for (const n of flatten(model.body)) {
            if (n.kind === "paragraph" && n.hasSectPr) n.xml = n.xml.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, fix);
        }
        if (model.bodySectPr) model.bodySectPr = fix(model.bodySectPr);
        if (borders) notes.push(`Bingkai halaman (page border) di ${borders} section dihapus. Format skripsi tidak memakai bingkai halaman.`);
        if (columns) notes.push(`${columns} section bertata letak beberapa kolom dikembalikan menjadi satu kolom.`);
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
    const landscapeAfter = sections.map((s, i) => {
        const x = sectXml(s);
        return Boolean(x) && isLandscape(x!) && !toPortrait[i];
    });

    sections.forEach((s, i) => {
        const num = numbering[i];
        const stillLandscape = landscapeAfter[i];
        const opensPage = !s.firstRole || startsNewPage(s.firstRole);
        const prev = i > 0 ? numbering[i - 1] : null;
        // Section baru hanya membuka halaman bila memang perlu: judul pembuka halaman,
        // ganti orientasi, ganti format/awal nomor halaman, atau sampul.
        const continuous = Boolean(prev) && !opensPage && !num.cover && !prev!.cover &&
            landscapeAfter[i] === landscapeAfter[i - 1] && num.fmt === prev!.fmt && num.start === null;

        const plan: SectionPlan = {
            page: rules.page(s.firstZone ?? "main"),
            numbering: numbered && num.fmt ? { fmt: num.fmt, start: num.start } : null,
            footer: !numbered ? "keep" : num.cover || !footerRId ? { blank: blankFooter! } : { rId: footerRId },
            headerRemap: num.cover && blankHeader
                ? new Map([...headerRemap, ...[...(sectXml(s) ?? "").matchAll(/<w:headerReference\b[^>]*r:id="([^"]+)"/g)].map((m) => [m[1], blankHeader!] as [string, string])])
                : headerRemap,
            firstPageBlank: numbered && num.hideFirst && blankFooter && blankHeader
                ? { footer: blankFooter, header: blankHeader } : undefined,
            orientation: toPortrait[i] ? "portrait" : undefined,
            continuous: continuous || undefined,
        };
        if (stillLandscape) landscapeSections++;
        if (toPortrait[i]) portraitConverted++;

        // Section yang diawali judul pembuka halaman (bab, daftar, pustaka, lampiran) wajib
        // mulai di halaman baru, walau di dokumen asli bertipe "continuous".
        const mustOpen = i > 0 && Boolean(s.firstRole) && startsNewPage(s.firstRole!);
        const finish = (x: string) => (mustOpen ? x.replace(/<w:type\s+w:val="continuous"\s*\/>/, "") : x);
        if (s.ref.where === "para") {
            const p = s.ref.node;
            const start = p.xml.search(/<w:sectPr(?=[\s>/])/);
            if (start === -1) return;
            const end = elementEnd(p.xml, start, "w:sectPr");
            p.xml = p.xml.slice(0, start) + finish(patchSectPr(p.xml.slice(start, end), plan)) + p.xml.slice(end);
        } else if (model.bodySectPr) {
            model.bodySectPr = finish(patchSectPr(model.bodySectPr, plan));
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

    const coverPage = rules.page("cover");
    const coverHeight = coverPage.heightTwips - coverPage.margin.top - coverPage.margin.bottom;
    const tpl = await buildCoverFromTemplate(model, nodes, outline, {
        style: rules.cover(),
        font: config.fontFamily,
        width: usableWidth,
        height: coverHeight,
        docTypeFallback: config.docType,
        ai: opts.coverAI,
    });
    const cover = tpl.built
        ? { pages: 1, reordered: false, coverLines: 0 }
        : layoutCover(model, nodes, outline, { height: coverHeight, width: usableWidth, style: rules.cover() });
    const logos = tpl.built
        ? { centered: 0, regions: [] as VerifyRegion[] }
        : centerCoverImages(nodes, outline);

    if (tpl.built) {
        notes.push(`Halaman sampul disusun ulang dengan template General (data sampul dibaca ${tpl.source === "ai" ? "dengan bantuan AI dan dicocokkan dengan teks asli" : "dengan aturan"}): judul di atas; jenis dokumen, identitas, pembimbing, dan logo di tengah; identitas institusi di bawah.`);
        const dt = (tpl.fields.docType ?? "").toLowerCase();
        const chosen = String(config.docType);
        const coverIsProposal = /usulan|proposal/.test(dt);
        if (chosen === "skripsi" && coverIsProposal) {
            notes.push(`Sampul bertuliskan "${tpl.fields.docType}", sedangkan jenis dokumen yang dipilih Skripsi. Periksa kembali pilihan jenis dokumen.`);
        } else if (chosen === "proposal" && !coverIsProposal && /skripsi|tesis|thesis/.test(dt)) {
            notes.push(`Sampul bertuliskan "${tpl.fields.docType}", sedangkan jenis dokumen yang dipilih Proposal. Periksa kembali pilihan jenis dokumen.`);
        }
    } else {
        if (tpl.reason) notes.push(`Sampul tidak disusun ulang dari template: ${tpl.reason} Tata letak sampul asli dirapikan.`);
        if (cover.pages) notes.push(`${cover.pages} halaman sampul disusun ke urutan baku: judul di atas; jenis dokumen, penulis, dan logo di tengah; identitas institusi di margin bawah.`);
    }

    if (config.cleanupBlankPages) {
        const headingIds = new Set(nodes.filter((_, i) => isHeadingRole(outline.items[i]?.role ?? "body")).map((n) => n.id));
        collapseEmptyPages(model, report, headingIds);
    }

    realizeCaps(model);
    const listCache = refreshListCaches(model, { dropTc: new Set([...tableCaps.appendix, ...figureCaps.appendix]) });

    /* -------- tulis kembali paket -------- */

    stripHeadingDirectBold(model);
    const newXml = serializeDocument(model);
    zip.file("word/document.xml", newXml);

    await patchStyles(zip, config.fontFamily, Math.round(config.fontSize * 2), usableWidth);
    await patchWidowControl(zip);
    await ensureBoldHeadingStyles(zip);
    await patchListEntryStyle(zip);
    await enableUpdateFields(zip);
    await patchCompat(zip);

    const after = fingerprint(newXml);
    const issues = verify(before, after, {
        unorderedPrefix: cover.reordered ? cover.coverLines : 0,
        regions: [
            ...headings.regions,
            ...tableCaps.regions, ...figureCaps.regions,
            ...signatures.regions,
            ...listCache.regions,
            ...logos.regions,
            ...listRep.regions,
            ...flow.regions, ...keywords.regions, ...refs.regions,
            ...demoted.regions,
            ...chapters.regions,
            ...minorRegions,
            ...labels.regions,
            ...tpl.regions,
        ],
        added: {
            ...signatures.added,
            tables: signatures.added.tables - tpl.removed.tables,
            rows: signatures.added.rows - tpl.removed.rows,
            cells: signatures.added.cells - tpl.removed.cells,
        },
    });

    const preview = nodes.slice(0, 300).map((n, i) => ({
        role: outline.items[i]?.role ?? "body",
        zone: outline.items[i]?.zone ?? "main",
        text: n.kind === "table"
            ? `[Tabel ${n.rows.length} baris × ${n.rows[0]?.cells.length ?? 0} kolom]`
            : n.hasDrawing && !n.text.trim() ? "[Gambar]" : n.text.replace(/\t/g, " ").slice(0, 280),
    }));

    await patchHeadingStyles(zip);

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
        headings,
        coverRemoved: tpl.removed,
    };

}
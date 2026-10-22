/**
 * Penyuntingan bagian-bagian paket DOCX selain body: styles.xml, sectPr,
 * footer, content types, dan relationship.
 */

import JSZip from "jszip";
import { elementEnd, escapeXml, innerXml, mergeProps, openTag, splitElements } from "./xml";
import type { PageFormat } from "../rules/format-rules";

const W_NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/* ------------------------------------------------------------------ */
/* sectPr                                                              */
/* ------------------------------------------------------------------ */

const SECTPR_ORDER = [
    "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders",
    "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
    "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
];

export type SectionPlan = {
    page: PageFormat;
    /** `null` = jangan ubah penomoran section ini. */
    numbering: { fmt: string; start: number | null } | null;
    /**
     * Footer:
     * - `{ rId }`   : pakai footer nomor halaman kita (default + genap)
     * - `{ blank }` : pakai footer kosong (sampul tidak bernomor)
     * - `"keep"`    : biarkan footer asli
     */
    footer: { rId: string } | { blank: string } | "keep";
    /** Header yang hanya berisi nomor halaman → diarahkan ke header kosong. */
    headerRemap: Map<string, string>;
    /**
     * Halaman pertama section ini dihitung tetapi nomornya tidak dicetak
     * (halaman judul). Berisi rId footer & header kosong.
     */
    firstPageBlank?: { footer: string; header: string };

    /** Paksa portrait (tabel landscape yang ternyata muat). */
    orientation?: "portrait";
    /** Section dimulai menyambung, bukan halaman baru. */
    continuous?: boolean;
};

/** Halaman landscape bila diberi atribut orient atau lebarnya > tinggi. */
export function isLandscape(sectPr: string): boolean {
    const pgSz = /<w:pgSz\b[^>]*>/.exec(sectPr)?.[0] ?? "";
    if (/w:orient="landscape"/.test(pgSz)) return true;
    const w = Number(/w:w="(\d+)"/.exec(pgSz)?.[1] ?? 0);
    const h = Number(/w:h="(\d+)"/.exec(pgSz)?.[1] ?? 0);
    return w > 0 && h > 0 && w > h;
}

/**
 * Tulis ulang satu `w:sectPr`.
 *
 * Orientasi asli SELALU dipertahankan: halaman landscape (biasanya untuk
 * tabel lebar) tetap landscape dengan ukuran kertas yang diseragamkan. Margin
 * diputar 90° sehingga sisi jilid (margin kiri 4 cm pada halaman portrait)
 * berada di atas — begitulah halaman landscape dijilid di skripsi.
 */
export function patchSectPr(sectPr: string, plan: SectionPlan): string {
    const children = splitElements(innerXml(sectPr));
    const landscape = isLandscape(sectPr) && plan.orientation !== "portrait";

    const refs: string[] = [];
    const rest: { tag: string; xml: string }[] = [];
    for (const c of children) {
        if (c.tag === "w:headerReference") {
            // Header halaman pertama tidak dipakai lagi setelah titlePg dibuang.
            if (plan.footer !== "keep" && /w:type="first"/.test(c.xml)) continue;
            const rId = /r:id="([^"]+)"/.exec(c.xml)?.[1];
            const to = rId ? plan.headerRemap.get(rId) : undefined;
            refs.push(to ? c.xml.replace(`r:id="${rId}"`, `r:id="${to}"`) : c.xml);
        } else if (c.tag === "w:footerReference") {
            if (plan.footer === "keep") refs.push(c.xml);
        } else if (c.tag === "w:titlePg" && (plan.footer !== "keep" || plan.firstPageBlank)) {
            // Footer kita satu untuk semua halaman; titlePg akan membuat halaman
            // pertama tiap section (halaman pertama bab) kehilangan nomornya.
            continue;
        } else rest.push(c);
    }
    if (plan.footer !== "keep") {
        const id = "rId" in plan.footer ? plan.footer.rId : plan.footer.blank;
        refs.push(`<w:footerReference w:type="default" r:id="${id}"/>`);
        refs.push(`<w:footerReference w:type="even" r:id="${id}"/>`);
    }
    if (plan.firstPageBlank) {
        refs.push(`<w:headerReference w:type="first" r:id="${plan.firstPageBlank.header}"/>`);
        refs.push(`<w:footerReference w:type="first" r:id="${plan.firstPageBlank.footer}"/>`);
    }

    const existing = `<w:sectPr>${rest.map((c) => c.xml).join("")}</w:sectPr>`;

    const prevPgMar = rest.find((c) => c.tag === "w:pgMar")?.xml ?? "";
    const keep = (name: string, dflt: string) =>
        new RegExp(`w:${name}="(\\d+)"`).exec(prevPgMar)?.[1] ?? dflt;

    const { widthTwips: w, heightTwips: h, margin: m } = plan.page;
    const pgSz = landscape
        ? `<w:pgSz w:w="${h}" w:h="${w}" w:orient="landscape"/>`
        : `<w:pgSz w:w="${w}" w:h="${h}"/>`;
    const mar = landscape
        ? { top: m.left, right: m.top, bottom: m.right, left: m.bottom }
        : m;

    const patch: Record<string, string | null> = {
        "w:pgSz": pgSz,
        "w:pgMar":
            `<w:pgMar w:top="${mar.top}" w:right="${mar.right}" ` +
            `w:bottom="${mar.bottom}" w:left="${mar.left}" ` +
            `w:header="${keep("header", "708")}" w:footer="${keep("footer", "708")}" w:gutter="${keep("gutter", "0")}"/>`,
    };
    if (plan.continuous) patch["w:type"] = `<w:type w:val="continuous"/>`;
    if (plan.firstPageBlank) patch["w:titlePg"] = "<w:titlePg/>";
    if (plan.numbering) {
        const { fmt, start } = plan.numbering;
        patch["w:pgNumType"] = `<w:pgNumType w:fmt="${fmt}"${start !== null ? ` w:start="${start}"` : ""}/>`;
    }

    const merged = mergeProps(existing, patch, SECTPR_ORDER);
    const attrs = openTag(sectPr).replace(/^<w:sectPr/, "").replace(/\/?>$/, "");
    return `<w:sectPr${attrs}>${refs.join("")}${merged.join("")}</w:sectPr>`;
}

/* ------------------------------------------------------------------ */
/* Footer nomor halaman                                                 */
/* ------------------------------------------------------------------ */

function footerXml(align: "center" | "right", font: string, halfPoints: number): string {
    const rpr = `<w:rPr><w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}" w:cs="${escapeXml(font)}"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${W_NS}><w:p><w:pPr><w:jc w:val="${align}"/>${rpr}</w:pPr>` +
        `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r>${rpr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
        `<w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r>${rpr}<w:t>1</w:t></w:r>` +
        `<w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>` +
        `</w:p></w:ftr>`;
}

export async function findOrAddRelationship(zip: JSZip, target: string, type: string): Promise<string> {
    const path = "word/_rels/document.xml.rels";
    const xml = await zip.file(path)!.async("string");

    // Memakai ulang relasi yang sudah ada membuat proses ini idempoten: memformat
    // dokumen yang sama berulang kali tidak menumpuk bagian footer yatim.
    const existing = new RegExp(`<Relationship[^>]*Id="(rId\\d+)"[^>]*Target="${target}"[^>]*/>`).exec(xml)
        ?? new RegExp(`<Relationship[^>]*Target="${target}"[^>]*Id="(rId\\d+)"[^>]*/>`).exec(xml);
    if (existing) return existing[1];

    const used = new Set([...xml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1])));
    let n = 1;
    while (used.has(n)) n++;
    const rId = `rId${n}`;
    const rel = `<Relationship Id="${rId}" Type="${type}" Target="${target}"/>`;
    zip.file(path, xml.replace(/<\/Relationships>\s*$/, `${rel}</Relationships>`));
    return rId;
}

export async function addContentTypeOverride(zip: JSZip, partName: string, contentType: string) {
    const path = "[Content_Types].xml";
    const xml = await zip.file(path)!.async("string");
    if (xml.includes(`PartName="${partName}"`)) return;
    zip.file(path, xml.replace(/<\/Types>\s*$/, `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`));
}

/** Header/footer kosong — dipakai untuk menggantikan header nomor halaman lama dan sampul. */
export async function createBlankPart(zip: JSZip, kind: "header" | "footer"): Promise<string> {
    const tag = kind === "header" ? "w:hdr" : "w:ftr";
    const name = kind === "header" ? "headerFmtBlank.xml" : "footerFmtBlank.xml";
    zip.file(`word/${name}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${tag} ${W_NS}><w:p/></${tag}>`);
    await addContentTypeOverride(
        zip, `/word/${name}`,
        `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`,
    );
    return findOrAddRelationship(
        zip, name,
        `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`,
    );
}

export type HeaderInfo = { rId: string; target: string; text: string; hasPageField: boolean };

/** Isi setiap header yang dipakai dokumen, untuk memutuskan mana yang dibuang. */
export async function inspectHeaders(zip: JSZip): Promise<HeaderInfo[]> {
    const rels = await zip.file("word/_rels/document.xml.rels")?.async("string") ?? "";
    const out: HeaderInfo[] = [];
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const r = m[0];
        if (!/relationships\/header"/.test(r)) continue;
        const rId = /Id="([^"]+)"/.exec(r)?.[1];
        const target = /Target="([^"]+)"/.exec(r)?.[1];
        if (!rId || !target) continue;
        const xml = await zip.file(`word/${target.replace(/^\/?word\//, "")}`)?.async("string");
        if (!xml) continue;
        const hasPageField = /<w:instrText[^>]*>\s*PAGE\b|<w:fldSimple[^>]*w:instr="\s*PAGE\b/.test(xml);
        // Teks di luar hasil field PAGE (angka contoh "1") dihitung sebagai isi.
        const text = [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("")
            .replace(hasPageField ? /^\s*[\divxlcdmIVXLCDMA-Za-z]{1,6}\s*$/ : /$^/, "").trim();
        out.push({ rId, target, text, hasPageField });
    }
    return out;
}

/** Buat satu bagian footer berisi field PAGE dan daftarkan ke paket. */
export async function createPageNumberFooter(
    zip: JSZip, align: "center" | "right", font: string, halfPoints: number,
): Promise<string> {
    // Nama tetap, bukan berurut: bagian yang sama ditulis ulang setiap kali.
    const name = "footerFmt1.xml";
    zip.file(`word/${name}`, footerXml(align, font, halfPoints));
    await addContentTypeOverride(
        zip, `/word/${name}`,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    );
    return findOrAddRelationship(
        zip, name,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    );
}

/* ------------------------------------------------------------------ */
/* styles.xml                                                           */
/* ------------------------------------------------------------------ */

const STYLE_PPR_ORDER = ["w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:spacing", "w:ind", "w:jc", "w:outlineLvl"];
const STYLE_RPR_ORDER = ["w:rFonts", "w:b", "w:bCs", "w:i", "w:color", "w:sz", "w:szCs"];

function styleBlock(id: string, name: string, level: number, font: string, halfPoints: number): string {
    return (
        `<w:style w:type="paragraph" w:styleId="${id}">` +
        `<w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>` +
        `<w:uiPriority w:val="${9 + level}"/><w:qFormat/>` +
        `<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="${level}"/></w:pPr>` +
        `<w:rPr><w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}" w:cs="${escapeXml(font)}"/>` +
        `<w:b/><w:bCs/><w:color w:val="000000"/><w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr>` +
        `</w:style>`
    );
}

function replaceStyle(xml: string, styleId: string, build: (old: string) => string): string {
    const re = new RegExp(`<w:style\\b[^>]*w:styleId="${styleId}"[^>]*>`);
    const m = re.exec(xml);
    if (!m) return xml;
    const end = elementEnd(xml, m.index, "w:style");
    return xml.slice(0, m.index) + build(xml.slice(m.index, end)) + xml.slice(end);
}

/**
 * Samakan gaya bawaan dokumen dengan aturan format, dan pastikan Heading 1–3
 * punya `outlineLvl` agar daftar isi otomatis Word dapat membacanya.
 */
export async function patchStyles(zip: JSZip, font: string, halfPoints: number, usableWidth?: number) {
    const path = "word/styles.xml";
    const file = zip.file(path);
    if (!file) return;
    let xml = await file.async("string");

    const fontsXml = `<w:rFonts w:ascii="${escapeXml(font)}" w:hAnsi="${escapeXml(font)}" w:cs="${escapeXml(font)}" w:eastAsia="${escapeXml(font)}"/>`;
    const sizeXml = `<w:sz w:val="${halfPoints}"/>`;
    const sizeCsXml = `<w:szCs w:val="${halfPoints}"/>`;

    // docDefaults — dasar bagi seluruh dokumen.
    const ddRe = /<w:rPrDefault>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<\/w:rPrDefault>/;
    if (ddRe.test(xml)) {
        xml = xml.replace(ddRe, (_m, rpr) => {
            const merged = mergeProps(rpr ?? null, { "w:rFonts": fontsXml, "w:sz": sizeXml, "w:szCs": sizeCsXml }, STYLE_RPR_ORDER);
            return `<w:rPrDefault><w:rPr>${merged.join("")}</w:rPr></w:rPrDefault>`;
        });
    } else {
        xml = xml.replace(/<w:docDefaults>/, `<w:docDefaults><w:rPrDefault><w:rPr>${fontsXml}${sizeXml}${sizeCsXml}</w:rPr></w:rPrDefault>`);
    }

    for (const [id, name, level] of [["Heading1", "heading 1", 0], ["Heading2", "heading 2", 1], ["Heading3", "heading 3", 2]] as const) {
        if (new RegExp(`w:styleId="${id}"`).test(xml)) {
            xml = replaceStyle(xml, id, (old) => {
                const children = splitElements(innerXml(old));
                const keep = children.filter((c) => c.tag !== "w:pPr" && c.tag !== "w:rPr");
                const oldPpr = children.find((c) => c.tag === "w:pPr")?.xml ?? null;
                const oldRpr = children.find((c) => c.tag === "w:rPr")?.xml ?? null;
                const ppr = mergeProps(oldPpr, {
                    "w:keepNext": "<w:keepNext/>", "w:keepLines": "<w:keepLines/>",
                    "w:outlineLvl": `<w:outlineLvl w:val="${level}"/>`,
                }, STYLE_PPR_ORDER);
                const rpr = mergeProps(oldRpr, {
                    "w:rFonts": fontsXml, "w:b": "<w:b/>", "w:bCs": "<w:bCs/>",
                    "w:color": `<w:color w:val="000000"/>`, "w:sz": sizeXml, "w:szCs": sizeCsXml,
                }, STYLE_RPR_ORDER);
                return `${openTag(old)}${keep.map((c) => c.xml).join("")}<w:pPr>${ppr.join("")}</w:pPr><w:rPr>${rpr.join("")}</w:rPr></w:style>`;
            });
        } else {
            xml = xml.replace(/<\/w:styles>\s*$/, `${styleBlock(id, name, level, font, halfPoints)}</w:styles>`);
        }
    }

    if (usableWidth) {
        const upsert = (src: string, id: string, block: string) =>
            new RegExp(`w:styleId="${id}"`).test(src)
                ? replaceStyle(src, id, () => block)
                : src.replace(/<\/w:styles>\s*$/, `${block}</w:styles>`);
        const rpr = `<w:rPr>${fontsXml}${sizeXml}${sizeCsXml}</w:rPr>`;
        const toc = (lvl: number, left: number) =>
            `<w:style w:type="paragraph" w:styleId="TOC${lvl}"><w:name w:val="toc ${lvl}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>` +
            `<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${usableWidth}"/></w:tabs><w:spacing w:before="0" w:after="60" w:line="240" w:lineRule="auto"/><w:ind w:left="${left}" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>${rpr}</w:style>`;
        xml = upsert(xml, "TOC1", toc(1, 0));
        xml = upsert(xml, "TOC2", toc(2, 567));
        xml = upsert(xml, "TOC3", toc(3, 1134));
        // Judul bab dua baris ("BAB I" + "PENDAHULUAN"): bukan heading agar tidak
        // muncul dua kali di daftar isi; entrinya lewat penanda TC.
        xml = upsert(xml, "JudulBab",
            `<w:style w:type="paragraph" w:customStyle="1" w:styleId="JudulBab"><w:name w:val="Judul Bab"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
            `<w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="center"/></w:pPr><w:rPr>${fontsXml}<w:b/><w:bCs/>${sizeXml}${sizeCsXml}</w:rPr></w:style>`);

        xml = upsert(xml, "Caption",
            `<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>` +
            `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>` +
            `<w:rPr>${fontsXml}<w:color w:val="000000"/><w:sz w:val="${Math.max(20, halfPoints - 2)}"/><w:szCs w:val="${Math.max(20, halfPoints - 2)}"/></w:rPr></w:style>`);
    }

    zip.file(path, xml);
}

/** Elemen CT_Compat yang skemanya berada SEBELUM doNotExpandShiftReturn. */
const COMPAT_BEFORE = new Set([
    "w:useSingleBorderforContiguousCells", "w:wpJustification", "w:noTabHangInd", "w:noLeading",
    "w:spaceForUL", "w:noColumnBalance", "w:balanceSingleByteDoubleByteWidth", "w:noExtraLineSpacing",
    "w:doNotLeaveBackslashAlone", "w:ulTrailSpace",
]);

/** Elemen CT_Settings yang skemanya berada SETELAH w:compat. */
const SETTINGS_AFTER_COMPAT = [
    "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions",
    "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
];

/**
 * Cegah Word meregangkan baris yang diakhiri line break manual (Shift+Enter)
 * di paragraf rata kiri-kanan — sumber celah kata yang sangat lebar.
 */
export async function patchCompat(zip: JSZip) {
    const path = "word/settings.xml";
    const file = zip.file(path);
    if (!file) return;
    let xml = await file.async("string");
    if (/<w:doNotExpandShiftReturn\b/.test(xml)) return;
    const tag = "<w:doNotExpandShiftReturn/>";

    const compatStart = xml.search(/<w:compat>/);
    if (compatStart !== -1) {
        const end = elementEnd(xml, compatStart, "w:compat");
        const compat = xml.slice(compatStart, end);
        const kids = splitElements(innerXml(compat));
        const idx = kids.findIndex((k) => !COMPAT_BEFORE.has(k.tag));
        const merged = idx === -1
            ? [...kids.map((k) => k.xml), tag]
            : [...kids.slice(0, idx).map((k) => k.xml), tag, ...kids.slice(idx).map((k) => k.xml)];
        xml = xml.slice(0, compatStart) + `<w:compat>${merged.join("")}</w:compat>` + xml.slice(end);
    } else if (/<w:compat\s*\/>/.test(xml)) {
        xml = xml.replace(/<w:compat\s*\/>/, `<w:compat>${tag}</w:compat>`);
    } else {
        const block = `<w:compat>${tag}</w:compat>`;
        const pos = SETTINGS_AFTER_COMPAT
            .map((t) => xml.search(new RegExp(`<${t}[\\s>/]`)))
            .filter((p) => p !== -1)
            .sort((a, b) => a - b)[0];
        xml = pos !== undefined
            ? xml.slice(0, pos) + block + xml.slice(pos)
            : xml.replace(/<\/w:settings>\s*$/, `${block}</w:settings>`);
    }
    zip.file(path, xml);
}

/** Minta Word memperbarui field (daftar isi, nomor halaman) saat dibuka. */
export async function enableUpdateFields(zip: JSZip) {
    const path = "word/settings.xml";
    const file = zip.file(path);
    if (!file) return;
    const xml = await file.async("string");
    if (/<w:updateFields\b/.test(xml)) {
        zip.file(path, xml.replace(/<w:updateFields[^>]*\/>/, `<w:updateFields w:val="true"/>`));
        return;
    }
    // Skema CT_Settings menempatkan updateFields tepat sebelum w:compat.
    const tag = `<w:updateFields w:val="true"/>`;
    zip.file(path, /<w:compat[\s>/]/.test(xml)
        ? xml.replace(/<w:compat([\s>/])/, `${tag}<w:compat$1`)
        : xml.replace(/<\/w:settings>\s*$/, `${tag}</w:settings>`));
}
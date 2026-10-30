/**
 * Spasi otomatis ala HTML (w:beforeAutospacing / w:afterAutospacing) dari style seperti
 * "Normal (Web)" — terbawa saat teks di-copy dari browser/AI. Selama aktif di style, Word
 * MENGABAIKAN nilai spasi paragraf (before="0" pun tetap jadi ±14 pt), sehingga muncul celah
 * yang tidak bisa diatur formatter. Di semua paragraf (kecuali sampul), termasuk isi tabel,
 * spasi otomatis yang diwarisi dari style dimatikan secara eksplisit.
 */
import { withProps, PPR_ORDER } from "./xml";
import { cellParagraphs, flatten, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { paragraphsInComplexTables } from "./layout-groups";
import type { PageInfo } from "./page-openers";

type Auto = { before: boolean; after: boolean };
const on = (v: string | undefined) => v !== undefined && !/^(0|false|off)$/.test(v);

const ownPPr = (xml: string) => {
    const s = xml.indexOf("<w:pPr>");
    const e = s === -1 ? -1 : xml.indexOf("</w:pPr>", s);
    if (s === -1 || e === -1) return "";
    return xml.slice(s, e).replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "").replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, "");
};
const spacingTag = (pPr: string) => /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
const attr = (tag: string, name: string) => new RegExp(`\\bw:${name}="([^"]*)"`).exec(tag)?.[1];
function setAttr(tag: string, name: string, value: string): string {
    const re = new RegExp(`\\bw:${name}="[^"]*"`);
    return re.test(tag) ? tag.replace(re, `w:${name}="${value}"`) : tag.replace(/\s*\/>$/, ` w:${name}="${value}"/>`);
}

/** Spasi otomatis yang diwarisi paragraf dari style-nya (lewat basedOn, lalu docDefaults). */
export function autoSpacingResolver(stylesXml: string): (styleId: string | null) => Auto {
    const styles = new Map<string, { before?: boolean; after?: boolean; basedOn: string | null }>();
    let defaultId: string | null = null;
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        if (!/w:type="paragraph"/.test(m[1])) continue;
        const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
        if (!id) continue;
        if (/w:default="(?:1|true|on)"/.test(m[1])) defaultId = id;
        const tag = spacingTag(/<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(m[2])?.[1] ?? "");
        const b = attr(tag, "beforeAutospacing");
        const a = attr(tag, "afterAutospacing");
        styles.set(id, {
            before: b === undefined ? undefined : on(b),
            after: a === undefined ? undefined : on(a),
            basedOn: /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1] ?? null,
        });
    }
    const dd = spacingTag(/<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/.exec(stylesXml)?.[0] ?? "");
    const ddAuto: Auto = { before: on(attr(dd, "beforeAutospacing")), after: on(attr(dd, "afterAutospacing")) };

    return (styleId) => {
        let before: boolean | undefined;
        let after: boolean | undefined;
        let cur: string | null = styleId ?? defaultId;
        for (let d = 0; cur && d < 10; d++) {
            const s = styles.get(cur);
            if (!s) break;
            if (before === undefined && s.before !== undefined) before = s.before;
            if (after === undefined && s.after !== undefined) after = s.after;
            cur = s.basedOn;
        }
        return { before: before ?? ddAuto.before, after: after ?? ddAuto.after };
    };
}

/** Spasi otomatis yang masih berlaku pada paragraf (milik sendiri > dari style). */
function effective(xml: string, resolve: (id: string | null) => Auto): Auto & { tag: string; styleId: string | null } {
    const pPr = ownPPr(xml);
    const styleId = /<w:pStyle\s+w:val="([^"]+)"/.exec(pPr)?.[1] ?? null;
    const tag = spacingTag(pPr);
    const inherited = resolve(styleId);
    const b = attr(tag, "beforeAutospacing");
    const a = attr(tag, "afterAutospacing");
    return {
        before: b !== undefined ? on(b) : inherited.before,
        after: a !== undefined ? on(a) : inherited.after,
        tag,
        styleId,
    };
}

function neutralize(p: ParaNode, resolve: (id: string | null) => Auto): boolean {
    const st = effective(p.xml, resolve);
    if (!st.before && !st.after) return false;
    let tag = st.tag || "<w:spacing/>";
    if (st.before) {
        tag = setAttr(tag, "beforeAutospacing", "0");
        if (attr(tag, "before") === undefined) tag = setAttr(tag, "before", "0");
    }
    if (st.after) {
        tag = setAttr(tag, "afterAutospacing", "0");
        if (attr(tag, "after") === undefined) tag = setAttr(tag, "after", "0");
    }
    p.xml = withProps(p.xml, "w:pPr", { "w:spacing": tag }, PPR_ORDER);
    return true;
}

/** Semua paragraf: badan dokumen + isi sel tabel (rekursif), beserta node pemilik zonanya. */
function everyParagraph(model: DocModel): { p: ParaNode; owner: ParaNode | TableNode }[] {
    const out: { p: ParaNode; owner: ParaNode | TableNode }[] = [];
    const visitTable = (t: TableNode, owner: ParaNode | TableNode) => {
        for (const r of t.rows) for (const c of r.cells) {
            for (const p of cellParagraphs(c)) out.push({ p, owner });
            for (const x of flatten(c.children)) if (x.kind === "table") visitTable(x, owner);
        }
    };
    for (const n of flatten(model.body)) {
        if (n.kind === "paragraph") out.push({ p: n, owner: n });
        else visitTable(n, n);
    }
    return out;
}

export function neutralizeAutoSpacing(model: DocModel, info: PageInfo, stylesXml: string): number {
    const resolve = autoSpacingResolver(stylesXml);
    const frozen = paragraphsInComplexTables(model);
    let count = 0;
    for (const { p, owner } of everyParagraph(model)) {
        if (info.get(owner)?.zone === "cover") continue;
        if (frozen.has(p)) continue;                     // tabel tata letak berisi objek melayang: utuh
        if (neutralize(p, resolve)) count++;
    }
    return count;
}

/** Untuk uji regresi: paragraf yang masih kena spasi otomatis dari style. */
export function countAutoSpacing(model: DocModel, stylesXml: string): { count: number; styles: string[] } {
    const resolve = autoSpacingResolver(stylesXml);
    const styles = new Set<string>();
    let count = 0;
    for (const { p } of everyParagraph(model)) {
        const st = effective(p.xml, resolve);
        if (st.before || st.after) { count++; styles.add(st.styleId ?? "(Normal)"); }
    }
    return { count, styles: [...styles] };
}
/**
 * Warna teks untuk naskah cetak:
 *  - Tautan (DOI/URL/email) di seluruh naskah kecuali sampul: hitam, tanpa garis bawah, tetap
 *    bisa diklik. Tautan dikenali dari SEMUA cara Word menyimpannya, bukan dari warnanya:
 *      · elemen <w:hyperlink> (hasil Ctrl+K, Mendeley/Zotero, copy-paste dari web/jurnal)
 *      · field HYPERLINK (<w:fldSimple> maupun fldChar begin/separate/end)
 *      · run ber-style karakter tautan ("Hyperlink", "FollowedHyperlink", "Internet Link"
 *        dari LibreOffice, serta style turunannya)
 *  - Daftar Pustaka: semua teks hitam (warna tema/langsung dari copy-paste dibuang), stabilo
 *    dan latar warna di belakang teks dihapus.
 * Warna dipasang langsung pada run sehingga menimpa warna dari style apa pun.
 */
import { withProps, RPR_ORDER } from "./xml";
import { cellParagraphs, flatten, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { PageInfo } from "./page-openers";

const BLACK = `<w:color w:val="000000"/>`;
const NO_UNDERLINE = `<w:u w:val="none"/>`;
const LINK_NAME = /^(?:hyperlink|followed\s*hyperlink|internet\s*link|visited\s*internet\s*link|tautan|hyperlink\s*\d*)$/i;
const LINK_ID = /hyperlink|internetlink/i;
const END_REFS = new Set(["chapter", "chapter_label", "appendix_heading", "front_heading", "list_heading"]);

/** Style karakter tautan (termasuk yang basedOn style tautan). */
function linkStyles(stylesXml: string): Set<string> {
    const base = new Map<string, string>();
    const link = new Set<string>();
    for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        if (!/w:type="character"/.test(m[1])) continue;
        const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
        if (!id) continue;
        const name = /<w:name\s+w:val="([^"]+)"/.exec(m[2])?.[1] ?? "";
        if (LINK_NAME.test(name.trim()) || LINK_ID.test(id)) link.add(id);
        const b = /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1];
        if (b) base.set(id, b);
    }
    for (const id of base.keys()) {
        for (let cur = base.get(id), n = 0; cur && n < 20; cur = base.get(cur), n++) {
            if (link.has(cur)) { link.add(id); break; }
        }
    }
    return link;
}

type FieldState = { stack: { link: boolean; result: boolean }[] };

const TOKEN_RE = /<w:hyperlink\b[^>]*\/>|<w:hyperlink\b[^>]*>|<\/w:hyperlink>|<w:fldSimple\b[^>]*\/>|<w:fldSimple\b[^>]*>|<\/w:fldSimple>|<w:r\b(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;

function patchParagraph(p: ParaNode, isRef: boolean, linkStyle: Set<string>, fields: FieldState): boolean {
    const end = p.xml.indexOf("</w:pPr>");
    const head = end === -1 ? "" : p.xml.slice(0, end);
    const body = end === -1 ? p.xml : p.xml.slice(end);
    let hyper = 0;
    const simple: boolean[] = [];
    let changed = false;
    const out = body.replace(TOKEN_RE, (tok) => {
        if (tok.startsWith("<w:hyperlink")) { if (!tok.endsWith("/>")) hyper++; return tok; }
        if (tok === "</w:hyperlink>") { hyper = Math.max(0, hyper - 1); return tok; }
        if (tok.startsWith("<w:fldSimple")) {
            if (!tok.endsWith("/>")) simple.push(/w:instr="\s*HYPERLINK\b/i.test(tok));
            return tok;
        }
        if (tok === "</w:fldSimple>") { simple.pop(); return tok; }

        // Run: perbarui status field kompleks (bisa melintasi beberapa run/paragraf).
        const instr = [...tok.matchAll(/<w:instrText\b[^>]*>([^<]*)<\/w:instrText>/g)].map((m) => m[1]).join("");
        const top = fields.stack[fields.stack.length - 1];
        if (instr && top && !top.result && /^\s*HYPERLINK\b/i.test(instr)) top.link = true;
        const fld = /<w:fldChar\b[^>]*w:fldCharType="(\w+)"/.exec(tok)?.[1];
        if (fld === "begin") fields.stack.push({ link: false, result: false });
        else if (fld === "separate" && top) top.result = true;
        else if (fld === "end") fields.stack.pop();
        if (fld || instr || !/<w:t[\s>]|<w:tab\s*\/>/.test(tok)) return tok;

        const rStyle = /<w:rStyle\s+w:val="([^"]+)"/.exec(/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(tok)?.[1] ?? "")?.[1];
        const isLink = hyper > 0 || simple.some(Boolean) ||
            fields.stack.some((f) => f.link && f.result) || (rStyle !== undefined && linkStyle.has(rStyle));
        if (!isLink && !isRef) return tok;
        const patch: Record<string, string | null> = { "w:color": BLACK };
        if (isLink) patch["w:u"] = NO_UNDERLINE;
        if (isRef) { patch["w:highlight"] = null; patch["w:shd"] = null; }
        const next = withProps(tok, "w:rPr", patch, RPR_ORDER);
        if (next !== tok) changed = true;
        return next;
    });
    if (changed) p.xml = head + out;
    return changed;
}

function paragraphsOf(n: ParaNode | TableNode): ParaNode[] {
    if (n.kind === "paragraph") return [n];
    return n.rows.flatMap((row) => row.cells.flatMap((cell) => [
        ...cellParagraphs(cell),
        ...cell.children.filter((c): c is TableNode => c.kind === "table").flatMap(paragraphsOf),
    ]));
}

/** Mengembalikan jumlah paragraf { tautan, pustaka } yang diubah. */
export function blackenLinks(model: DocModel, info: PageInfo, stylesXml: string): { links: number; refs: number } {
    const linkStyle = linkStyles(stylesXml);
    const fields: FieldState = { stack: [] };
    let inRefs = false;
    let links = 0;
    let refs = 0;
    for (const node of flatten(model.body)) {
        const meta = info.get(node);
        if (meta?.role === "reference_heading") inRefs = true;
        else if (meta && END_REFS.has(meta.role)) inRefs = false;
        if (meta?.zone === "cover") continue;
        for (const p of paragraphsOf(node)) {
            if (patchParagraph(p, inRefs, linkStyle, fields)) {
                if (inRefs) refs++;
                else links++;
            }
        }
    }
    return { links, refs };
}
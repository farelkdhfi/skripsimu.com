/**
 * Gambar isi yang "melayang" (wrap Square/Tight/Top and Bottom/In Front/Behind) dijadikan
 * "In Line with Text" di paragrafnya sendiri. Gambar melayang yang nyangkut di paragraf teks
 * mengacak susunan paragraf; sebagai inline, gambar ikut alur teks dan dikenali sebagai gambar
 * (keterangan, Daftar Gambar).
 *
 * Yang SENGAJA tetap melayang (ditentukan dari ciri gambarnya, bukan dari dokumen tertentu):
 *  - logo/gambar di sampul (diatur modul sampul)
 *  - tanda tangan & materai: gambar kecil di paragraf kosong/pendek yang berdekatan dengan
 *    penanda tanda tangan ("Penulis", "Yang menyatakan", tanggal, nama dalam kurung, NIP/NPM, …)
 *  - bentuk (kotak, garis, panah, text box) dan gambar yang ditumpuki bentuk seperti itu
 *  - gambar latar di belakang teks yang berukuran besar (scan lembar penuh, watermark)
 *  - gambar di dalam mc:AlternateContent (bentuk Word 2010+ beserta cadangannya)
 *
 * Teks tidak berubah: gambar hanya dipindah keluar dari paragraf teks ke paragrafnya sendiri,
 * tepat di atas keterangannya (bila ada) atau di dekat posisi tampilnya.
 */
import { withProps, PPR_ORDER } from "./xml";
import { flatten, nextVisible, parseBlocks, prevVisible, type BlockNode, type DocModel, type ParaNode } from "./document-model";
import { captionKind, scanStructure, type Zone } from "../structure/roles";
import { inlineOf, maxId } from "./figure-split";
import { nearFloatingShape, tableClass } from "./layout-groups";

const CUE_RE = /^(mengetahui|menyetujui|mengesahkan|disetujui|disahkan|pembimbing|penguji|ketua|dekan|wakil|rektor|direktur|kepala|sekretaris|koordinator|kaprodi|yang\s+menyatakan|yang\s+membuat|penulis|peneliti|mahasiswa|materai|meterai|hormat\s+(saya|kami))\b/i;
const DATE_RE = /^[A-Z][A-Za-z.'\- ]{1,40},\s*(\d{1,2}\s+)?[A-Za-z]{3,12}\s+\d{4}\.?$/;
const ID_RE = /^(NIP|NIDN|NIDK|NIK|NIM|NPM|NRP|NUPTK)\b/i;
const PAREN_NAME_RE = /^\(\s*[A-Z][^()]{2,70}\)$/;
const isCue = (t: string) => t.length <= 80 && (CUE_RE.test(t) || DATE_RE.test(t) || ID_RE.test(t) || PAREN_NAME_RE.test(t));

const EMU_CM = 360000;
const SMALL_EMU = 7 * EMU_CM;
const BACKGROUND_EMU = 10 * EMU_CM;
const DRAWING_RE = /<w:drawing>[\s\S]*?<\/w:drawing>/g;
const RUN_RE = /<w:r\b(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const CONTENT_RE = /<w:t[\s>]|<w:tab\s*\/>|<w:br\b|<w:cr\b|<w:drawing\b|<w:pict\b|<w:object\b|<w:fldChar\b|<w:instrText\b|<w:sym\b|<w:footnoteReference\b|<w:endnoteReference\b|<w:commentReference\b|<w:ptab\b/;

type Info = {
    xml: string;
    anchor: boolean;
    picture: boolean;
    shapeOnly: boolean;
    behind: boolean;
    cx: number;
    cy: number;
    wrap: string | null;
    vRel: string | null;
    vOffset: number | null;
    hAlign: string | null;
    hOffset: number | null;
};

function infoOf(d: string): Info {
    const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(d);
    const head = /<wp:anchor\b[^>]*>/.exec(d)?.[0] ?? "";
    const pv = /<wp:positionV\b[^>]*relativeFrom="(\w+)"[^>]*>([\s\S]*?)<\/wp:positionV>/.exec(d);
    const ph = /<wp:positionH\b[^>]*>([\s\S]*?)<\/wp:positionH>/.exec(d);
    const picture = /<pic:pic\b|<c:chart\b|<dgm:relIds\b|<wpc:wpc\b/.test(d);
    return {
        xml: d,
        anchor: d.includes("<wp:anchor"),
        picture,
        shapeOnly: !picture && /<wps:wsp\b|<wpg:wgp\b/.test(d),
        behind: /\bbehindDoc="(?:1|true)"/.test(head),
        cx: Number(ext?.[1] ?? 0),
        cy: Number(ext?.[2] ?? 0),
        wrap: /<wp:wrap(\w+)/.exec(d)?.[1] ?? null,
        vRel: pv?.[1] ?? null,
        vOffset: pv ? (/<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(pv[2]) ? Number(/<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(pv[2])![1]) : null) : null,
        hAlign: ph ? /<wp:align>(\w+)<\/wp:align>/.exec(ph[1])?.[1] ?? null : null,
        hOffset: ph ? (/<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(ph[1]) ? Number(/<wp:posOffset>(-?\d+)<\/wp:posOffset>/.exec(ph[1])![1]) : null) : null,
    };
}

/** Gambar (bukan bentuk) di luar mc:AlternateContent. */
function drawingsOf(p: ParaNode): Info[] {
    const out: Info[] = [];
    for (const run of p.xml.matchAll(RUN_RE)) {
        if (run[0].includes("<mc:AlternateContent")) continue;
        for (const d of run[0].matchAll(DRAWING_RE)) out.push(infoOf(d[0]));
    }
    return out;
}

const hasShapes = (p: ParaNode) =>
    /<wps:wsp\b|<wpg:wgp\b|<v:shape\b|<v:rect\b|<v:line\b|<v:oval\b|<v:group\b|<w:txbxContent\b/.test(p.xml);

type Slot = { p: ParaNode; list: BlockNode[]; zone: Zone };

export function inlineFloatingImages(model: DocModel): { inlined: number; moved: number } {
    const out = { inlined: 0, moved: 0 };
    const top = flatten(model.body);
    const { zones } = scanStructure(top.map((n) => n.kind === "paragraph"
        ? { text: n.text, hasDrawing: n.hasDrawing, isTable: false, pageEnd: n.hasSectPr || n.hasPageBreak, pageStart: n.hadPageBreakBefore }
        : { text: "", hasDrawing: false, isTable: true }));
    const zoneOf = new Map(top.map((n, i) => [n, zones[i]] as const));

    // Semua paragraf (termasuk isi sel tabel) berurutan, dengan daftar induk & zonanya.
    const slots: Slot[] = [];
    const walk = (list: BlockNode[], inherited: Zone | null) => {
        for (const n of list) {
            if (n.kind === "container") { walk(n.children, inherited); continue; }
            const z = inherited ?? (n.kind === "paragraph" || n.kind === "table" ? zoneOf.get(n) ?? "main" : "main");
            if (n.kind === "paragraph") slots.push({ p: n, list, zone: z });
            // Tabel tata letak berisi objek melayang: satu susunan, gambarnya tidak dipindah.
            else if (n.kind === "table" && tableClass(n) !== "complex") for (const r of n.rows) for (const c of r.cells) walk(c.children, z);
        }
    };
    walk(model.body, null);
    const indexOfSlot = new Map(slots.map((s, i) => [s.p, i] as const));

    /** Penanda tanda tangan di paragraf ini atau di ≤ 2 baris berteks di atas/bawahnya. */
    const nearCue = (p: ParaNode) => {
        if (isCue(p.text.trim())) return true;
        const si = indexOfSlot.get(p)!;
        for (const step of [-1, 1]) {
            let seen = 0;
            for (let k = si + step; k >= 0 && k < slots.length && seen < 2; k += step) {
                const t = slots[k].p.text.trim();
                if (!t) continue;
                seen++;
                if (isCue(t)) return true;
            }
        }
        return false;
    };

    const sect = model.bodySectPr ?? "";
    const pgW = Number(/<w:pgSz\b[^>]*w:w="(\d+)"/.exec(sect)?.[1] ?? 11906);
    const marL = Number(/<w:pgMar\b[^>]*w:left="(\d+)"/.exec(sect)?.[1] ?? 1701);
    const marR = Number(/<w:pgMar\b[^>]*w:right="(\d+)"/.exec(sect)?.[1] ?? 1701);
    const textW = (pgW - marL - marR) * 635;

    const alignOf = (d: Info, zone: Zone): "left" | "center" | "right" => {
        if (zone === "main") return "center";
        if (d.hAlign) return /left|inside/.test(d.hAlign) ? "left" : /right|outside/.test(d.hAlign) ? "right" : "center";
        if (d.hOffset !== null) {
            const mid = d.hOffset + d.cx / 2;
            return mid < textW * 0.4 ? "left" : mid > textW * 0.6 ? "right" : "center";
        }
        return "center";
    };

    let nextId = maxId(model.body) + 1;

    for (const slot of slots) {
        const { p, list, zone } = slot;
        if (zone === "cover" || !p.hasDrawing) continue;
        const ds = drawingsOf(p).filter((d) => d.anchor);
        if (!ds.length) continue;

        // Gambar yang ditumpuki bentuk (panah, label) di paragraf ini/sebelah: satu susunan.
        const li = list.indexOf(p);
        const neighbors = [list[prevVisible(list, li)], list[nextVisible(list, li)]].filter((n): n is ParaNode => n?.kind === "paragraph");
        // Kotak teks/shape melayang dalam ≤ 3 paragraf (mis. keterangan "Ilustrasi 1" berupa kotak
        // teks yang posisinya relatif ke paragraf pengganjal di bawah foto) = satu susunan.
        const composite = hasShapes(p) || neighbors.some(hasShapes) || nearFloatingShape(list, li, 3);
        const shortPara = p.text.trim().length <= 80;
        const cue = nearCue(p);

        const convert = ds.filter((d) => {
            if (!d.picture || d.shapeOnly || composite) return false;
            if (d.behind && (d.cx >= BACKGROUND_EMU || d.cy >= BACKGROUND_EMU)) return false;
            const small = Math.max(d.cx, d.cy) <= SMALL_EMU;
            if (small && shortPara && cue) return false;
            return true;
        });
        if (!convert.length) continue;

        const align = alignOf(convert[0], zone);
        const keep = new Set(convert.map((d) => d.xml));

        if (!p.text.trim()) {
            // Paragraf khusus gambar: diubah di tempat.
            let xml = p.xml.replace(DRAWING_RE, (d) => (keep.has(d) ? inlineOf(d) : d));
            xml = withProps(xml, "w:pPr", {
                "w:jc": `<w:jc w:val="${align}"/>`,
                "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
                "w:framePr": null,
            }, PPR_ORDER);
            Object.assign(p, { ...(parseBlocks(xml, 0)[0] as ParaNode), id: p.id });
            out.inlined += convert.length;
            continue;
        }

        // Gambar di paragraf teks: dikeluarkan ke paragraf sendiri.
        const runs: string[] = [];
        let xml = p.xml.replace(RUN_RE, (run) => {
            if (run.includes("<mc:AlternateContent")) return run;
            const rPr = /^<w:r\b[^>]*>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(run)?.[1] ?? "";
            const rest = run.replace(DRAWING_RE, (d) => {
                if (!keep.has(d)) return d;
                runs.push(`<w:r>${rPr}${inlineOf(d)}</w:r>`);
                return "";
            });
            return CONTENT_RE.test(rest.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "")) ? rest : "";
        });
        if (!runs.length) continue;

        let after = true;
        const nextText = list.slice(li + 1).find((n) => n.kind !== "paragraph" || n.text.trim() || n.hasDrawing);
        const captionNext = nextText?.kind === "paragraph" && captionKind(nextText.text.trim()) === "figure_caption";
        const d0 = convert[0];
        if (!captionNext) {
            if (d0.wrap && /^(Square|Tight|Through)$/.test(d0.wrap)) after = false;        // foto di samping teks → di atasnya
            else if (d0.vRel && /paragraph|line/.test(d0.vRel) && (d0.vOffset ?? 0) < 0) after = false;
        }
        const figXml = `<w:p><w:pPr>${captionNext ? "<w:keepNext/>" : ""}<w:spacing w:before="0" w:after="0"/>` +
            `<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="${align}"/></w:pPr>${runs.join("")}</w:p>`;
        const fig = parseBlocks(figXml, nextId++)[0] as ParaNode;
        Object.assign(p, { ...(parseBlocks(xml, 0)[0] as ParaNode), id: p.id });
        list.splice(after ? li + 1 : li, 0, fig);
        out.inlined += convert.length;
        out.moved++;
    }
    return out;
}
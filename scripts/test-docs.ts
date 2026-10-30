/**
 * Uji regresi formatter skripsi.
 *
 *   npm run test:docs                       → format semua .docx di test-docs/, periksa, bandingkan snapshot
 *   npm run test:docs -- --update           → setujui hasil sekarang sebagai snapshot baru
 *   npm run test:docs -- riki               → hanya file yang namanya memuat "riki"
 *   npm run test:docs -- dokumen --trace    → tunjukkan di tahap mana teks mulai berubah
 *
 * Tanpa AI, jadi hasilnya deterministik. Hasil format disimpan di test-docs/__out__/.
 */
import { useNodeIO } from "../src/features/thesis-formatter/front/node";
import { readFile, writeFile, mkdir, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { transformDocx } from "../src/features/thesis-formatter/docx/transform";
import { configForPreset } from "../src/features/thesis-formatter/schemas/format-config";
import { flatten, parseDocument, type ParaNode, type TableNode } from "../src/features/thesis-formatter/docx/document-model";
import { fingerprint } from "../src/features/thesis-formatter/docx/verify";
import { countAutoSpacing } from "../src/features/thesis-formatter/docx/autospacing";
useNodeIO();

const ROOT = path.resolve("test-docs");
const OUT = path.join(ROOT, "__out__");
const SNAP = path.join(ROOT, "__snapshots__");
const args = process.argv.slice(2);
const UPDATE = args.includes("--update");
const TRACE = args.includes("--trace");
const FILTER = args.filter((a) => !a.startsWith("--")).map((a) => a.toLowerCase());

/**
 * `frontAI`: berkas JSON (di test-docs/) berisi jawaban AI bagian depan yang diputar ulang — untuk
 * menguji jalur AI tanpa memanggil AI (mis. jawaban AI sungguhan yang disalin dari log, atau
 * jawaban yang sengaja salah).
 */
type Case = { preset?: string; config?: Record<string, unknown>; frontAI?: string };

/* ------------------------------------------------------------------ */
/* Pembacaan XML ringkas                                                */
/* ------------------------------------------------------------------ */

const decode = (s: string) => s
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function textOf(xml: string): string {
    let out = "";
    for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br(?![^>]*w:type="page")[^>]*\/>/g)) {
        out += m[1] !== undefined ? decode(m[1]) : m[0].startsWith("<w:tab") ? "→" : "↵";
    }
    return out.replace(/\s+/g, " ").trim();
}

/** pPr milik paragraf sendiri, tanpa rPr penanda paragraf, sectPr, dan riwayat revisi. */
function ownPPr(xml: string): string {
    const s = xml.indexOf("<w:pPr>");
    const e = s === -1 ? -1 : xml.indexOf("</w:pPr>", s);
    if (s === -1 || e === -1) return "";
    return xml.slice(s, e)
        .replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "")
        .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, "")
        .replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, "");
}

const attr = (tag: string, name: string) => new RegExp(`w:${name}="([^"]*)"`).exec(tag)?.[1];

function describePara(p: ParaNode): string {
    const pPr = ownPPr(p.xml);
    const el = (t: string) => new RegExp(`<w:${t}\\b[^>]*\\/?>`).exec(pPr)?.[0] ?? "";
    const style = attr(el("pStyle"), "val") ?? "-";
    const jc = attr(el("jc"), "val") ?? "-";
    const ind = el("ind");
    const indS = ind
        ? `${attr(ind, "left") ?? attr(ind, "start") ?? 0}/${attr(ind, "hanging") ? "h" + attr(ind, "hanging") : "f" + (attr(ind, "firstLine") ?? 0)}`
        : "-";
    const sp = el("spacing");
    const spS = sp
        ? `${attr(sp, "before") ?? "."}/${attr(sp, "after") ?? "."}/${attr(sp, "line") ?? "."}${(attr(sp, "lineRule") ?? "")[0] ?? ""}`
        : "-";
    const tabs = [...(/<w:tabs>([\s\S]*?)<\/w:tabs>/.exec(pPr)?.[1] ?? "").matchAll(/w:pos="(-?\d+)"/g)].map((m) => m[1]).join(",");
    const flags = [
        /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPr) ? "PBB" : "",
        /<w:br\b[^>]*w:type="page"/.test(p.xml) ? "PB" : "",
        /<w:sectPr\b/.test(p.xml) ? "SECT" : "",
        /<w:numPr>[\s\S]*?<w:numId\s+w:val="(?!0")/.test(pPr) ? "NUM" : "",
        /<wp:inline\b/.test(p.xml) ? "IMG" : "",
        /<wp:anchor\b/.test(p.xml) ? "FLOAT" : "",
        /<w:keepNext(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPr) ? "KN" : "",
        /<w:instrText[^>]*>\s*SEQ\b/.test(p.xml) ? "SEQ" : "",
        /<w:instrText[^>]*>\s*(TOC|PAGEREF)\b/.test(p.xml) ? "TOCF" : "",
    ].filter(Boolean).join(" ");
    const sizes = [...new Set([...p.xml.slice(p.xml.indexOf("</w:pPr>") + 1).matchAll(/<w:sz\s+w:val="(\d+)"/g)].map((m) => m[1]))].join(",");
    return `P ${style} | ${jc} | ind ${indS} | sp ${spS}${tabs ? ` | tab ${tabs}` : ""}${sizes ? ` | sz ${sizes}` : ""}${flags ? ` | ${flags}` : ""} | ${textOf(p.xml).slice(0, 70)}`;
}

function describeTable(t: TableNode): string {
    const cols = Math.max(0, ...t.rows.map((r) => r.cells.length));
    const tb = /<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/.exec(t.props ?? "")?.[1] ?? "";
    const lines = [...tb.matchAll(/<w:(top|bottom|left|right|insideH|insideV)\b[^>]*w:val="([^"]+)"/g)]
        .filter((m) => m[2] !== "nil" && m[2] !== "none").map((m) => m[1][0] + (m[1].startsWith("inside") ? m[1].slice(-1) : ""));
    const text = t.rows.slice(0, 2).map((r) => r.cells.map((c) => flatten(c.children)
        .map((x) => (x.kind === "paragraph" ? textOf(x.xml) : "")).join(" ")).join("|")).join(" / ");
    return `T ${t.rows.length}x${cols} | border ${lines.join("") || "none"} | ${text.slice(0, 60)}`;
}

function describeSections(xml: string): string[] {
    return [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m, i) => {
        const s = m[0];
        const pg = /<w:pgNumType\b[^>]*>/.exec(s)?.[0] ?? "";
        const sz = /<w:pgSz\b[^>]*>/.exec(s)?.[0] ?? "";
        const mar = /<w:pgMar\b[^>]*>/.exec(s)?.[0] ?? "";
        return `S${i} type=${attr(/<w:type\b[^>]*>/.exec(s)?.[0] ?? "", "val") ?? "nextPage"} num=${attr(pg, "fmt") ?? "-"}/${attr(pg, "start") ?? "-"}` +
            ` ${attr(sz, "orient") ?? "portrait"} mar=${attr(mar, "top")},${attr(mar, "bottom")},${attr(mar, "left")},${attr(mar, "right")}` +
            ` footer=${(s.match(/<w:footerReference\b/g) ?? []).length}${/<w:titlePg\b/.test(s) ? " titlePg" : ""}`;
    });
}

/* ------------------------------------------------------------------ */
/* Pemeriksaan wajib                                                    */
/* ------------------------------------------------------------------ */

function xmlProblem(xml: string): string | null {
    const stack: { tag: string; ns: Set<string> }[] = [];
    const re = /<!\[CDATA\[[\s\S]*?\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[^\s=>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
    // Prefix namespace harus dideklarasikan di elemen itu sendiri atau leluhurnya (Word menolak file bila tidak).
    const declared = (attrs: string) => new Set([...attrs.matchAll(/\sxmlns:([\w.-]+)\s*=/g)].map((x) => x[1]));
    const inScope = (p: string, own: Set<string>) => p === "xml" || own.has(p) || stack.some((s) => s.ns.has(p));
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
        if (!m[2]) continue;
        if (!m[1]) {
            const own = declared(m[3] ?? "");
            const prefixes = [m[2], ...[...(m[3] ?? "").matchAll(/\s([\w.-]+):[\w.-]+\s*=/g)].map((x) => `${x[1]}:`)]
                .map((n) => (n.includes(":") ? n.split(":")[0] : "")).filter((p) => p && p !== "xmlns");
            const bad = prefixes.find((p) => !inScope(p, own));
            if (bad) return `prefix namespace "${bad}:" tidak dideklarasikan pada <${m[2]}> (karakter ${m.index})`;
            if (!m[4]) stack.push({ tag: m[2], ns: own });
            continue;
        }
        if (m[4]) continue;
        const top = stack.pop();
        if (top?.tag !== m[2]) return `</${m[2]}> tidak cocok dengan <${top?.tag ?? "?"}> (karakter ${m.index})`;
    }
    if (stack.length) return `<${stack[stack.length - 1].tag}> tidak ditutup`;
    const amp = /&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.exec(xml);
    if (amp) return `karakter & tanpa escape (karakter ${amp.index})`;
    return null;
}

async function hardChecks(zip: JSZip, docXml: string): Promise<string[]> {
    const fails: string[] = [];

    for (const name of Object.keys(zip.files)) {
        if (!/\.(xml|rels)$/.test(name)) continue;
        const p = xmlProblem(await zip.file(name)!.async("string"));
        if (p) fails.push(`XML rusak di ${name}: ${p}`);
    }

    const rels = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "";
    const ids = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
    const missing = [...new Set([...docXml.matchAll(/r:(?:id|embed|link)="([^"]+)"/g)].map((m) => m[1]))].filter((id) => !ids.has(id));
    if (missing.length) fails.push(`Relasi tidak ada di document.xml.rels: ${missing.join(", ")}`);

    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const target = /Target="([^"]+)"/.exec(m[0])?.[1];
        if (!target || /TargetMode="External"/.test(m[0]) || /^[a-z]+:/i.test(target)) continue;
        const file = path.posix.normalize(target.startsWith("/") ? target.slice(1) : path.posix.join("word", target));
        if (!zip.file(file)) fails.push(`File relasi tidak ada: ${file}`);
    }

    const types = (await zip.file("[Content_Types].xml")?.async("string")) ?? "";
    for (const name of Object.keys(zip.files)) {
        if (!/^word\/(header|footer)[^/]*\.xml$/.test(name)) continue;
        if (!types.includes(`PartName="/${name}"`)) fails.push(`Content type tidak terdaftar untuk ${name}`);
    }

    const stray = [...docXml.matchAll(/<w:t(?:\s[^>]*)?>\s*(period|space|undefined|null|NaN|\[object Object\])\s*<\/w:t>/g)];
    if (stray.length) fails.push(`${stray.length} run berisi teks nyasar "${stray[0][1]}" (nilai config/variabel tertulis ke dokumen)`);

    // Aturan umum (semua naskah):
    // 1) Judul bernomor empat tingkat ("2.1.4.1 …") bukan subbab: tidak ber-style Heading 1–3.
    const paras = [...docXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((m) => m[0]);
    const plain = (x: string) => [...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g)].map((m) => m[1] ?? "\t").join("");
    const deep = paras.filter((x) => /^\s*\d+\.\d+\.\d+\.\d+\.?[\t ]/.test(plain(x)) && /<w:pStyle w:val="(?:Heading|heading)[1-3]"/.test(x));
    if (deep.length) fails.push(`${deep.length} judul empat tingkat ber-style Heading 1–3, mis. "${plain(deep[0]).slice(0, 40)}"`);
    // 2) Daftar Isi (entri yang menunjuk ke judul lewat PAGEREF) paling dalam tiga tingkat.
    const tocDeep = paras.filter((x) => /PAGEREF _(?:Fmt|Toc)/.test(x) && /^\s*\d+\.\d+\.\d+\.\d+\.?\t/.test(plain(x)));
    if (tocDeep.length) fails.push(`Daftar Isi memuat ${tocDeep.length} judul tingkat 4, mis. "${plain(tocDeep[0]).slice(0, 40)}"`);
    // 3) Paragraf ber-indentasi baku formatter tidak boleh masih menunjuk ke kotak HTML bermargin.
    const ws = zip.file("word/webSettings.xml") ? await zip.file("word/webSettings.xml")!.async("string") : "";
    if (/<w:marLeft w:val="[1-9]/.test(ws)) {
        const stuck = paras.filter((x) => /<w:divId\b/.test(x) && /<w:ind w:left="720" w:hanging="720"\/>/.test(x));
        if (stuck.length) fails.push(`${stuck.length} paragraf (mis. Daftar Pustaka) berindentasi baku masih menunjuk ke kotak HTML bermargin → bergeser di Word`);
    }
    // 4) Catatan tabel ("Keterangan", "Sumber") menempel ke tabelnya: paragraf tepat setelah tabel
    //    yang berupa catatan harus keepNext bila masih ada catatan sesudahnya.
    //    Hanya di badan teks sebelum lampiran (lampiran beku & isi tabel tata letak tidak diubah).
    const refAt = docXml.search(/>DAFTAR PUSTAKA<\/w:t>(?![\s\S]*>DAFTAR PUSTAKA<\/w:t>)/);
    for (const m of docXml.matchAll(/<\/w:tbl>((?:<w:p\b[^>]*>[\s\S]*?<\/w:p>){2})/g)) {
        const before = docXml.slice(0, m.index! + "</w:tbl>".length);
        if (refAt !== -1 && m.index! > refAt) continue;
        if ((before.match(/<w:tbl>/g) ?? []).length - (before.match(/<\/w:tbl>/g) ?? []).length > 0) continue;
        const ps = [...m[1].matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((x) => x[0]);
        if (/^\s*Keterangan\s*:?\s*$/.test(plain(ps[0])) && plain(ps[1]).trim() && !/<w:keepNext\/>/.test(ps[0])) {
            fails.push(`"Keterangan" tabel tidak menempel ke baris catatan berikutnya (bisa terpisah halaman)`);
            break;
        }
    }
    // 5) Blok catatan tabel (baris berhuruf kecil tepat di bawah tabel): tidak ada jarak di tengah,
    //    dan tidak memuat judul isi lain ("Kaidah Keputusan:").
    let gapFail = "";
    let alienFail = "";
    for (const m of docXml.matchAll(/<\/w:tbl>((?:(?:<w:bookmark(?:Start|End)\b[^>]*\/>)|<w:p\b[^>]*>[\s\S]*?<\/w:p>)+)/g)) {
        const before = docXml.slice(0, m.index! + "</w:tbl>".length);
        if (refAt !== -1 && m.index! > refAt) continue;
        if ((before.match(/<w:tbl>/g) ?? []).length - (before.match(/<\/w:tbl>/g) ?? []).length > 0) continue;
        const ps = [...m[1].matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)].map((x) => x[0]);
        const block: string[] = [];
        for (const p of ps) {
            const sz = Number(/<w:sz w:val="(\d+)"/.exec(p)?.[1] ?? 0);
            if (!plain(p).trim() || !sz || sz >= 24) break;
            block.push(p);
        }
        if (block.length < 2) continue;
        const sp = (p: string, k: "before" | "after") => Number(new RegExp(`<w:spacing\\b[^>]*w:${k}="(\\d+)"`).exec(p)?.[1] ?? 0);
        if (!gapFail && (block.slice(1).some((p) => sp(p, "before") > 0) || block.slice(0, -1).some((p) => sp(p, "after") > 0))) {
            gapFail = `ada jarak kosong di tengah blok catatan tabel, mis. sekitar "${plain(block[0]).slice(0, 30)}"`;
        }
        const alien = block.find((p) => /:\s*$/.test(plain(p).trim()) && !/^\s*(keterangan|ket\.?|sumber|source|catatan|note)/i.test(plain(p)));
        if (alien && !alienFail) alienFail = `catatan tabel memuat judul isi lain: "${plain(alien).slice(0, 40)}"`;
    }
    if (gapFail) fails.push(gapFail);
    if (alienFail) fails.push(alienFail);
    // 7) Daftar Tabel / Ilustrasi memuat SEMUA keterangan bernomor di isi bab (termasuk keterangan di
    //    sel tabel tata letak & kotak teks). Dibandingkan per nomor unik.
    if (refAt !== -1) {
        const mainXml = docXml.slice(0, refAt);
        const listSeg = (head: RegExp) => {
            const hs = [...docXml.matchAll(head)];
            if (!hs.length) return null;
            const from = hs[hs.length > 1 ? 1 : 0].index!;             // [0] bisa entri Daftar Isi
            const rest = docXml.slice(from + 10);
            const to = rest.search(/>DAFTAR (?:ISI|TABEL|ILUSTRASI|GAMBAR|LAMPIRAN)<\/w:t>|>(?:BAB )?I<\/w:t>/);
            return rest.slice(0, to === -1 ? 20000 : to);
        };
        const nums = (label: RegExp) => {
            const set = new Set<string>();
            for (const m of mainXml.matchAll(/<w:p\b[^>]*>(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g)) {
                const t = plain(m[0]).trim();
                const k = label.exec(t);
                if (k && t.length < 250) set.add(k[1]);
            }
            return set;
        };
        for (const [head, label, name] of [
            [/>DAFTAR (?:ILUSTRASI|GAMBAR)<\/w:t>/g, /^(?:ilustrasi|gambar)\s+(\d+)\s*[.:]/i, "Ilustrasi/Gambar"],
            [/>DAFTAR TABEL<\/w:t>/g, /^tabel\s+(\d+)\s*[.:]/i, "Tabel"],
        ] as const) {
            const seg = listSeg(head);
            if (seg === null) continue;
            const listed = (seg.match(/PAGEREF /g) ?? []).length;
            const body = nums(label).size;
            if (body && listed < body) fails.push(`Daftar ${name} memuat ${listed} entri, padahal di isi bab ada ${body} keterangan bernomor`);
        }
    }
    // 8) Kalimat pengantar "…:" menempel ke daftar / keterangan "simbol = arti" sesudahnya; rumus yang
    //    diikuti kalimat dalam satu paragraf sudah dipisah (isi bab, sebelum Daftar Pustaka).
    if (refAt !== -1) {
        const mainParas = [...docXml.slice(0, refAt).matchAll(/<w:p\b[^>]*>(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g)].map((m) => m[0]);
        const loose = mainParas.find((p, k) => {
            const nx = mainParas[k + 1];
            if (!nx || !/:\s*$/.test(plain(p).trim()) || plain(p).trim().length < 8) return false;
            const intro = /<w:numId w:val="[1-9]/.test(nx) || /^\S{1,22}\t=\t/.test(plain(nx));
            return intro && !/<w:keepNext\/>/.test(p) && /<w:sz w:val="24"/.test(p);
        });
        if (loose) fails.push(`kalimat pengantar tidak menempel ke daftar sesudahnya: "${plain(loose).slice(0, 40)}"`);
        const glued = mainParas.find((p) => /^(?:<w:pPr>[\s\S]*?<\/w:pPr>)?(?:<w:bookmark\w+[^>]*\/>)*<m:oMath\b/.test(p.replace(/^<w:p\b[^>]*>/, "")) &&
            /^[A-Z]\S*\s+\S+\s+\S+/.test(plain(p.slice(p.lastIndexOf("</m:oMath>"))).trim()) &&
            !/<m:oMath\b/.test(p.slice(p.indexOf("</m:oMath>") + 10)));
        if (glued) fails.push(`rumus dan kalimat masih dalam satu paragraf: "${plain(glued).slice(0, 40)}"`);
    }
    // 6) Baris keterangan rumus "simbol ⇥ = arti" tidak ditengahkan seperti rumus.
    const centeredLegend = paras.filter((x) => /<w:jc w:val="center"\/>/.test(/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(x)?.[0] ?? "") &&
        /^\s*\S{1,12}\s*\t\s*=\s*\S/.test(plain(x)) && !/<w:tbl/.test(x));
    if (centeredLegend.length) fails.push(`${centeredLegend.length} baris keterangan rumus ("simbol = arti") ditengahkan, mis. "${plain(centeredLegend[0]).slice(0, 40)}"`);

    return fails;
}

/** Peringatan: kemungkinan halaman kosong (dua pindah halaman tanpa isi di antaranya). */
function blankPageWarnings(blocks: (ParaNode | TableNode)[]): string[] {
    const out: string[] = [];
    let content = true;
    for (const n of blocks) {
        if (n.kind === "table") { content = true; continue; }
        const pPr = ownPPr(n.xml);
        if (/<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(pPr) && !content) {
            out.push(`kemungkinan halaman kosong sebelum "${textOf(n.xml).slice(0, 40)}"`);
        }
        if (textOf(n.xml) || /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/.test(n.xml)) content = true;
        const sect = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0];
        const breaks = /<w:br\b[^>]*w:type="page"/.test(n.xml) || (sect && !/<w:type\s+w:val="continuous"/.test(sect));
        if (breaks) {
            if (!content) out.push(`kemungkinan halaman kosong di sekitar "${textOf(n.xml).slice(0, 40) || "(paragraf kosong)"}"`);
            content = false;
        }
    }
    return out;
}

/* ------------------------------------------------------------------ */
/* Snapshot & diff                                                      */
/* ------------------------------------------------------------------ */

function diff(a: string[], b: string[]): string[] {
    let s = 0;
    while (s < a.length && s < b.length && a[s] === b[s]) s++;
    let e = 0;
    while (e < a.length - s && e < b.length - s && a[a.length - 1 - e] === b[b.length - 1 - e]) e++;
    const A = a.slice(s, a.length - e);
    const B = b.slice(s, b.length - e);
    const dp = Array.from({ length: A.length + 1 }, () => new Int32Array(B.length + 1));
    for (let i = A.length - 1; i >= 0; i--)
        for (let j = B.length - 1; j >= 0; j--)
            dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    const out: string[] = [];
    let i = 0;
    let j = 0;
    while (i < A.length && j < B.length) {
        if (A[i] === B[j]) { i++; j++; }
        else if (dp[i + 1][j] >= dp[i][j + 1]) out.push(`  - ${A[i++]}`);
        else out.push(`  + ${B[j++]}`);
    }
    while (i < A.length) out.push(`  - ${A[i++]}`);
    while (j < B.length) out.push(`  + ${B[j++]}`);
    return out;
}

/* ------------------------------------------------------------------ */
/* Main                                                                 */
/* ------------------------------------------------------------------ */

const normName = (s: string) => s.toLowerCase().replace(/\.docx$/, "").replace(/[\s_\-]+/g, " ").trim();

async function main() {
    if (!existsSync(ROOT)) { console.error(`Folder ${ROOT} belum ada.`); process.exit(1); }
    await mkdir(OUT, { recursive: true });
    await mkdir(SNAP, { recursive: true });
    const cases: Record<string, Case> = existsSync(path.join(ROOT, "cases.json"))
        ? JSON.parse(await readFile(path.join(ROOT, "cases.json"), "utf8")) : {};

    const all = (await readdir(ROOT)).filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$"));
    // Nama di cases.json dicocokkan longgar: huruf besar/kecil, spasi/_/- dianggap sama, boleh sebagian nama.
    const used = new Set<string>();
    const caseOf = (file: string): Case => {
        const k = normName(file);
        const hit = Object.keys(cases).find((key) => normName(key) === k) ??
            Object.keys(cases).find((key) => k.includes(normName(key)));
        if (hit) { used.add(hit); return cases[hit]; }
        return {};
    };
    const files = all.filter((f) => !FILTER.length || FILTER.some((x) => f.toLowerCase().includes(x)));
    if (!files.length) { console.error("Tidak ada .docx di test-docs/."); process.exit(1); }

    let failed = 0;
    let changed = 0;
    for (const file of files) {
        const c = caseOf(file);
        const preset = c.preset ?? "umum";
        const config = { ...configForPreset(preset), ...(c.config ?? {}) } as Parameters<typeof transformDocx>[1];
        const base = file.replace(/\.docx$/i, "");
        const started = Date.now();
        let lines: string[] = [];
        const fails: string[] = [];
        let warns: string[] = [];
        const stages: { name: string; xml: string }[] = [];
        let originalXml = "";
        let finalXml = "";
        let unexplained: { a: string[]; b: string[] } | null = null;
        let saveNote = "";

        try {
            const input = await readFile(path.join(ROOT, file));
            originalXml = (await (await JSZip.loadAsync(input)).file("word/document.xml")?.async("string")) ?? "";
            const frontAI = c.frontAI
                ? async () => JSON.parse(await readFile(path.join(ROOT, c.frontAI!), "utf8"))
                : undefined;
            const result = await transformDocx(input, config, {
                ...(TRACE ? { trace: (name: string, xml: string) => stages.push({ name, xml }) } : {}),
                frontAI,
            });
            const outFile = path.join(OUT, `${base}-FORMATTED-${preset}.docx`);
            try {
                await writeFile(outFile, result.buffer);
            } catch (e) {
                const code = (e as NodeJS.ErrnoException).code;
                if (code !== "EBUSY" && code !== "EPERM") throw e;
                // File hasil lama sedang dibuka di Word: simpan dengan nama lain, pemeriksaan tetap jalan.
                const alt = outFile.replace(/\.docx$/, `-${Date.now()}.docx`);
                await writeFile(alt, result.buffer);
                saveNote = `file hasil sedang dibuka di Word, disimpan sebagai ${path.basename(alt)}`;
            }
            const zip = await JSZip.loadAsync(result.buffer);
            finalXml = await zip.file("word/document.xml")!.async("string");

            for (const i of result.issues) if (i.level === "error") fails.push(`Verifikasi: ${i.message}`);
            unexplained = result.unexplained;
            fails.push(...await hardChecks(zip, finalXml));

            const blocks = flatten(parseDocument(finalXml).body);
            warns = blankPageWarnings(blocks);
            const auto = countAutoSpacing(parseDocument(finalXml), (await zip.file("word/styles.xml")?.async("string")) ?? "");
            if (auto.count) warns.push(`${auto.count} paragraf masih memakai spasi otomatis dari style (${auto.styles.join(", ")}) — celah tersembunyi`);
            lines = [
                `# preset: ${preset}`,
                ...describeSections(finalXml),
                ...blocks.map((n) => (n.kind === "paragraph" ? describePara(n) : describeTable(n))),
                ...result.notes.map((n) => `# catatan: ${n}`),
            ];
        } catch (err) {
            fails.push(`Formatter error: ${(err as Error).stack ?? err}`);
        }

        const snapPath = path.join(SNAP, `${base}.${preset}.txt`);
        const label = `${file} [${preset}] (${((Date.now() - started) / 1000).toFixed(1)} dtk)`;
        let d: string[] = [];
        if (lines.length) {
            if (UPDATE || !existsSync(snapPath)) {
                await writeFile(snapPath, lines.join("\n") + "\n");
            } else {
                d = diff((await readFile(snapPath, "utf8")).trimEnd().split("\n"), lines);
            }
        }

        if (fails.length) failed++;
        if (d.length) changed++;
        const status = fails.length ? "GAGAL" : d.length ? "BERUBAH" : UPDATE ? "DISETUJUI" : "OK";
        console.log(`\n${status.padEnd(9)} ${label}`);
        for (const f of fails) console.log(`  ✗ ${f}`);
        if (unexplained) {
            const dd = diff(unexplained.a, unexplained.b);
            console.log(`  baris yang berubah TANPA dilaporkan modul (${dd.length}):`);
            for (const l of dd.slice(0, 20)) console.log(l);
            if (dd.length > 20) console.log(`  … ${dd.length - 20} lagi`);
        }
        if (saveNote) console.log(`  ! ${saveNote}`);
        for (const w of warns.slice(0, 5)) console.log(`  ! ${w}`);
        if (d.length) {
            console.log(`  ${d.length} baris berbeda dari snapshot:`);
            for (const l of d.slice(0, 80)) console.log(l);
            if (d.length > 80) console.log(`  … ${d.length - 80} baris lagi`);
        }

        // Jejak per tahap: baris teks (versi verifikasi) yang berubah di tiap tahap formatter.
        if (TRACE && originalXml) {
            console.log("  --- jejak per tahap ---");
            const linesOf = (xml: string) => fingerprint(xml).text.split("\n");
            let prev = { name: "00 asli", lines: linesOf(originalXml) };
            for (const s of [...stages, { name: "12 hasil akhir", xml: finalXml }]) {
                if (!s.xml) continue;
                const cur = linesOf(s.xml);
                const dd = diff(prev.lines, cur);
                if (dd.length) {
                    console.log(`  [${prev.name} → ${s.name}] ${dd.length} baris berubah`);
                    for (const l of dd.slice(0, 14)) console.log(`  ${l}`);
                    if (dd.length > 14) console.log(`    … ${dd.length - 14} lagi`);
                }
                prev = { name: s.name, lines: cur };
            }
        }
    }

    const unused = Object.keys(cases).filter((k) => !used.has(k));
    if (unused.length && !FILTER.length) console.log(`\n! cases.json tidak cocok dengan file mana pun: ${unused.join(", ")}`);
    console.log(`\n${files.length} dokumen · ${failed} gagal · ${changed} berubah${changed && !UPDATE ? " (jalankan dengan --update bila perubahan memang benar)" : ""}`);
    process.exit(failed ? 1 : 0);
}

main();
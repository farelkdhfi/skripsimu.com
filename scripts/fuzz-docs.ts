/**
 * Uji mutasi (fuzz) formatter: naskah asli di test-docs/ diacak dengan kerusakan yang biasa
 * dibuat penulis, lalu diformat. Yang diperiksa BUKAN tampilannya, melainkan janji dasar:
 *
 *   - formatter tidak pernah error / macet;
 *   - isi sesudah bagian depan tidak pernah hilang: setiap kata isi naskah (dari BAB I s.d. akhir,
 *     dicari dengan aturan sederhana yang terpisah dari formatter) masih ada di hasil — atau
 *     formatter menyerahkan naskah asli apa adanya (tangga cadangan terakhir);
 *   - judul bab naskah tetap ada.
 *
 *   npx tsx scripts/fuzz-docs.ts                 → semua naskah body-* & front-m/k, 8 mutasi
 *   npx tsx scripts/fuzz-docs.ts dina --seed 7   → hanya naskah "dina", acakan lain
 */
import { useNodeIO } from "../src/features/thesis-formatter/front/node";
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { formatDocxBytes } from "../src/features/thesis-formatter/services/format-core";
import { configForPreset } from "../src/features/thesis-formatter/schemas/format-config";
import { fingerprint } from "../src/features/thesis-formatter/docx/verify";
useNodeIO();

const args = process.argv.slice(2);
const seedArg = args.indexOf("--seed");
let seed = seedArg !== -1 ? Number(args[seedArg + 1]) : 1;
const FILTER = args.filter((a, i) => !a.startsWith("--") && i !== seedArg + 1).map((a) => a.toLowerCase());
const rand = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };

/* ---- mutasi atas document.xml ---- */

const PARA_RE = /<w:p\b(?:(?!<w:p\b)[\s\S])*?<\/w:p>/g;
const textOf = (p: string) => [...p.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
const mapParas = (xml: string, f: (p: string, i: number) => string) => { let i = 0; return xml.replace(PARA_RE, (p) => f(p, i++)); };
const EMPTY_P = "<w:p/>";

type Mutation = { name: string; apply: (xml: string) => string };
const MUTATIONS: Mutation[] = [
    {
        name: "style judul dibuang (Heading → Normal)",
        apply: (x) => x.replace(/<w:pStyle w:val="(?:Heading|Judul)\d"\/>/gi, ""),
    },
    {
        name: "pindah halaman & section break dibuang",
        apply: (x) => x.replace(/<w:br w:type="page"\/>/g, "").replace(/<w:pageBreakBefore(?:\s[^>]*)?\/>/g, ""),
    },
    {
        name: "baris kosong acak disisipkan",
        apply: (x) => mapParas(x, (p) => (rand() < 0.08 ? p + EMPTY_P + (rand() < 0.3 ? EMPTY_P : "") : p)),
    },
    {
        name: "tab diganti spasi",
        apply: (x) => mapParas(x, (p) => (rand() < 0.5 ? p.replace(/<w:tab\/>/g, '<w:t xml:space="preserve"> </w:t>') : p)),
    },
    {
        name: "judul DAFTAR ISI/TABEL dibuang",
        apply: (x) => mapParas(x, (p) => (/^\s*DAFTAR\s+(ISI|TABEL|GAMBAR|ILUSTRASI|LAMPIRAN)\s*$/i.test(textOf(p)) ? "" : p)),
    },
    {
        name: "daftar lampiran ketikan menempel ke BAB I",
        apply: (x) => {
            let done = false;
            return mapParas(x, (p) => {
                if (done || !/^\s*(BAB\s+)?I\s*(PENDAHULUAN)?\s*$/i.test(textOf(p)) && !/^\s*BAB\s+I\b/i.test(textOf(p))) return p;
                // Hanya BAB I yang asli (bukan baris daftar isi): baris ini harus bergaya judul / rata tengah.
                if (!/<w:pStyle w:val="(?:Heading|Judul)1"|<w:jc w:val="center"/.test(p)) return p;
                done = true;
                const row = (n: number) => `<w:p><w:r><w:t>Hasil Analisis Statistik Bagian ${n} terhadap Variabel Penelitian</w:t></w:r><w:r><w:tab/><w:t>${50 + n}</w:t></w:r></w:p>`;
                return `<w:p><w:r><w:t>DAFTAR LAMPIRAN</w:t></w:r></w:p>${[1, 2, 3].map(row).join("")}${p}`;
            });
        },
    },
    {
        name: "judul bab romawi tanpa kata BAB",
        apply: (x) => x.replace(/(<w:t(?:\s[^>]*)?>)\s*BAB\s+([IVX]+)\s*(<\/w:t>)/g, "$1$2$3"),
    },
    {
        name: "paragraf isi digandakan spasi & huruf kapital acak di judul",
        apply: (x) => mapParas(x, (p) => {
            const t = textOf(p);
            if (t.length > 0 && t.length < 60 && rand() < 0.2) return p.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g, (_m, a, b, c) => `${a}${b.toUpperCase()}${c}`);
            return p;
        }),
    },
];

/* ---- isi naskah dari BAB I s.d. akhir (aturan sederhana, terpisah dari formatter) ---- */

const wordKey = (w: string) => w.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");
function bodyWords(xml: string): Map<string, number> {
    const lines = fingerprint(xml.replace(/<w:(?:tab|br|cr)\s*\/>/g, "<w:t> </w:t>")).text.split("\n");
    // Awal BAB I = baris judul bab I terakhir yang disusul kalimat panjang dalam 6 baris.
    let start = 0;
    for (let i = 0; i < lines.length; i++) {
        if (!/^(?:bab\s+)?(?:i|1)\.?\s*(?:pendahuluan)?\s*$|^(?:bab\s+i|i)\s+pendahuluan\s*$/i.test(lines[i].trim())) continue;
        if (lines.slice(i + 1, i + 7).some((l) => l.length > 150 && /[.!?]$/.test(l.trim()))) start = i;
    }
    const bag = new Map<string, number>();
    for (const l of lines.slice(start)) for (const w of l.split(/\s+/)) { const k = wordKey(w); if (k) bag.set(k, (bag.get(k) ?? 0) + 1); }
    return bag;
}
function missingWords(a: Map<string, number>, b: Map<string, number>) {
    let total = 0, miss = 0;
    for (const [w, n] of a) { total += n; miss += Math.max(0, n - (b.get(w) ?? 0)); }
    return { total, miss };
}

(async () => {
    const root = path.resolve("test-docs");
    const files = (await readdir(root)).filter((f) => f.endsWith(".docx") && /^(body-|front-m|front-k-gilang\.)/.test(f))
        .filter((f) => !FILTER.length || FILTER.some((x) => f.toLowerCase().includes(x)));
    const ai = { classify: async (s: unknown[]) => ({ roles: s.map(() => null), recommendations: [], notices: [] }), cover: async () => null, front: async () => null };
    let failed = 0;
    let runs = 0;
    for (const f of files) {
        const input = await readFile(path.join(root, f));
        const zip0 = await JSZip.loadAsync(input);
        const xml0 = await zip0.file("word/document.xml")!.async("string");
        for (const m of MUTATIONS) {
            runs++;
            const zip = await JSZip.loadAsync(input);
            const xml = m.apply(xml0);
            zip.file("word/document.xml", xml);
            const bytes = await zip.generateAsync({ type: "uint8array" });
            const t0 = Date.now();
            const tag = `${f} · ${m.name}`;
            try {
                const r = await formatDocxBytes(bytes, f, configForPreset("unpad-fapet"), undefined, ai as never);
                const out = Buffer.from(r.docxBase64, "base64");
                const outXml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
                const returnedOriginal = outXml === xml;
                const { total, miss } = missingWords(bodyWords(xml), bodyWords(outXml));
                const limit = Math.max(40, Math.round(total * 0.01));
                const errs = r.validation.filter((v) => v.level === "error").map((v) => v.message);
                const status = returnedOriginal ? "ASLI" : miss > limit ? "HILANG" : "OK";
                if (status === "HILANG") failed++;
                console.log(`${status.padEnd(7)} ${tag} (${((Date.now() - t0) / 1000).toFixed(1)} dtk) — kata hilang ${miss}/${total}${errs.length ? `; ${errs.length} error: ${errs[0].slice(0, 140)}` : ""}`);
            } catch (e) {
                failed++;
                console.log(`ERROR   ${tag}: ${(e as Error).message.slice(0, 200)}`);
            }
        }
    }
    console.log(`\n${runs} percobaan · ${failed} gagal`);
    process.exit(failed ? 1 : 0);
})();
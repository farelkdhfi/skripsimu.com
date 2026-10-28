/**
 * Pembacaan data bagian depan dengan bantuan AI.
 *
 * AI hanya MENUNJUK baris (dan gambar) di naskah — tidak pernah menulis isi. Setiap nilai yang
 * dipakai disalin dari baris naskah itu sendiri, lalu diperiksa aturan (NPM harus angka, nama
 * tanpa angka, rentang paragraf masuk akal, …). Tunjukan yang tidak lolos diabaikan dan
 * pembacaan aturan dipakai sebagai gantinya.
 */
import { z } from "zod";
import type { ParaNode } from "../docx/document-model";
import { collapseRepeat, isSidangDateLabel, joinRich, isPersonName, npmOf } from "./extract";
import { ABSTRACT_RE, ABSTRAK_RE, APPROVAL_RE, DOC_TYPE_RE, INST_RE, KP_RE, OLEH_RE, YEAR_RE, bare, norm } from "./segment";
import type { FrontFieldKey, FrontPageKind, Item, LineItem, Pic, PicItem, Rich } from "./types";

/* ------------------------------------------------------------------ */
/* Skema jawaban AI                                                     */
/* ------------------------------------------------------------------ */

const num = z.number().int();
const Line = num.nullable().optional().catch(null);
const Lines = z.array(num).nullable().optional().catch(null);
const LineText = z.object({ baris: num, teks: z.string() }).nullable().optional().catch(null);
const Range = z.object({ dari: num, sampai: num }).nullable().optional().catch(null);
const Pics = z.array(z.string()).nullable().optional().catch(null);
const PicId = z.string().nullable().optional().catch(null);

export const frontAiSchema = z.object({
    judul: Lines, judulInggris: Lines,
    nama: LineText, npm: LineText, tahun: LineText,
    pembimbingUtama: LineText, pembimbingAnggota: LineText, ketuaSidang: LineText, dekan: LineText,
    tanggalSidang: LineText, tempatTanggalPernyataan: LineText, tempatTanggalKataPengantar: LineText,
    kataPengantar: Range, abstrak: Range, kataKunci: Line, abstract: Range, keywords: Line,
    ttdPembimbingUtama: PicId, ttdPembimbingAnggota: PicId, ttdKetuaSidang: PicId, ttdDekan: PicId,
    ttdPernyataan: Pics, ttdKataPengantar: Pics,
}).catch({});

export type FrontPointers = z.infer<typeof frontAiSchema>;

/** Pemanggil AI (disediakan service; tidak ada → hanya aturan). */
export type FrontAI = (listing: string) => Promise<FrontPointers | null>;

export const FRONT_AI_SYSTEM = `Kamu membaca BAGIAN DEPAN skripsi Indonesia (sebelum BAB I) yang diberikan sebagai daftar baris bernomor, lalu MENUNJUK baris tempat setiap data berada. Kamu TIDAK menulis ulang isi apa pun.

FORMAT DAFTAR
- "=== HALAMAN <n> (dugaan: <jenis>) ===" menandai awal halaman. Dugaan jenis berasal dari program dan bisa salah; "tidak dikenali" berarti program tidak yakin.
- "L<n>: <teks>" = baris teks bernomor. Baris panjang dipotong dan diakhiri "…".
- "G<n>: [gambar … cm]" = gambar (logo, tanda tangan, materai). "[bentuk/kotak …]" = hiasan, bukan tanda tangan.

ATURAN UMUM (WAJIB)
1. Jangan mengarang. Ragu atau tidak ada → null. Jawaban null lebih baik daripada salah.
2. "teks" harus potongan PERSIS dari baris yang ditunjuk (huruf, ejaan, gelar, tanda baca sama), tanpa label seperti "Nama :", "NPM.", "Dibuat di", "Tanggal Sidang ... :".
3. Setiap data hanya diambil dari SATU tempat. Jangan pernah menggabungkan baris dari dua halaman.

ATURAN PER DATA
- judul: nomor baris judul skripsi berbahasa Indonesia di SAMPUL MUKA (halaman pertama). Hanya baris-baris BERURUTAN di halaman itu. Judul yang sama di sampul kedua, pengesahan, pernyataan, atau abstrak JANGAN ikut ditunjuk. Jangan sertakan baris "SKRIPSI", "Oleh", nama, NPM, atau nama institusi.
- judulInggris: baris judul bahasa Inggris (biasanya miring, tepat di bawah judul Indonesia) di sampul muka; bila tidak ada, baris judul di atas "ABSTRACT". Tetap baris berurutan dari satu halaman. Baris judul bahasa Inggris TIDAK PERNAH ikut ditunjuk sebagai judul, dan sebaliknya: judul dan judulInggris tidak boleh memakai nomor baris yang sama.
- nama: nama mahasiswa penulis (bukan nama pembimbing/penguji/dekan).
- npm: nomor pokok mahasiswa, angka saja.
- tahun: HANYA dari baris yang isinya tahun, mis. "2025" atau "Jatinangor, 2025" di bagian bawah sampul. Angka di dalam NPM (mis. 200110220107) BUKAN tahun.
- pembimbingUtama, pembimbingAnggota, ketuaSidang, dekan: NAMA orang (dengan gelar) di LEMBAR PENGESAHAN, yaitu baris nama yang berpasangan dengan jabatannya (tepat di atas atau di bawah baris jabatan). Baris jabatan bukan nama. Nama penulis tidak pernah menjadi penanda tangan. Bila jabatan di naskah ditulis lain (mis. "An. Ketua Program Studi" atau "Ketua Komisi"), pakai pasangan yang paling sesuai: penanda tangan di bawah "Mengesahkan" yang bukan dekan → ketuaSidang.
- tanggalSidang: HANYA tanggal sidang/ujian sarjana di lembar pengesahan (label seperti "Tanggal Sidang Ujian Sarjana"), tanpa labelnya. Tanggal seminar usulan penelitian ("SUP"), kolokium, atau seminar hasil BUKAN tanggal sidang → null.
- tempatTanggalPernyataan: tempat dan tanggal di lembar pernyataan, mis. "Jatinangor, Juni 2026" (tanpa "Dibuat di").
- tempatTanggalKataPengantar: baris tempat dan tanggal di penutup kata pengantar, mis. "Sumedang, 8 Juni 2026".
- kataPengantar: rentang baris ISI kata pengantar, dari paragraf pertama sesudah judul "KATA PENGANTAR" sampai paragraf terakhir SEBELUM baris tempat-tanggal penutup. Jangan sertakan judul halaman, tempat-tanggal, "Penulis", atau halaman lain.
- abstrak: rentang baris ISI abstrak bahasa Indonesia sesudah judul "ABSTRAK" sampai SEBELUM baris "Kata Kunci". kataKunci: nomor baris "Kata Kunci: ...".
- abstract: sama untuk abstract bahasa Inggris (sesudah "ABSTRACT" sampai sebelum "Keywords"). keywords: nomor baris "Keywords: ...".
- ttdPembimbingUtama, ttdPembimbingAnggota, ttdKetuaSidang, ttdDekan: kode gambar tanda tangan orang itu di lembar pengesahan (gambar tepat di atas namanya), atau null.
- ttdPernyataan: kode gambar tanda tangan dan materai di lembar pernyataan. ttdKataPengantar: kode gambar tanda tangan di penutup kata pengantar.
- Logo (gambar besar di sampul) BUKAN tanda tangan.

CONTOH
Daftar:
=== HALAMAN 1 (dugaan: sampul muka) ===
L1: Pengaruh Pakan Fermentasi terhadap Bobot Badan
L2: Ayam Broiler
L3: The Effect of Fermented Feed on Broiler Body Weight
L4: SKRIPSI
L5: Siti Rahma
G1: [gambar 4.0×4.0 cm]
L6: Jatinangor
L7: 2025
=== HALAMAN 2 (dugaan: sampul kedua) ===
L8: Pengaruh Pakan Fermentasi terhadap Bobot Badan Ayam Broiler
L9: The Effect of Fermented Feed on Broiler Body Weight
L10: SKRIPSI
L11: Siti Rahma
L12: NPM. 200110210045
=== HALAMAN 3 (dugaan: lembar pengesahan) ===
L13: Menyetujui:
G2: [gambar 3.0×1.2 cm]
L14: Dr. Ir. Budi Santoso, M.P.
L15: Pembimbing Utama
L16: Mengesahkan:
L17: Prof. Dr. Ani Lestari, M.Si.
L18: Dekan Fakultas Peternakan
Jawaban:
{"judul":[1,2],"judulInggris":[3],"nama":{"baris":5,"teks":"Siti Rahma"},"npm":{"baris":12,"teks":"200110210045"},"tahun":{"baris":7,"teks":"2025"},"pembimbingUtama":{"baris":14,"teks":"Dr. Ir. Budi Santoso, M.P."},"pembimbingAnggota":null,"ketuaSidang":null,"dekan":{"baris":17,"teks":"Prof. Dr. Ani Lestari, M.Si."},"tanggalSidang":null,"tempatTanggalPernyataan":null,"tempatTanggalKataPengantar":null,"kataPengantar":null,"abstrak":null,"kataKunci":null,"abstract":null,"keywords":null,"ttdPembimbingUtama":"G2","ttdPembimbingAnggota":null,"ttdKetuaSidang":null,"ttdDekan":null,"ttdPernyataan":null,"ttdKataPengantar":null}
Perhatikan: judul hanya [1,2] (BUKAN [1,2,8], dan BUKAN [1,2,3] karena L3 judul bahasa Inggris); tahun dari L7, BUKAN dari angka "2001…" di NPM; G1 adalah logo, bukan tanda tangan.

Balas HANYA JSON dengan kunci: judul, judulInggris, nama, npm, tahun, pembimbingUtama, pembimbingAnggota, ketuaSidang, dekan, tanggalSidang, tempatTanggalPernyataan, tempatTanggalKataPengantar, kataPengantar, abstrak, kataKunci, abstract, keywords, ttdPembimbingUtama, ttdPembimbingAnggota, ttdKetuaSidang, ttdDekan, ttdPernyataan, ttdKataPengantar.
Bentuk nilai: judul/judulInggris = [nomor, ...]; data teks = {"baris": nomor, "teks": "..."}; rentang = {"dari": nomor, "sampai": nomor}; kataKunci/keywords = nomor; tanda tangan = "G<n>" atau ["G<n>", ...]; tidak ada = null.`;

/* ------------------------------------------------------------------ */
/* Daftar baris untuk AI                                                */
/* ------------------------------------------------------------------ */

export type Listing = {
    text: string;
    lines: Map<number, LineItem>;
    pics: Map<number, PicItem>;
    /** Nomor halaman (urutan potongan pindah halaman) setiap baris & gambar. */
    pageOfLine: Map<number, number>;
};

const KIND_LABEL: Record<FrontPageKind, string> = {
    sampulMuka: "sampul muka", sampulKedua: "sampul kedua", pengesahan: "lembar pengesahan", pernyataan: "lembar pernyataan",
    kataPengantar: "kata pengantar", abstrak: "abstrak", abstract: "abstract",
};

/** Daftar baris bernomor per halaman; `pageOf` (hasil aturan) dipakai sebagai dugaan jenis halaman. */
export function buildListing(items: Item[], pageOf: (it: Item) => FrontPageKind | undefined = () => undefined): Listing {
    // Halaman = potongan hasil aturan (pindah halaman, judul halaman, tahun penutup sampul): halaman
    // baru juga dimulai bila jenis halaman hasil aturan berganti walau tanpa pindah halaman.
    const groups: Item[][] = [[]];
    let kindNow: FrontPageKind | undefined;
    for (const it of items) {
        const g = groups[groups.length - 1];
        if (it.k === "break") { if (g.length) groups.push([]); kindNow = undefined; continue; }
        if (it.k === "blank") continue;
        const k = pageOf(it);
        if (k && kindNow && k !== kindNow && g.length) groups.push([]);
        if (k) kindNow = k;
        groups[groups.length - 1].push(it);
    }
    const out: string[] = [];
    const lines = new Map<number, LineItem>();
    const pics = new Map<number, PicItem>();
    const pageOfLine = new Map<number, number>();
    let l = 0;
    let g = 0;
    let page = 0;
    for (const grp of groups) {
        if (!grp.length) continue;
        page++;
        const kinds = [...new Set(grp.map(pageOf).filter((k): k is FrontPageKind => Boolean(k)))];
        out.push(`=== HALAMAN ${page} (dugaan: ${kinds.length ? kinds.map((k) => KIND_LABEL[k]).join(" / ") : "tidak dikenali"}) ===`);
        for (const it of grp) {
            if (it.k === "line") {
                lines.set(++l, it);
                pageOfLine.set(l, page);
                const t = it.text.length > 220 ? `${it.text.slice(0, 220)}…` : it.text;
                out.push(`L${l}: ${t}`);
            } else if (it.k === "pic") {
                pics.set(++g, it);
                const cm = (emu: number) => (emu / 360000).toFixed(1);
                out.push(`G${g}: [${it.pic.picture ? "gambar" : "bentuk/kotak"} ${cm(it.pic.cx)}×${cm(it.pic.cy)} cm]`);
            }
        }
    }
    return { text: out.join("\n"), lines, pics, pageOfLine };
}

/* ------------------------------------------------------------------ */
/* Tunjukan AI → data (disalin dari naskah & diperiksa)                 */
/* ------------------------------------------------------------------ */

export type AiData = {
    fields: Partial<Record<FrontFieldKey, Rich | string>>;
    kataPengantar?: ParaNode[];
    abstrak?: { paras: ParaNode[]; keywords: ParaNode | null };
    abstract?: { paras: ParaNode[]; keywords: ParaNode | null };
    signatures: Partial<Record<"pembimbingUtama" | "pembimbingAnggota" | "ketuaSidang" | "dekan", Pic>>;
    pernyataanPics?: Pic[];
    kataPengantarPics?: Pic[];
    /** Tunjukan yang ditolak pemeriksaan (untuk catatan). */
    rejected: string[];
};

const MONTH = /\b(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|january|february|march|may|june|july|august|october|december)\b/i;

const CLEAN: Partial<Record<FrontFieldKey, (t: string) => string>> = {
    nama: (t) => t.replace(/^\s*(nama(\s+mahasiswa)?|oleh|disusun\s+oleh)\s*[:：]?\s*/i, "").replace(/^\(|\)$/g, "").trim(),
    npm: (t) => npmOf(t) ?? npmOf(t.replace(/^\s*(npm|nim)\s*\.?\s*[:：]?\s*/i, "")) ?? t.replace(/\D/g, ""),
    tahun: (t) => /(?:19|20)\d{2}/.exec(t)?.[0] ?? t,
    tanggalSidang: (t) => t.replace(/^[^:：]*[:：]\s*/, "").replace(/[.…:\s]+$/, "").trim(),
    tempatTanggalPernyataan: (t) => t.replace(/^\s*dibuat\s+di\s*/i, "").replace(/[.…\s]+$/, "").trim(),
    tempatTanggalKataPengantar: (t) => t.replace(/\.$/, "").trim(),
};

const VALID: Partial<Record<FrontFieldKey, (t: string) => boolean>> = {
    nama: (t) => isPersonName(t),
    npm: (t) => /^\d{8,15}$/.test(t),
    tahun: (t) => /^(?:19|20)\d{2}$/.test(t) && Number(t) >= 1990 && Number(t) <= new Date().getFullYear() + 1,
    tanggalSidang: (t) => /\d/.test(t) || MONTH.test(t),
    tempatTanggalPernyataan: (t) => /(?:19|20)\d{2}/.test(t) || MONTH.test(t),
    tempatTanggalKataPengantar: (t) => /(?:19|20)\d{2}/.test(t) || MONTH.test(t),
};
const SIGNER: FrontFieldKey[] = ["pembimbingUtama", "pembimbingAnggota", "ketuaSidang", "dekan"];
const DEGREE = /\b(?:Prof|Dr|Drs|Dra|Ir|H|Hj)\.|,\s*[A-Z][A-Za-z]{0,5}\.|\bPh\.?\s?D\b/;
/** Nama penanda tangan: bergelar atau berbentuk nama orang; bukan jabatan/label/nomor. */
const signerOk = (t: string) => t.length >= 3 && t.length <= 120 && /[A-Za-z]{2}/.test(t) &&
    !/^(?:(?:a\.?\s?n|u\.?\s?b|plh|plt)\.?\s+)?(pembimbing|ketua|dekan|wakil|menyetujui|mengesahkan|nip|nidn|npm|nim|nik|tanggal|oleh)\b/i.test(t) &&
    !/[:：]\s*$/.test(t) && !/\d{4,}/.test(t) && (DEGREE.test(t) || isPersonName(t));
/** Baris jabatan penanda tangan (bukti di sekitar nama). */
export const SIGNER_ROLE_RE = /^(?:(?:a\.?\s?n|u\.?\s?b|plh|plt)\.?\s+)?(?:pembimbing|ketua|(?:wakil\s+)?dekan)\b/i;

const ID_WORDS = new Set(["dan", "yang", "pada", "terhadap", "sebagai", "dari", "dengan", "di", "ke", "untuk", "dalam", "oleh", "atau", "serta", "pengaruh", "penambahan", "pemberian", "penggunaan", "tingkat", "kualitas", "produksi", "terhadap", "secara", "pakan", "ransum", "ayam", "sapi", "domba", "kambing", "itik", "kelinci", "daun", "tepung", "kulit", "buah", "cairan", "imbangan", "sisa", "penyusutan", "pertumbuhan", "lama", "waktu", "mencapai"]);
const EN_WORDS = new Set(["the", "of", "and", "on", "in", "as", "to", "for", "with", "by", "effect", "effects", "from", "at", "a", "an", "its", "their", "growth", "production", "quality", "level", "levels", "addition", "use", "supplementation", "feed", "meal", "peel", "peels", "leaf", "leaves", "chicken", "cattle", "sheep", "goat", "rumen", "fluid", "time", "reach", "stage", "residues", "feces", "performance"]);
/** Bahasa teks dari kata-kata umum: "id", "en", atau null (tidak jelas, mis. nama latin saja). */
export function langOf(text: string): "id" | "en" | null {
    const ws = text.toLowerCase().replace(/\([^)]*\)/g, " ").split(/[^a-z\u00C0-\u024F]+/).filter(Boolean);
    let id = 0;
    let en = 0;
    for (const w of ws) { if (ID_WORDS.has(w)) id++; if (EN_WORDS.has(w)) en++; }
    if (id === en) return null;
    return id > en ? "id" : "en";
}

const squash = (s: string) => s.replace(/\s+/g, " ").trim();

/**
 * Pemeriksaan umum satu nilai data (dari AI maupun aturan). `nama` = nama penulis, untuk menolak
 * penanda tangan yang ternyata nama penulis.
 */
export function fieldOk(k: FrontFieldKey, value: string, nama: string | null = null): boolean {
    const v = squash(value);
    if (!v) return false;
    if (k === "judul" || k === "judulInggris") {
        const lang = langOf(v);
        return v.length >= 8 && v.length <= 500 && v.split(" ").length >= 2 && /[A-Za-z]{3}/.test(v) && !notTitle(v) &&
            collapseRepeat([{ t: v, it: false, sub: false, sup: false, b: false }])[0]?.t.length === v.length &&
            // Judul berbahasa Indonesia, judul Inggris berbahasa Inggris.
            !(k === "judul" && lang === "en") && !(k === "judulInggris" && lang === "id");
    }
    if (SIGNER.includes(k)) {
        const same = nama && v.toLowerCase().replace(/[^a-z]/g, "") === nama.toLowerCase().replace(/[^a-z]/g, "");
        return signerOk(v) && !same;
    }
    return (VALID[k] ?? (() => true))(v);
}

/** Halaman asal data yang hanya ada di satu halaman. */
const PAGE_OF_FIELD: Partial<Record<FrontFieldKey, FrontPageKind[]>> = {
    pembimbingUtama: ["pengesahan"], pembimbingAnggota: ["pengesahan"], ketuaSidang: ["pengesahan"], dekan: ["pengesahan"],
    tanggalSidang: ["pengesahan"], tempatTanggalPernyataan: ["pernyataan"], tempatTanggalKataPengantar: ["kataPengantar"],
    tahun: ["sampulMuka", "sampulKedua"],
};

/** Baris yang jelas BUKAN bagian judul (jenis dokumen, label penulis, institusi, tahun, NPM, …). */
const notTitle = (t: string) => {
    const b = bare(t);
    return DOC_TYPE_RE.test(b) || OLEH_RE.test(b) || APPROVAL_RE.test(b) || YEAR_RE.test(b) || (INST_RE.test(b) && b.length < 60) ||
        /^((disusun|diajukan|ditulis)\s+oleh|nama|npm|nim)\b/i.test(b) || /^\d[\d .]{7,}$/.test(b) ||
        KP_RE.test(b) || ABSTRAK_RE.test(b) || ABSTRACT_RE.test(b) || /^(lembar|halaman|surat)\s/i.test(b);
};

/**
 * Tunjukan AI → data. `pageOf` (jenis halaman setiap item, bila halaman itu dikenali aturan)
 * dipakai untuk menolak tunjukan yang jelas salah tempat, mis. logo sampul sebagai tanda tangan.
 */
export function applyPointers(ptr: FrontPointers, listing: Listing, pageOf: (it: Item) => FrontPageKind | undefined = () => undefined): AiData {
    const res: AiData = { fields: {}, signatures: {}, rejected: [] };
    const line = (n: number | null | undefined) => (typeof n === "number" ? listing.lines.get(n) : undefined);
    const reject = (what: string) => res.rejected.push(what);
    /** Item berada di halaman yang benar (atau halamannya tidak dikenali aturan). */
    const on = (it: Item | undefined, ...kinds: FrontPageKind[]) => {
        if (!it) return false;
        const k = pageOf(it);
        return k === undefined || kinds.includes(k);
    };

    // Judul (beberapa baris berurutan). Judul & judul Inggris yang menunjuk baris yang sama:
    // salah satunya pasti keliru → keduanya ditolak (dipakai bacaan aturan).
    const overlap = (ptr.judul ?? []).some((n) => (ptr.judulInggris ?? []).includes(n));
    for (const k of ["judul", "judulInggris"] as const) {
        const ns = ptr[k];
        if (!ns?.length) continue;
        if (overlap) { reject(`${k} (judul & judul Inggris menunjuk baris yang sama)`); continue; }
        const ls = ns.map(line);
        // Satu blok baris BERURUTAN dari SATU halaman (judul yang sama di halaman lain tidak boleh ikut).
        const sorted = [...new Set(ns)].sort((a, b) => a - b);
        const onePage = new Set(sorted.map((n) => listing.pageOfLine.get(n))).size === 1;
        if (ls.some((x) => !x) || !onePage || sorted.some((n, i) => i > 0 && n !== sorted[i - 1] + 1)) { reject(`${k} (bukan baris berurutan di satu halaman)`); continue; }
        if ((ls as LineItem[]).some((l) => notTitle(l.text))) { reject(k); continue; }
        // Setiap baris harus berbahasa yang benar: blok "judul Indonesia + judul Inggris" ditolak.
        const wrong = k === "judul" ? "en" : "id";
        if ((ls as LineItem[]).some((l) => langOf(l.text) === wrong)) { reject(`${k} (ada baris berbahasa ${wrong === "en" ? "Inggris" : "Indonesia"})`); continue; }
        const where: FrontPageKind[] = k === "judul" ? ["sampulMuka", "sampulKedua", "pengesahan"] : ["sampulMuka", "sampulKedua", "abstract"];
        if (!(ls as LineItem[]).every((l) => on(l, ...where))) { reject(k); continue; }
        const segs = (ls as LineItem[]).map((l, i) => {
            if (i > 0) return l.segs;
            const m = /^\s*(?:judul(?:\s+(?:skripsi|penelitian))?|title)\s*[:：]\s*/i.exec(l.raw);
            return m ? cutSegs(l.segs, m[0].length) : l.segs;
        });
        const r = collapseRepeat(joinRich(segs));
        const t = r.map((s) => s.t).join("");
        if (t.length < 8 || t.length > 500 || t.split(/\s+/).length < 2 || !/[A-Za-z]{3}/.test(t)) { reject(k); continue; }
        res.fields[k] = r;
    }

    // Data satu baris: teks harus bagian dari baris yang ditunjuk.
    const singles: FrontFieldKey[] = ["nama", "npm", "tahun", ...SIGNER, "tanggalSidang", "tempatTanggalPernyataan", "tempatTanggalKataPengantar"];
    for (const k of singles) {
        const p = ptr[k as keyof FrontPointers] as { baris: number; teks: string } | null | undefined;
        if (!p) continue;
        const l = line(p.baris);
        if (!l) { reject(`${k} (baris tidak ada)`); continue; }
        // Bukti di sekitar baris lebih kuat daripada dugaan jenis halaman (yang bisa meleset pada naskah
        // tanpa pindah halaman): nama penanda tangan harus menempel pada baris jabatannya; data lain
        // boleh di halaman yang tidak sesuai dugaan asal bukti barisnya sendiri kuat.
        if (SIGNER.includes(k)) {
            const near = [-2, -1, 1, 2].some((d) => { const x = line(p.baris + d); return x && SIGNER_ROLE_RE.test(bare(x.text)); });
            if (!near) { reject(`${k} (tidak ada baris jabatan di dekatnya)`); continue; }
        } else if (PAGE_OF_FIELD[k] && !on(l, ...PAGE_OF_FIELD[k]!)) {
            const t = norm(l.text);
            const next = line(p.baris + 1)?.text ?? "";
            const strong = k === "tanggalSidang" ? /^tanggal\b/i.test(t)
                : k === "tahun" ? YEAR_RE.test(bare(t))
                : k === "tempatTanggalPernyataan" ? /^dibuat\s+di\b/i.test(t)
                : k === "tempatTanggalKataPengantar" ? /^(penulis|peneliti|penyusun)\b/i.test(bare(next))
                : false;
            if (!strong) { reject(`${k} (baris/halaman tidak sesuai)`); continue; }
        }
        // Tanggal sidang hanya dari baris berlabel sidang/ujian sarjana (bukan "Tanggal Pelaksanaan SUP").
        if (k === "tanggalSidang" && !isSidangDateLabel(norm(l.text).split(/[:\uFF1A]/)[0])) { reject("tanggalSidang (bukan tanggal sidang/ujian sarjana)"); continue; }
        // Tahun hanya dari baris tahun ("2025" / "Jatinangor, 2025"), bukan angka di baris lain (NPM).
        if (k === "tahun" && !YEAR_RE.test(bare(l.text))) { reject("tahun (baris yang ditunjuk bukan baris tahun)"); continue; }
        const lt = squash(l.raw.replace(/\t/g, " "));
        const given = squash(p.teks ?? "");
        let v = given && lt.includes(given) ? given : lt;
        v = (CLEAN[k] ?? ((t: string) => t.trim()))(v);
        const ok = SIGNER.includes(k) ? signerOk(v) : (VALID[k] ?? (() => v.length > 0))(v);
        if (!ok) { reject(k); continue; }
        res.fields[k] = v;
    }

    // Rentang paragraf prosa.
    const paras = (r: { dari: number; sampai: number } | null | undefined, head: RegExp, kind: FrontPageKind): ParaNode[] | null => {
        if (!r || r.dari > r.sampai || !listing.lines.has(r.dari) || !listing.lines.has(r.sampai)) return null;
        const out: ParaNode[] = [];
        for (let n = r.dari; n <= r.sampai; n++) {
            const l = listing.lines.get(n);
            if (!l) continue;
            if (!on(l, kind)) return null;
            const t = norm(l.text);
            if (head.test(bare(t)) || /^(kata\s*kunci|key\s*words?)\s*[:：]/i.test(t)) continue;
            if (!out.includes(l.para)) out.push(l.para);
        }
        return out.length ? out : null;
    };
    const kp = paras(ptr.kataPengantar, KP_RE, "kataPengantar");
    if (kp) res.kataPengantar = kp; else if (ptr.kataPengantar) reject("kataPengantar");
    for (const [k, kw, head] of [["abstrak", "kataKunci", ABSTRAK_RE], ["abstract", "keywords", ABSTRACT_RE]] as const) {
        const ps = paras(ptr[k], head, k);
        const kl = line(ptr[kw]);
        const kwOk = kl && /^(kata\s*kunci|key\s*words?)\s*[:：]/i.test(norm(kl.text));
        if (ps) res[k] = { paras: ps.filter((p) => p !== kl?.para), keywords: kwOk ? kl!.para : null };
        else if (ptr[k]) reject(k);
    }

    // Gambar hanya dari halamannya sendiri (logo sampul bukan tanda tangan).
    const pic = (id: string | null | undefined, kind: FrontPageKind, what: string): Pic | null => {
        if (!id) return null;
        const n = Number(/^G?(\d+)$/i.exec(id.trim())?.[1]);
        const it = listing.pics.get(n);
        const ok = it && it.pic.picture && !it.pic.unsupported && on(it, kind);
        if (!ok) reject(what);
        return ok ? it!.pic : null;
    };
    const ttd = { pembimbingUtama: ptr.ttdPembimbingUtama, pembimbingAnggota: ptr.ttdPembimbingAnggota, ketuaSidang: ptr.ttdKetuaSidang, dekan: ptr.ttdDekan };
    const usedPic = new Set<Pic>();
    for (const [k, id] of Object.entries(ttd) as [keyof AiData["signatures"], string | null | undefined][]) {
        const p = pic(id, "pengesahan", `tanda tangan ${k}`);
        if (p && !usedPic.has(p)) { res.signatures[k] = p; usedPic.add(p); }
    }
    if (ptr.ttdPernyataan) res.pernyataanPics = ptr.ttdPernyataan.map((id) => pic(id, "pernyataan", "tanda tangan pernyataan")).filter((p): p is Pic => Boolean(p));
    if (ptr.ttdKataPengantar) res.kataPengantarPics = ptr.ttdKataPengantar.map((id) => pic(id, "kataPengantar", "tanda tangan kata pengantar")).filter((p): p is Pic => Boolean(p));
    return res;
}

function cutSegs(r: Rich, n: number): Rich {
    const out: Rich = [];
    let left = n;
    for (const s of r) {
        if (left >= s.t.length) { left -= s.t.length; continue; }
        out.push({ ...s, t: s.t.slice(left) });
        left = 0;
    }
    return out.filter((s) => s.t);
}
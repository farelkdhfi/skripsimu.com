/**
 * Pemotongan bagian depan menjadi halaman dan penentuan jenis tiap halaman.
 *
 * Batas halaman diambil dari pindah halaman / section break / "page break before". Karena banyak
 * penulis berpindah halaman cukup dengan Enter, ada dua batas cadangan yang aman:
 *  - judul halaman yang khas (KATA PENGANTAR, LEMBAR PERNYATAAN, LEMBAR PENGESAHAN),
 *  - baris tahun penutup sampul ("2026") bila sesudahnya masih ada isi.
 *
 * Pembacaan ini tidak pernah membatalkan template: halaman yang tidak dikenali dilewati (barisnya
 * dilaporkan), halaman ganda hanya dipakai yang pertama, dan halaman yang memuat dua jenis isi
 * (pindah halaman lupa dibuat) dipotong pada judul halaman berikutnya.
 */
import { CHAPTER_RE } from "../structure/roles";
import type { FrontPage, FrontPageKind, Item, LineItem } from "./types";

export const KP_RE = /^(kata\s+pengantar|prakata)$/i;
export const ABSTRAK_RE = /^abstrak$/i;
export const ABSTRACT_RE = /^abstract$/i;
export const PERNYATAAN_RE = /^(?:(?:lembar|surat|halaman)\s+)?pernyataan\b/i;
export const PENGESAHAN_HEAD_RE = /^(?:lembar|halaman)\s+pengesahan\b/i;
export const OLEH_RE = /^oleh\s*:?$/i;
export const APPROVAL_RE = /^(menyetujui|mengesahkan|disetujui|disahkan|pembimbing)\b/i;
export const DOC_TYPE_RE = /^(skripsi|usulan\s+penelitian|hasil\s+penelitian|tugas\s+akhir|artikel|skripsi\s+atau\s+artikel|thesis|tesis|proposal(?:\s+(?:penelitian|skripsi))?|research\s+proposal|undergraduate\s+thesis)$/i;
export const INST_RE = /\b(program\s+studi|prodi|jurusan|departemen|fakultas|universitas|institut|sekolah\s+tinggi|politeknik|degree\s+program|faculty\s+of|university)\b/i;
export const YEAR_RE = /^(?:[A-Za-z][A-Za-z .'-]{1,30},?\s+)?(?:19|20)\d{2}$/;

export const norm = (s: string) => s.replace(/\s+/g, " ").trim();
/** Untuk mencocokkan jenis halaman: tanda bintang, titik dua, titik di ujung diabaikan. */
export const bare = (s: string) => norm(s).replace(/[*:.\s]+$/, "").replace(/^[*\s]+/, "");

const isLine = (x: Item): x is LineItem => x.k === "line";

/** Awal BAB I: "BAB I …", "I. PENDAHULUAN", "I" ↵ "PENDAHULUAN", atau "PENDAHULUAN". */
export function isChapterStart(text: string, next: string | null): boolean {
    const t = bare(text);
    if (t.length > 120) return false;
    if (CHAPTER_RE.test(t)) return true;
    if (/^I\.?\s+[A-Z][A-Z ,&/-]{3,}$/.test(t)) return true;               // "I. PENDAHULUAN"
    if (/^I\.?$/.test(t) && next && /^[A-Z][A-Z ,&/-]{3,}$/.test(bare(next))) return true; // "I" ↵ "PENDAHULUAN"
    if (/^PENDAHULUAN$/.test(t)) return true;
    return false;
}

function signals(items: Item[]) {
    const lines = items.filter(isLine).map((l) => bare(l.text));
    const pics = items.filter((x) => x.k === "pic");
    const has = (re: RegExp) => lines.some((l) => l.length <= 120 && re.test(l));
    return {
        kp: has(KP_RE),
        abstrak: has(ABSTRAK_RE),
        abstract: has(ABSTRACT_RE),
        pernyataan: has(PERNYATAAN_RE),
        pengesahan: has(PENGESAHAN_HEAD_RE) || (has(OLEH_RE) && has(APPROVAL_RE)),
        cover: has(DOC_TYPE_RE) || (pics.some((p) => p.k === "pic" && p.pic.picture) && has(INST_RE)),
        coverSecond: lines.some((l) => /^(diajukan|submitted)\b/i.test(l)) || lines.some((l) => /\b\d{9,15}\b/.test(l.replace(/[\s.]/g, ""))),
        // Tanpa teks dan tanpa foto (paling banyak bentuk/garis hiasan) = halaman kosong.
        empty: !lines.length && !pics.some((p) => p.k === "pic" && p.pic.picture),
    };
}

export type Segmented = {
    pages: FrontPage[];
    blankPages: number;
    blankShapes: Item[];
    /** Halaman yang tidak dipakai (tidak dikenali / ganda): baris teksnya dilaporkan. */
    skipped: { reason: string; items: Item[] }[];
};

const PROSA: FrontPageKind[] = ["kataPengantar", "abstrak", "abstract"];
const CLOSING_RE = /^(penulis|peneliti|penyusun|hormat\s+saya)\b/i;
const KEYWORDS_RE = /^(kata\s*kunci|key\s*words?)\s*[:\uFF1A]/i;

/** Jenis halaman dari isinya (bisa lebih dari satu bila pindah halaman tidak dibuat). */
function kindsOf(group: Item[]): FrontPageKind[] {
    const s = signals(group);
    const kinds: FrontPageKind[] = [];
    if (s.kp) kinds.push("kataPengantar");
    if (s.abstrak) kinds.push("abstrak");
    if (s.abstract) kinds.push("abstract");
    if (s.pernyataan) kinds.push("pernyataan");
    if (s.pengesahan) kinds.push("pengesahan");
    if (s.cover && !kinds.length) kinds.push(s.coverSecond ? "sampulKedua" : "sampulMuka");
    return kinds;
}

/**
 * Halaman berisi kata pengantar + abstrak (atau abstrak + abstract) tanpa pindah halaman: dipotong
 * sebelum blok judul/nama di atas judul ABSTRAK/ABSTRACT, yaitu sesudah baris penutup sebelumnya
 * ("Penulis", baris kata kunci, atau kalimat panjang).
 */
function splitProsa(group: Item[]): Item[][] {
    const cuts: number[] = [];
    group.forEach((it, i) => {
        if (it.k !== "line" || it.inTable) return;
        const t = bare(it.text);
        if (!ABSTRAK_RE.test(t) && !ABSTRACT_RE.test(t)) return;
        let cut = i;
        let seen = 0;
        for (let k = i - 1; k >= 0 && seen < 8; k--) {
            const c = group[k];
            if (c.k !== "line") continue;
            const ct = norm(c.text);
            if (CLOSING_RE.test(ct) && ct.length <= 30) break;
            if (KEYWORDS_RE.test(ct)) break;
            if (ct.length > 100 && /[.!?]$/.test(ct)) break;
            if (KP_RE.test(bare(ct)) || ABSTRAK_RE.test(bare(ct)) || ABSTRACT_RE.test(bare(ct))) break;
            seen++;
            cut = k;
        }
        if (cut > 0) cuts.push(cut);
    });
    if (!cuts.length) return [group];
    const out: Item[][] = [];
    let from = 0;
    for (const c of cuts) { if (c > from) { out.push(group.slice(from, c)); from = c; } }
    out.push(group.slice(from));
    return out;
}

export function segmentPages(items: Item[]): Segmented {
    // 1. Potong pada batas halaman (pindah halaman, section, judul khas, tahun penutup sampul).
    const raw: Item[][] = [[]];
    let closeAfterYear = false;
    const cur = () => raw[raw.length - 1];
    const hasContent = (xs: Item[]) => xs.some((x) => x.k === "line" || x.k === "pic");
    for (const it of items) {
        if (it.k === "break") { if (hasContent(cur())) raw.push([]); closeAfterYear = false; continue; }
        if (it.k !== "blank" && closeAfterYear) { raw.push([]); closeAfterYear = false; }
        if (it.k === "line") {
            const t = bare(it.text);
            const pageHead = KP_RE.test(t) || (PERNYATAAN_RE.test(t) && t.length <= 120) || PENGESAHAN_HEAD_RE.test(t);
            // Hanya bila halaman berjalan sudah berisi TEKS: bentuk/bingkai yang ditambatkan sebelum judul
            // tetap milik halaman judul itu.
            if (pageHead && !it.inTable && cur().some((x) => x.k === "line")) raw.push([]);
            cur().push(it);
            const page = cur();
            if (YEAR_RE.test(t) && page.some((x) => x.k === "line" && INST_RE.test(x.text))) closeAfterYear = true;
            continue;
        }
        cur().push(it);
    }

    // 2. Jenis halaman.
    const pages: FrontPage[] = [];
    const skipped: Segmented["skipped"] = [];
    let blankPages = 0;
    const blankShapes: Item[] = [];
    const groups = raw.flatMap((g) => (kindsOf(g).filter((k) => PROSA.includes(k)).length > 1 ? splitProsa(g) : [g]));
    for (const group of groups) {
        const s = signals(group);
        if (s.empty) {
            if (group.length) blankPages++;
            blankShapes.push(...group.filter((x) => x.k === "pic"));
            continue;
        }
        const kinds = kindsOf(group);
        const first = group.find(isLine)?.text ?? "(halaman tanpa teks)";
        if (!kinds.length) {
            const prev = pages[pages.length - 1];
            // Lanjutan halaman prosa yang terpotong pindah halaman manual.
            if (prev && PROSA.includes(prev.kind)) { prev.items.push(...group); continue; }
            skipped.push({ reason: `halaman yang diawali "${first.slice(0, 60)}" bukan bagian template`, items: group });
            continue;
        }
        // Dua jenis sekaligus: jenis dengan judul halaman paling atas yang dipakai.
        let kind = kinds[0];
        if (kinds.length > 1) {
            const at = (k: FrontPageKind) => group.findIndex((x) => x.k === "line" && kindsOf([x]).includes(k));
            kind = [...kinds].sort((a, b) => (at(a) === -1 ? 1e9 : at(a)) - (at(b) === -1 ? 1e9 : at(b)))[0];
        }
        // Sampul: yang PERTAMA selalu sampul muka, yang kedua sampul kedua (ada naskah yang mencantumkan
        // NPM/"Diajukan…" juga di sampul muka, jadi ciri isi saja tidak cukup).
        const covers = pages.filter((p) => p.kind === "sampulMuka" || p.kind === "sampulKedua").length;
        if (kind === "sampulMuka" || kind === "sampulKedua") kind = covers === 0 ? "sampulMuka" : covers === 1 ? "sampulKedua" : kind;
        if (pages.some((p) => p.kind === kind)) {
            skipped.push({ reason: `halaman "${kind}" kedua (yang dipakai halaman pertama)`, items: group });
            continue;
        }
        pages.push({ kind, items: group });
    }
    return { pages, blankPages, blankShapes, skipped };
}
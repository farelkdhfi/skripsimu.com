/**
 * Baris gabungan → baris terpisah.
 *
 * Naskah hasil konversi PDF → Word (dan sebagian naskah ketikan) menaruh dua "baris" yang tampak
 * terpisah di layar dalam SATU paragraf; pemisahnya hanya lebar paragraf yang sempit:
 *   "Prof. Dr. Ir. Yuli Astuti Hidayati, MP. Pembimbing Utama"   (nama + jabatan)
 *   "Muhammad Omar Bastari 200110220107"                          (nama + NPM)
 *   "OLEH: MUHAMMAD OMAR BASTARI"                                 (label + nama)
 *   "Dibuat di Jatinangor, 13 Oktober 2026 Penulis,"              (tempat-tanggal + penutup)
 * Baris seperti ini dipecah pada penanda yang umum (jabatan penanda tangan, "OLEH", NPM, "Penulis"),
 * supaya aturan & AI melihat baris yang sama dengan yang dilihat pembaca. Hanya baris pendek
 * (bukan alinea) yang diperiksa.
 */
import type { Item, LineItem, Rich } from "./types";

const DEGREE = /\b(?:Prof|Dr|Drs|Dra|Ir|H|Hj)\.|,\s*[A-Z][A-Za-z]{0,5}\.|\bPh\.?\s?D\b/;
const ROLE_START = /(?:^|\s)((?:(?:a\.?\s?n|u\.?\s?b|plh|plt)\.?\s+)?(?:pembimbing\s+(?:utama|anggota|pendamping|pertama|kedua|i{1,2}|[12])|ketua\s+(?:panitia|sidang|penguji|program\s+studi|prodi|departemen|jurusan|komisi)|(?:wakil\s+)?dekan)\b.*)$/i;
const NAME_CHARS = /^[A-Za-zÀ-ɏ .,'’-]+$/;

/** Potong Rich pada posisi karakter `at` (dalam teks gabungan segs). */
function cut(segs: Rich, at: number): [Rich, Rich] {
    const a: Rich = [];
    const b: Rich = [];
    let pos = 0;
    for (const s of segs) {
        const end = pos + s.t.length;
        if (end <= at) a.push(s);
        else if (pos >= at) b.push(s);
        else { a.push({ ...s, t: s.t.slice(0, at - pos) }); b.push({ ...s, t: s.t.slice(at - pos) }); }
        pos = end;
    }
    const trim = (r: Rich) => {
        const x = r.filter((s) => s.t);
        if (x.length) { x[0] = { ...x[0], t: x[0].t.replace(/^\s+/, "") }; x[x.length - 1] = { ...x[x.length - 1], t: x[x.length - 1].t.replace(/\s+$/, "") }; }
        return x.filter((s) => s.t);
    };
    return [trim(a), trim(b)];
}

/** Posisi pemisah (indeks karakter di `raw`) atau -1. */
function splitAt(raw: string): number {
    const t = raw.replace(/\s+$/, "");
    if (t.length > 180) return -1;
    // "OLEH: NAMA" / "Oleh NAMA"
    let m = /^(oleh\s*:?)\s+(?=[A-Za-z\u00C0-\u024F])/i.exec(t);
    if (m) return m[1].length;
    // "Dibuat di Kota, 13 Oktober 2026 Penulis," / "Kota, 8 Juni 2026 Penulis"
    m = /^(.*(?:19|20)\d{2})\s+(?:penulis|peneliti|penyusun|yang\s+menyatakan)\s*,?\s*$/i.exec(t);
    if (m) return m[1].length;
    // "Nama Penulis 200110220107" / "Nama NPM. 200110220107"
    m = /^(.+?)\s+(?:(?:npm|nim)\s*\.?\s*:?\s*)?\d{8,15}$/i.exec(t);
    if (m && NAME_CHARS.test(m[1].trim()) && !/^(npm|nim)\.?$/i.test(m[1].trim()) && m[1].trim().split(/\s+/).length <= 6) return m[1].length;
    // "Nama, Gelar. Pembimbing Utama" / "… IPU. Wakil Dekan I Fakultas Peternakan"
    m = ROLE_START.exec(t);
    if (m && m.index > 0) {
        const left = t.slice(0, m.index).trim();
        const prose = /\b(selaku|kepada|yang|dan|sebagai|terima|atas|oleh)\b/i.test(left);
        if (left.length >= 5 && left.length <= 110 && DEGREE.test(left) && !prose && !/[:：]$/.test(left)) return m.index;
    }
    return -1;
}

export function splitMergedLines(items: Item[]): { items: Item[]; split: number } {
    const out: Item[] = [];
    let split = 0;
    for (const it of items) {
        if (it.k !== "line") { out.push(it); continue; }
        let rest: LineItem = it;
        for (let guard = 0; guard < 3; guard++) {
            const at = splitAt(rest.raw);
            if (at <= 0) break;
            const [a, b] = cut(rest.segs, at);
            const ra = a.map((s) => s.t).join("");
            const rb = b.map((s) => s.t).join("");
            if (!ra.trim() || !rb.trim()) break;
            const mk = (segs: Rich, raw: string): LineItem => ({ ...rest, segs, raw, text: raw.replace(/\t+/g, " ").replace(/\s+/g, " ").trim() });
            out.push(mk(a, ra));
            rest = mk(b, rb);
            split++;
        }
        out.push(rest);
    }
    return { items: out, split };
}
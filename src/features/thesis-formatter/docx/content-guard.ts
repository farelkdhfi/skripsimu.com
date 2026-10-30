/**
 * Pengaman anti-kehilangan isi, terpisah dari verifikasi baris per baris.
 *
 * Verifikasi biasa membandingkan urutan baris dan bisa "dijelaskan" oleh blok yang dilaporkan
 * modul. Pemeriksaan di sini sengaja kasar dan independen: dihitung dari kata, bukan dari laporan
 * modul, sehingga kesalahan besar (satu bab ikut terbuang bersama bagian depan, daftar yang menelan
 * isi bab, catatan kaki hilang) tetap tertangkap walau ada modul yang salah melapor.
 *
 *  1. Kata isi naskah (semua yang tersisa sesudah bagian depan diganti template) harus tetap ada:
 *     selisihnya paling banyak 1% (penomoran ulang, "Gambar" → "Ilustrasi", dsb.; hasil uji ≤ 0,1%).
 *  2. Setiap judul bab & subbab bernomor naskah harus masih ada.
 *  3. Kata di catatan kaki / catatan akhir tidak boleh berkurang.
 *
 * Hasilnya daftar masalah; kosong = aman. Dipakai format-core untuk tangga cadangan (tanpa
 * template bagian depan → naskah asli dikembalikan).
 */
import type JSZip from "jszip";
import { fingerprint } from "./verify";

const wordKey = (w: string) => w.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");
const bag = (text: string) => {
    const m = new Map<string, number>();
    for (const w of text.split(/\s+/)) {
        const k = wordKey(w);
        if (k) m.set(k, (m.get(k) ?? 0) + 1);
    }
    return m;
};
const squash = (s: string) => s.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");
/** Teks per baris dengan tab/ganti baris sebagai spasi (kata yang dipisah tab tidak menempel). */
const textOf = (xml: string) => fingerprint(xml.replace(/<w:(?:tab|br|cr)\s*\/>/g, "<w:t> </w:t>")).text;

export type GuardBaseline = {
    /** Teks isi naskah (sesudah bagian depan diganti template, sebelum modul lain). */
    bodyText: string;
    /** Judul bab/subbab naskah yang wajib tetap ada (teks tanpa nomor). */
    headings: string[];
    /** Kata catatan kaki & catatan akhir naskah asli. */
    notesWords: number;
};

async function notesText(zip: JSZip): Promise<string> {
    const parts = await Promise.all(["word/footnotes.xml", "word/endnotes.xml"].map(async (p) => (await zip.file(p)?.async("string")) ?? ""));
    return parts.map(textOf).join("\n");
}

export async function guardBaseline(zip: JSZip, bodyXml: string, headings: string[]): Promise<GuardBaseline> {
    const notes = await notesText(zip);
    return { bodyText: textOf(bodyXml), headings, notesWords: notes.split(/\s+/).filter((w) => wordKey(w)).length };
}

export async function contentLoss(base: GuardBaseline, zip: JSZip, finalXml: string): Promise<string[]> {
    const out: string[] = [];
    const afterText = textOf(finalXml);

    // 1. Kata isi.
    const a = bag(base.bodyText);
    const b = bag(afterText);
    let total = 0;
    let missing = 0;
    const lost: [string, number][] = [];
    for (const [w, n] of a) {
        total += n;
        const d = n - (b.get(w) ?? 0);
        if (d > 0) { missing += d; lost.push([w, d]); }
    }
    const limit = Math.max(40, Math.round(total * 0.01));
    if (missing > limit) {
        lost.sort((x, y) => y[1] - x[1]);
        out.push(`${missing} kata isi naskah (dari ${total}) tidak ada di hasil, mis. ${lost.slice(0, 6).map(([w]) => `"${w}"`).join(", ")}.`);
    }

    // 2. Judul bab & subbab.
    const flat = squash(afterText);
    const gone = base.headings.filter((h) => squash(h).length >= 4 && !flat.includes(squash(h)));
    if (gone.length) out.push(`${gone.length} judul bab/subbab naskah tidak ada di hasil, mis. ${gone.slice(0, 3).map((h) => `"${h.slice(0, 50)}"`).join(", ")}.`);

    // 3. Catatan kaki / akhir.
    const notesNow = (await notesText(zip)).split(/\s+/).filter((w) => wordKey(w)).length;
    if (notesNow < base.notesWords) out.push(`Catatan kaki/akhir berkurang: ${base.notesWords} → ${notesNow} kata.`);

    return out;
}
/**
 * Lebar huruf Times New Roman (metrik identik dengan Liberation Serif), per 1000 em, untuk
 * menghitung jumlah baris paragraf yang dibungkus — dipakai untuk mengimbangi judul yang lebih
 * panjang/pendek dari contoh template supaya halaman tidak meluber.
 */

const CHARS = " !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~“”‘’–—…éèáàíóúüöäñçÉ×±°µ·";

const TABLE: Record<"r" | "b" | "i" | "bi", string> = {
    r: "250,333,408,500,500,833,778,180,333,333,500,564,250,333,250,278,500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,480,200,480,541,444,444,333,333,500,1000,1000,444,444,444,444,278,500,500,500,500,444,500,444,611,564,549,400,576,333",
    b: "250,333,555,500,500,1000,833,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,394,220,394,520,500,500,333,333,500,1000,1000,444,444,500,500,278,500,556,556,500,500,556,444,667,570,549,400,576,333",
    i: "250,333,420,500,500,833,778,214,333,333,500,675,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,675,675,675,500,920,611,611,667,722,611,611,722,722,333,444,667,556,833,667,722,611,722,611,500,556,722,611,833,611,556,556,389,278,389,422,500,333,500,500,444,500,444,278,500,500,278,278,444,278,722,500,500,500,500,389,389,278,500,444,667,444,444,389,400,275,400,541,556,556,333,333,500,889,889,444,444,500,500,278,500,500,500,500,500,500,444,611,675,549,400,576,250",
    bi: "250,389,555,500,500,833,778,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,832,667,667,667,722,667,667,722,778,389,500,667,611,889,722,722,611,722,667,556,611,722,667,889,667,611,611,333,278,333,570,500,333,500,500,444,500,444,333,500,500,278,278,500,278,778,556,500,500,500,389,389,278,556,444,667,500,444,389,348,220,348,570,500,500,333,333,500,1000,1000,444,444,500,500,278,500,556,556,500,500,556,444,667,570,549,400,576,250",
};

const WIDTH = Object.fromEntries(Object.entries(TABLE).map(([k, v]) => {
    const nums = v.split(",").map(Number);
    return [k, new Map([...CHARS].map((c, i) => [c, nums[i]]))];
})) as Record<keyof typeof TABLE, Map<string, number>>;

/** Potongan teks dengan gaya huruf. `halfPts`: ukuran huruf (setengah poin, seperti w:sz). */
export type TextPiece = { t: string; b: boolean; i: boolean; halfPts: number };

/** Lebar satu karakter (twip). Karakter di luar tabel dianggap selebar huruf "n" (lebih aman: sedikit lebar). */
function charWidth(ch: string, p: TextPiece): number {
    const table = WIDTH[p.b && p.i ? "bi" : p.b ? "b" : p.i ? "i" : "r"];
    const em = table.get(ch) ?? (/\p{Lu}/u.test(ch) ? 722 : 556);
    return (em / 1000) * (p.halfPts / 2) * 20;
}

/**
 * Jumlah baris paragraf bila dibungkus pada lebar `width` (twip), meniru Word: putus di spasi
 * (spasi di ujung baris tidak dihitung) dan sesudah tanda hubung; "\n" = ganti baris manual.
 * Ada cadangan 1% lebar untuk perbedaan perataan/kerning.
 */
export function measureLines(pieces: TextPiece[], width: number): number {
    const limit = width * 0.99;
    let lines = 1;
    let cur = 0;        // lebar baris tanpa spasi di ujung
    let pending = 0;    // spasi di belakang kata terakhir
    let word = 0;       // kata yang sedang disusun
    const place = () => {
        if (!word) return;
        if (cur > 0 && cur + pending + word > limit) { lines++; cur = 0; pending = 0; }
        if (cur === 0) {
            // Kata lebih panjang dari satu baris: dipotong per karakter oleh Word.
            while (word > limit) { lines++; word -= limit; }
            cur = word;
        } else cur += pending + word;
        pending = 0;
        word = 0;
    };
    for (const p of pieces) {
        for (const ch of p.t) {
            if (ch === "\n") { place(); lines++; cur = 0; pending = 0; continue; }
            if (ch === " " || ch === "\t") {
                place();
                if (cur > 0) pending += charWidth(" ", p);
                continue;
            }
            word += charWidth(ch, p);
            // Word boleh memutus sesudah tanda hubung.
            if (ch === "-" || ch === "–" || ch === "/") place();
        }
    }
    place();
    return lines;
}
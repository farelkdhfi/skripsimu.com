/**
 * Gambar yang lebih lebar dari area teks diperkecil secara proporsional.
 *
 * Mengganti margin (mis. 2,5 cm → kiri 4 cm) menyempitkan area teks, sehingga
 * gambar yang tadinya pas bisa meluber melewati margin kanan. Hanya ukuran
 * tampilan (`wp:extent` + `a:ext`) yang diubah; berkas gambar di
 * `word/media` dan rasio aspeknya tidak disentuh.
 */

const EMU_PER_TWIP = 635;

export type ImageReport = { scaled: number; floatingOverflow: number };

/** Perkecil drawing inline di satu paragraf agar lebarnya ≤ `maxTwips`. */
export function fitDrawings(xml: string, maxTwips: number, rep: ImageReport): string {
  const maxEmu = Math.floor(maxTwips * EMU_PER_TWIP);
  return xml.replace(/<wp:(inline|anchor)\b[\s\S]*?<\/wp:\1>/g, (block, kind: string) => {
    const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"\s*\/>/.exec(block);
    if (!ext) return block;
    const cx = Number(ext[1]);
    const cy = Number(ext[2]);
    if (cx <= maxEmu) return block;

    // Gambar melayang (anchor) punya posisi absolut; memperkecilnya tanpa
    // menghitung ulang posisi bisa menggeser tata letak. Cukup dilaporkan.
    if (kind === "anchor") { rep.floatingOverflow++; return block; }

    const k = maxEmu / cx;
    const ncx = Math.floor(cx * k);
    const ncy = Math.floor(cy * k);
    rep.scaled++;
    return block
      .replace(ext[0], `<wp:extent cx="${ncx}" cy="${ncy}"/>`)
      // Ukuran yang sama dicatat ulang di bentuk gambar (pic:spPr/a:xfrm).
      .replace(new RegExp(`<a:ext\\s+cx="${cx}"\\s+cy="${cy}"\\s*/>`, "g"), `<a:ext cx="${ncx}" cy="${ncy}"/>`);
  });
}
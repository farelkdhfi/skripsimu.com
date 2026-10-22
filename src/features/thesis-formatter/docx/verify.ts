/**
 * Pemeriksaan integritas: membandingkan dokumen asli dengan hasil format.
 *
 * Ini berjalan di produksi, bukan hanya saat pengujian. Kalau ada gambar yang
 * hilang, tabel yang berubah bentuk, atau teks yang bergeser, kita ingin
 * mengetahuinya sebelum file sampai ke pengguna.
 */

export type VerifyIssue = { level: "error" | "warn"; message: string };

export type DocFingerprint = {
  drawings: number;
  embeddedImageIds: string[];
  tables: number;
  rows: number;
  cells: number;
  paragraphsWithText: number;
  textRuns: number;
  text: string;
};

const count = (xml: string, re: RegExp) => (xml.match(re) || []).length;

/**
 * Teks dokumen, satu baris per paragraf.
 *
 * Spasi di awal/akhir paragraf dan spasi beruntun tidak dihitung: itu tata
 * letak (orang sering "memusatkan" baris dengan mengetik 20 spasi), bukan
 * isi. Formatter boleh membuangnya. Setiap kata, tanda baca, dan urutan
 * paragraf tetap harus identik.
 */

const LIST_MARK = /^(?:\(?(?:\d{1,2}|[a-zA-Z]{1,5})[.)]|[\u2022\u25CF\u25BA\u2713\u25AA\u2023\u2043*\-\u2013])\s+/;

const decode = (s: string) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Teks satu paragraf seperti yang dibandingkan verifikasi ("" bila kosong). */
export function paragraphLine(chunk: string): string {
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m: RegExpExecArray | null;
  let line = "";
  while ((m = re.exec(chunk))) line += m[1];
  return decode(line).replace(/\s+/g, " ").trim().replace(LIST_MARK, "");
}

function extractText(xml: string): string {
  return xml.split(/<\/w:p>/).map(paragraphLine).filter(Boolean).join("\n");
}

export function fingerprint(documentXml: string): DocFingerprint {
  return {
    drawings: count(documentXml, /<w:drawing[\s>]/g) + count(documentXml, /<w:pict[\s>]/g) + count(documentXml, /<w:object[\s>]/g),
    embeddedImageIds: [...documentXml.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]).sort(),
    tables: count(documentXml, /<w:tbl[\s>]/g),
    rows: count(documentXml, /<w:tr[\s>]/g),
    cells: count(documentXml, /<w:tc[\s>]/g),
    paragraphsWithText: count(documentXml, /<w:t(?:\s[^>]*)?>/g),
    textRuns: count(documentXml, /<w:r[\s>]/g),
    text: extractText(documentXml),
  };
}

/** Potongan teks pertama yang berbeda, untuk pesan galat yang bisa ditindak. */
function firstTextDivergence(a: string, b: string): string | null {
  if (a === b) return null;
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  const at = Math.max(0, i - 40);
  return `sekitar "...${a.slice(at, i + 40)}..." → "...${b.slice(at, i + 40)}..."`;
}

export type VerifyRegion = { beforeLines: string[]; afterLines: string[]; beforeText: string; afterText: string };

/**
 * @param opts.unorderedPrefix baris sampul yang boleh berubah urutan.
 * @param opts.regions blok yang disusun ulang (tanda tangan): kata-katanya harus
 *   sama persis, urutan & pemecahan barisnya boleh berubah.
 * @param opts.added tabel/baris/sel yang SENGAJA ditambahkan formatter.
 */
export function verify(
  before: DocFingerprint,
  after: DocFingerprint,
  opts: { unorderedPrefix?: number; regions?: VerifyRegion[]; added?: { tables: number; rows: number; cells: number } } = {},
): VerifyIssue[] {
  const issues: VerifyIssue[] = [];
  const add = opts.added ?? { tables: 0, rows: 0, cells: 0 };

  if (after.drawings < before.drawings)
    issues.push({ level: "error", message: `${before.drawings - after.drawings} gambar/objek hilang dari dokumen hasil.` });

  const lost = before.embeddedImageIds.filter((id) => !after.embeddedImageIds.includes(id));
  if (lost.length)
    issues.push({ level: "error", message: `${lost.length} referensi gambar tertanam tidak ditemukan lagi di dokumen hasil.` });

  if (after.tables - add.tables !== before.tables)
    issues.push({ level: "error", message: `Jumlah tabel berubah: ${before.tables} → ${after.tables - add.tables}.` });
  if (after.rows - add.rows !== before.rows)
    issues.push({ level: "error", message: `Jumlah baris tabel berubah: ${before.rows} → ${after.rows - add.rows}.` });
  if (after.cells - add.cells !== before.cells)
    issues.push({ level: "error", message: `Jumlah sel tabel berubah: ${before.cells} → ${after.cells - add.cells}.` });

  // Keluarkan baris blok yang disusun ulang, sisanya wajib identik & berurutan.
  const removeLines = (lines: string[], targets: string[]) => {
    const out = [...lines];
    let cursor = 0;
    for (const t of targets) {
      let k = out.indexOf(t, cursor);
      if (k === -1) k = out.indexOf(t);
      if (k !== -1) { out.splice(k, 1); cursor = k; }
    }
    return out;
  };
  const regions = opts.regions ?? [];
  let a = before.text.split("\n");
  let b = after.text.split("\n");
  for (const r of regions) { a = removeLines(a, r.beforeLines); b = removeLines(b, r.afterLines); }

  const k = opts.unorderedPrefix ?? 0;
  const norm = (lines: string[]) => {
    const low = lines.map((l) => l.toLowerCase());
    return [...low.slice(0, k).sort(), ...low.slice(k)].join("\n");
  };
  const na = norm(a);
  const nb = norm(b);
  if (na !== nb) {
    const where = firstTextDivergence(na, nb);
    issues.push({ level: "error", message: `Teks dokumen berubah${where ? ` ${where}` : ""}.` });
  }

  // Isi blok tanda tangan: kumpulan kata harus sama persis.
  const bag = (s: string) => s.toLowerCase().split(/\s+/).filter(Boolean).sort().join(" ");
  const ba = bag(regions.map((r) => r.beforeText).join(" "));
  const bb = bag(regions.map((r) => r.afterText).join(" "));
  if (ba !== bb) issues.push({ level: "error", message: "Teks dokumen berubah di blok tanda tangan." });

  return issues;
}
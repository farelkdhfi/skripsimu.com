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
  /** Rumus Office Math (m:oMath), tidak termasuk salinan cadangan di mc:Fallback. */
  equations: number;
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
 * Penanda daftar ketikan yang boleh hilang (diganti numbering otomatis Word):
 * angka ("1.", "12)"), SATU huruf ("a.", "B)"), angka Romawi ("iv.", "II."), bullet.
 * Singkatan/gelar ("Prof.", "Dr.", "NIDN.", "NIP.") BUKAN penanda dan tidak dibuang.
 */
const LIST_MARK = /^(?:\(?(?:\d{1,3}|[a-zA-Z]|[ivxlcdm]{1,5}|[IVXLCDM]{1,5})[.)]|[\u2022\u25CF\u25BA\u2713\u25AA\u2023\u2043*\-\u2013])\s+/;

const decode = (s: string) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/**
 * Teks satu paragraf seperti yang dibandingkan verifikasi ("" bila kosong).
 *
 * Spasi di awal/akhir paragraf dan spasi beruntun tidak dihitung: itu tata
 * letak (orang sering "memusatkan" baris dengan mengetik 20 spasi), bukan
 * isi. Formatter boleh membuangnya. Setiap kata, tanda baca, dan urutan
 * paragraf tetap harus identik.
 */
export function paragraphLine(chunk: string): string {
  // Teks biasa (w:t) DAN teks rumus (m:t): rumus yang hilang/berubah ikut terdeteksi.
  const re = /<[wm]:t(?:\s[^>]*)?>([\s\S]*?)<\/[wm]:t>/g;
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
    equations: count(documentXml.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/g, ""), /<m:oMath[\s>]/g),
    embeddedImageIds: [...documentXml.matchAll(/<a:blip[^>]*r:embed="([^"]+)"/g)].map((m) => m[1]).sort(),
    tables: count(documentXml, /<w:tbl[\s>]/g),
    rows: count(documentXml, /<w:tr[\s>]/g),
    cells: count(documentXml, /<w:tc[\s>]/g),
    paragraphsWithText: count(documentXml, /<w:t(?:\s[^>]*)?>/g),
    textRuns: count(documentXml, /<w:r[\s>]/g),
    text: extractText(documentXml),
  };
}

export type VerifyRegion = { beforeLines: string[]; afterLines: string[]; beforeText: string; afterText: string };

/**
 * @param unorderedPrefix baris sampul yang boleh berubah urutan.
 * @param regions blok yang disusun ulang modul (tanda tangan, keterangan, judul, daftar):
 *   barisnya dikeluarkan dari perbandingan urut; kata-katanya dibandingkan sebagai kumpulan.
 * @param added tabel/baris/sel/gambar yang SENGAJA ditambahkan (positif) atau dibuang (negatif) formatter.
 * @param removedImageIds gambar tertanam yang SENGAJA tidak dipakai lagi (mis. bingkai/hiasan sampul).
 */
export type VerifyOptions = {
  unorderedPrefix?: number;
  regions?: VerifyRegion[];
  added?: { tables: number; rows: number; cells: number; drawings?: number; equations?: number };
  removedImageIds?: string[];
};

/** Kunci pencocokan baris: huruf besar-kecil & spasi tidak dihitung (judul bisa dikapitalkan tahap berikutnya). */
const keyOf = (s: string) => s.toLowerCase().replace(/\s+/g, "");

/** Kemunculan `want` yang tidak terpakai dan paling dekat dengan `from`. */
function nearest(keys: string[], want: string, from: number, gone: Set<number>): number {
  for (let d = 0; d < keys.length; d++) {
    const f = from + d;
    const b = from - d;
    if (f < keys.length && !gone.has(f) && keys[f] === want) return f;
    if (d > 0 && b >= 0 && !gone.has(b) && keys[b] === want) return b;
  }
  return -1;
}

/**
 * Keluarkan baris satu blok. Teks yang sama bisa muncul lebih dari sekali ("KATA PENGANTAR"
 * di Daftar Isi dan sebagai judul), jadi blok ditempatkan di posisi yang paling banyak
 * memuat barisnya secara berurutan, lalu tiap baris diambil dari kemunculan terdekat.
 */
function removeGroup(lines: string[], targets: string[]): string[] {
  const gone = removeGroupIdx(lines, targets);
  return gone.size ? lines.filter((_, i) => !gone.has(i)) : lines;
}

/** Seperti removeGroup, tetapi mengembalikan indeks baris yang dikeluarkan. */
export function removeGroupIdx(lines: string[], targets: string[]): Set<number> {
  if (!targets.length) return new Set();
  const keys = lines.map(keyOf);
  const want = targets.map(keyOf);
  const gone = new Set<number>();

  let start = -1;
  let best = -1;
  for (let s = 0; s < keys.length; s++) {
    if (keys[s] !== want[0]) continue;
    let score = 1;
    let j = s + 1;
    for (let t = 1; t < want.length && j < keys.length; t++) {
      const k = keys.indexOf(want[t], j);
      if (k !== -1 && k - j <= 40) { score++; j = k + 1; }
    }
    if (score > best) { best = score; start = s; }
  }

  // Baris blok dicari MAJU dari baris blok sebelumnya (blok = baris berurutan). Teks yang sama bisa
  // ada di tempat lain (entri Daftar Isi "1.1 Latar Belakang" = judul subbabnya); kalau kemunculan
  // di dalam blok sudah diambil blok lain, kemunculan terdekat di luar dipakai TANPA memindahkan
  // posisi blok — supaya baris berikutnya tetap diambil dari blok ini, bukan dari isi bab.
  const WINDOW = 8;
  let cursor = start === -1 ? 0 : start;
  for (const w of want) {
    let k = -1;
    for (let f = cursor; f < keys.length && f <= cursor + WINDOW; f++) {
      if (!gone.has(f) && keys[f] === w) { k = f; break; }
    }
    if (k !== -1) { gone.add(k); cursor = k; continue; }
    const far = nearest(keys, w, cursor, gone);
    if (far !== -1) gone.add(far);
  }
  return gone;
}

/**
 * Baris yang tersisa setelah blok yang dilaporkan modul dikeluarkan — inilah yang wajib
 * identik & berurutan. Dipakai juga untuk menampilkan baris yang berubah TANPA dilaporkan.
 */
export function residualLines(before: DocFingerprint, after: DocFingerprint, opts: VerifyOptions = {}): { a: string[]; b: string[] } {
  let a = before.text.split("\n");
  let b = after.text.split("\n");
  // Blok besar dulu: blok panjang (mis. Daftar Isi) paling pasti posisinya. Blok satu baris (judul
  // bab yang diseragamkan) yang diproses lebih dulu bisa "mengambil" baris kembarnya di Daftar Isi,
  // lalu blok Daftar Isi terpaksa mengambil baris dari isi bab.
  // Blok yang hanya mengubah spasi/tab (keterangan "R4 = 30%" → "R4⇥=⇥30%") tidak dikeluarkan:
  // teks yang sama sering muncul berkali-kali (keterangan tabel yang diulang), dan mengeluarkan
  // kemunculan yang salah merusak urutan. Baris itu dibandingkan tanpa spasi di kedua sisi.
  const wsOnly = (r: VerifyRegion) => r.beforeLines.length === r.afterLines.length && r.beforeLines.every((l, i) => keyOf(l) === keyOf(r.afterLines[i]));
  const all = opts.regions ?? [];
  const wsKeys = new Set(all.filter(wsOnly).flatMap((r) => r.beforeLines.map(keyOf)));
  const regions = all.filter((r) => !wsOnly(r));
  for (const r of [...regions].sort((x, y) => y.beforeLines.length - x.beforeLines.length)) a = removeGroup(a, r.beforeLines);
  for (const r of [...regions].sort((x, y) => y.afterLines.length - x.afterLines.length)) b = removeGroup(b, r.afterLines);
  const k = opts.unorderedPrefix ?? 0;
  const prep = (lines: string[]) => {
    const low = lines.map((l) => (wsKeys.has(keyOf(l)) ? keyOf(l) : l.toLowerCase().replace(/\s+/g, " ").trim())).filter(Boolean);
    return [...low.slice(0, k).sort(), ...low.slice(k)];
  };
  return { a: prep(a), b: prep(b) };
}

/** Pasangan baris pertama yang berbeda, untuk pesan yang bisa ditindaklanjuti. */
function firstLineDivergence(a: string[], b: string[]): string {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  const cut = (s: string | undefined) => (s === undefined ? "(akhir dokumen)" : `"${s.length > 90 ? s.slice(0, 90) + "…" : s}"`);
  return `sekitar baris ${cut(a[i])} → ${cut(b[i])}`;
}

/** Kunci kata untuk perbandingan kumpulan: huruf & angka saja (koma/titik yang berpindah tidak dihitung). */
const wordKey = (w: string) => w.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");

/** Jumlah kemunculan tiap baris (kunci tanpa spasi & huruf besar-kecil). */
function lineCounts(text: string): Map<string, number> {
  const m = new Map<string, number>();
  for (const l of text.split("\n")) {
    const k = keyOf(l);
    if (k) m.set(k, (m.get(k) ?? 0) + 1);
  }
  return m;
}

/** Baris yang posisinya berbeda antara a dan b (bagian yang tidak masuk urutan bersama terpanjang). */
function movedLines(a: string[], b: string[]): string[] {
  const n = a.length;
  const m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const out: string[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) out.push(a[i++]);
    else out.push(b[j++]);
  }
  while (i < n) out.push(a[i++]);
  while (j < m) out.push(b[j++]);
  return out;
}

export function verify(before: DocFingerprint, after: DocFingerprint, opts: VerifyOptions = {}): VerifyIssue[] {
  const issues: VerifyIssue[] = [];
  const add = opts.added ?? { tables: 0, rows: 0, cells: 0 };

  const expectedDrawings = before.drawings + (add.drawings ?? 0);
  if (after.drawings < expectedDrawings)
    issues.push({ level: "error", message: `${expectedDrawings - after.drawings} gambar/objek hilang dari dokumen hasil.` });

  const expectedEq = before.equations + (add.equations ?? 0);
  if (after.equations < expectedEq)
    issues.push({ level: "error", message: `${expectedEq - after.equations} rumus hilang dari dokumen hasil (asli ${before.equations}, hasil ${after.equations}).` });

  const allowed = [...(opts.removedImageIds ?? [])];
  const lost = before.embeddedImageIds.filter((id) => {
    if (after.embeddedImageIds.includes(id)) return false;
    const k = allowed.indexOf(id);
    if (k === -1) return true;
    allowed.splice(k, 1);
    return false;
  });
  if (lost.length)
    issues.push({ level: "error", message: `${lost.length} referensi gambar tertanam tidak ditemukan lagi di dokumen hasil.` });

  if (after.tables - add.tables !== before.tables)
    issues.push({ level: "error", message: `Jumlah tabel berubah: ${before.tables} → ${after.tables - add.tables}.` });
  if (after.rows - add.rows !== before.rows)
    issues.push({ level: "error", message: `Jumlah baris tabel berubah: ${before.rows} → ${after.rows - add.rows}.` });
  if (after.cells - add.cells !== before.cells)
    issues.push({ level: "error", message: `Jumlah sel tabel berubah: ${before.cells} → ${after.cells - add.cells}.` });

  const { a, b } = residualLines(before, after, opts);
  if (a.join("\n") !== b.join("\n")) {
    // Isi sama, hanya urutan berbeda? Baris yang teksnya muncul lebih dari sekali (judul & entri
    // Daftar Isi, "Sumber: …", "Keterangan:") tidak bisa dibedakan lewat teks: itu ambiguitas
    // pencocokan, bukan kerusakan. Paragraf unik yang berpindah tetap dilaporkan.
    const tally = new Map<string, number>();
    for (const l of a) tally.set(l, (tally.get(l) ?? 0) + 1);
    for (const l of b) tally.set(l, (tally.get(l) ?? 0) - 1);
    const sameContent = [...tally.values()].every((v) => v === 0);
    const cBefore = lineCounts(before.text);
    const cAfter = lineCounts(after.text);
    const ambiguous = (l: string) => Math.max(cBefore.get(keyOf(l)) ?? 0, cAfter.get(keyOf(l)) ?? 0) >= 2;
    const moved = sameContent ? movedLines(a, b) : [];
    if (!sameContent) {
      issues.push({ level: "error", message: `Teks dokumen berubah ${firstLineDivergence(a, b)}.` });
    } else if (moved.some((l) => !ambiguous(l))) {
      const l = moved.find((x) => !ambiguous(x))!;
      issues.push({ level: "error", message: `Teks dokumen berubah sekitar baris "${l.slice(0, 90)}": urutan paragraf berpindah.` });
    }
  }

  // Blok yang disusun ulang: kumpulan kata harus sama persis; kata yang berbeda disebutkan.
  const tally = new Map<string, number>();
  const words = (s: string) => s.split(/\s+/).map(wordKey).filter(Boolean);
  for (const r of opts.regions ?? []) {
    for (const w of words(r.beforeText)) tally.set(w, (tally.get(w) ?? 0) + 1);
    for (const w of words(r.afterText)) tally.set(w, (tally.get(w) ?? 0) - 1);
  }
  const missing = [...tally].filter(([, n]) => n > 0).map(([w]) => w);
  const extra = [...tally].filter(([, n]) => n < 0).map(([w]) => w);
  if (missing.length || extra.length) {
    const list = (xs: string[]) => xs.slice(0, 8).map((w) => `"${w}"`).join(", ") + (xs.length > 8 ? ", …" : "");
    issues.push({
      level: "error",
      message: `Teks dokumen berubah di bagian yang disusun ulang (tanda tangan/keterangan/judul)` +
        `${missing.length ? `: hilang ${list(missing)}` : ""}${extra.length ? `${missing.length ? ";" : ":"} bertambah ${list(extra)}` : ""}.`,
    });
  }

  return issues;
}
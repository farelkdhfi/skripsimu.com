/**
 * LAPISAN STRUKTUR.
 *
 * Modul ini hanya menjawab satu pertanyaan: "bagian apa ini?" — dan tidak tahu
 * apa pun soal font, margin, atau spasi. Pemisahan ini yang nanti membuat
 * template pedoman per universitas cukup mengganti lapisan `rules/` tanpa
 * menyentuh deteksi struktur sama sekali.
 */

export type Zone = "cover" | "front" | "main" | "back";

export type Role =
  | "cover"
  | "front_heading"      // MOTTO, PERSEMBAHAN, LEMBAR PENGESAHAN, KATA PENGANTAR, ABSTRAK, ...
  | "list_heading"       // DAFTAR ISI / DAFTAR TABEL / DAFTAR GAMBAR / DAFTAR LAMPIRAN
  | "list_entry"         // baris di dalam daftar tersebut (dot leader + nomor halaman)
  | "chapter_label"      // "BAB I"
  | "chapter"            // "PENDAHULUAN" atau "BAB I PENDAHULUAN" satu baris
  | "section"            // 1.1
  | "subsection"         // 1.1.1
  | "figure"             // paragraf berisi gambar
  | "figure_caption"
  | "table_caption"
  | "body"
  | "reference_heading"  // DAFTAR PUSTAKA
  | "reference"          // entri di dalamnya
  | "appendix_heading"
  | "empty";

/**
 * Judul bagian awal dicocokkan dalam dua tingkat ketegasan:
 *
 * - Frasa majemuk yang khas ("LEMBAR PERSETUJUAN …", "PERNYATAAN KEASLIAN …",
 *   "KATA PENGANTAR") boleh diikuti hingga 3 kata tambahan, karena kampus
 *   menulisnya bervariasi: "LEMBAR PERSETUJUAN PEMBIMBING",
 *   "PERNYATAAN KEASLIAN KARYA ILMIAH".
 * - Kata tunggal yang juga bisa muncul di judul skripsi ("JUDUL", "SAMPUL",
 *   "RINGKASAN") harus berdiri sendiri. Kalau tidak, judul sampul seperti
 *   "JUDUL SKRIPSI TENTANG ..." akan dikira halaman judul dan dipaksa pindah
 *   halaman.
 */
const TAIL = "(?:\\s+[^\\s]+){0,3}";
const SUFFIX = "(?:\\s+(?:skripsi|tesis|disertasi|proposal|penelitian|penulis))?";

const FRONT_COMPOUND =
  "lembar\\s+(?:pengesahan|persetujuan|pernyataan)|" +
  "halaman\\s+(?:pengesahan|persetujuan|pernyataan|judul|sampul|persembahan|motto)|" +
  "surat\\s+pernyataan|pernyataan\\s+(?:keaslian|orisinalitas|keasliaan|bebas\\s+plagiat)|" +
  "kata\\s+pengantar|ucapan\\s+terima\\s+kasih|riwayat\\s+hidup|" +
  "motto\\s+dan\\s+persembahan|moto\\s+dan\\s+persembahan|" +
  "daftar\\s+(?:singkatan|istilah|simbol|lambang|notasi|arti)";

const FRONT_SINGLE =
  "motto|moto|persembahan|pengesahan|persetujuan|abstrak|abstract|intisari|ringkasan|biodata";

export const FRONT_HEADING_RE = new RegExp(
  `^(?:(?:${FRONT_COMPOUND})${TAIL}|(?:${FRONT_SINGLE})${SUFFIX})$`, "i",
);

/**
 * Jenis halaman bagian awal. Ini menentukan bagaimana paragraf DI BAWAH
 * judulnya diformat, karena ketiganya punya tata letak yang sangat berbeda:
 *
 * - form     : persetujuan, pengesahan, pernyataan, motto, persembahan.
 *              Tata letaknya disengaja (tanggal rata kanan, kolom tanda
 *              tangan dengan tab, nama di tengah) dan harus dipertahankan.
 * - prose    : kata pengantar, ucapan terima kasih — paragraf biasa.
 * - abstract : abstrak/abstract — paragraf biasa dengan spasi tunggal.
 */
export type FrontKind = "form" | "prose" | "abstract";

/** Penanda butir daftar ketikan (bukan daftar otomatis Word). */
export const LIST_MARKER_RE = /^(?:\(?(?:\d{1,2}|[a-zA-Z]|[ivxIVX]{1,4})[.)]|[\u2022\u25CF\u25BA\u2713\u25AA\u2023\u2043\-\u2013*])[\s\t]+\S/;

export function frontKindOf(heading: string): FrontKind {
  const t = heading.toLowerCase();
  if (/abstra|intisari|ringkasan/.test(t)) return "abstract";
  if (/kata\s+pengantar|ucapan\s+terima\s+kasih/.test(t)) return "prose";
  return "form";
}

export const LIST_HEADING_RE = new RegExp(
  "^daftar\\s+(?:isi|tabel|gambar|lampiran|grafik|bagan)" + SUFFIX + "$", "i",
);

export const REFERENCE_HEADING_RE = new RegExp(
  "^(?:daftar\\s+(?:pustaka|referensi|rujukan)|referensi|bibliografi|bibliography|references)" + SUFFIX + "$", "i",
);
/** "Lampiran : 1 (satu) berkas" (baris surat) bukan judul lampiran → titik dua langsung setelah kata ditolak. */
export const APPENDIX_HEADING_RE = /^(lampiran|appendix|appendices)\b(?!\s*[:：])/i;
export const CHAPTER_RE = /^(?:bab|chapter)\s+([ivxlcdm]+|\d+)\b/i;

/** "BAB I" berdiri sendiri — judulnya ada di paragraf berikutnya. */
export const CHAPTER_LABEL_ONLY_RE = /^(?:bab|chapter)\s+([ivxlcdm]+|\d+)\s*[.:]?\s*$/i;

const LIST_ENTRY_RE = /(\.{3,}|\u2026|\t)\s*[\divxlcdm]+\s*$/i;
/**
 * Di dalam blok daftar isi, akhiran nomor halaman sudah cukup jadi penanda.
 * Pemisahnya harus tegas (titik-titik, tab, atau dua spasi) — satu spasi saja
 * akan membuat "BAB I" terbaca sebagai entri yang berakhir angka romawi.
 */
const PAGE_NUM_TAIL = /(\.{2,}|\u2026|\t|\s{2,})\s*[\divxlcdm]+\s*$/i;
const SECTION_RE = /^(\d+)\.(\d+)[.)]?\s+\S/;
const SUBSECTION_RE = /^(\d+)\.(\d+)\.(\d+)[.)]?\s+\S/;
const ALPHA_SECTION_RE = /^[A-Z]\.\s+\S/;
const CAPTION_RE = /^(gambar|gbr\.?|figure|fig\.?|grafik|bagan|diagram|tabel|table|tbl\.?)\s*\.?\s*(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*\s*[.:]?\s*/i;

/**
 * Keterangan gambar/tabel memakai Title Case setelah nomornya
 * ("Tabel 3.1 Variabel Penelitian"). Kalimat isi yang kebetulan diawali
 * rujukan ("Tabel 3.1 menunjukkan bahwa ...") dilanjutkan huruf kecil, dan
 * itulah pembedanya. Pemeriksaan dilakukan di luar regex supaya backtracking
 * tidak bisa meloloskan bentuk kedua.
 */
/** Keterangan tanpa nomor: "Tabel Sebaran Populasi …" (judul pendek, tidak diakhiri titik). */
const UNNUMBERED_CAPTION_RE = /^(tabel|table|gambar|figure)\s+[A-Z][^.]{3,140}$/;

export function captionKind(t: string): Role | null {
  const m = CAPTION_RE.exec(t);
  if (m) {
    const rest = t.slice(m[0].length);
    if (/^[a-z]/.test(rest)) return null;
    return /^(tabel|table|tbl\.?)$/i.test(m[1]) ? "table_caption" : "figure_caption";
  }
  const u = UNNUMBERED_CAPTION_RE.exec(t);
  if (u) return /^(tabel|table)$/i.test(u[1]) ? "table_caption" : "figure_caption";
  return null;
}

const headingish = (t: string) => !/[.!?;,]$/.test(t) && t.split(/\s+/).length <= 16;

export type Signals = {
  text: string;
  hasDrawing: boolean;
  isTable: boolean;
  /** Paragraf ber-style keterangan (Caption, CaptionGambar, …): keterangan panjang tetap dikenali. */
  captionStyled?: boolean;
  /** Paragraf judul bab bernomor romawi saja ("I" ↵ "PENDAHULUAN"), dari unifyChapterHeadings. */
  chapterHint?: boolean;
};

/**
 * Petunjuk tata letak asli sebuah paragraf. Ini bukan isi dokumen, melainkan
 * jejak keputusan penulis (rata kanan, rata tengah, kolom tab) yang dipakai
 * lapisan aturan untuk membedakan "berantakan" dari "disengaja".
 */
export type LayoutHint = {
  align: "left" | "center" | "right" | "both" | null;
  tabCount: number;
  /** Semua tab berada di awal paragraf — indentasi manual, bukan kolom. */
  leadingTabsOnly: boolean;
  /** Butir daftar yang diketik manual: "1. …", "a) …", "• …". */
  listMarker: boolean;
  /** Seluruh teks paragraf dicetak tebal. */
  allBold: boolean;
  length: number;
  frontKind: FrontKind | null;

  /** Bagian dari judul bab dua baris ("BAB I" + "PENDAHULUAN"). */
  pairedChapter?: boolean;
};

/**
 * Klasifikasi berbasis aturan untuk satu node. Mengembalikan `null` bila
 * tidak yakin — di situlah AI mengambil alih.
 */
export function classifyByRule(s: Signals, zone: Zone, inList = false): Role | null {
  if (s.isTable) return null;
  const t = s.text.trim();

  if (s.hasDrawing) return "figure";
  if (!t) return "empty";
  if (s.chapterHint) return "chapter";

  const bare = t.replace(/[:\s]+$/, "");

  // "BAB I" yang berdiri sendiri selalu label bab, tidak pernah entri daftar.
  if (t.length <= 120 && CHAPTER_LABEL_ONLY_RE.test(bare)) return "chapter_label";

  if (LIST_ENTRY_RE.test(t)) return "list_entry";
  // Di dalam halaman DAFTAR ISI/TABEL/GAMBAR, baris berakhiran nomor halaman
  // adalah entri daftar — bukan judul bab, sekalipun bunyinya sama persis.
  if (inList && PAGE_NUM_TAIL.test(t)) return "list_entry";
  if (t.length <= (s.captionStyled ? 700 : 200)) {
    const caption = captionKind(t);
    if (caption) return caption;
  }
  if (t.length <= 120) {
    if (CHAPTER_RE.test(bare)) return "chapter";
    if (LIST_HEADING_RE.test(bare) && headingish(bare)) return "list_heading";
    if (REFERENCE_HEADING_RE.test(bare) && headingish(bare)) return "reference_heading";
    if (APPENDIX_HEADING_RE.test(bare) && headingish(bare) && bare.length <= 60) return "appendix_heading";
    if (FRONT_HEADING_RE.test(bare) && headingish(bare)) return "front_heading";
    if (headingish(t)) {
      if (SUBSECTION_RE.test(t)) return "subsection";
      if (SECTION_RE.test(t)) return "section";
      if (ALPHA_SECTION_RE.test(t) && zone === "main") return "section";
    }
  }
  if (zone === "cover") return "cover";
  return null;
}

/**
 * Tentukan zona setiap node berdasarkan penanda struktural dokumen asli.
 * Zona inilah yang dipakai untuk penomoran halaman romawi vs arab — dan ia
 * mengikuti dokumen, bukan template.
 */
export function scanStructure(signals: Signals[]): { zones: Zone[]; inList: boolean[] } {
  const zones: Zone[] = new Array(signals.length).fill("cover");
  const inList: boolean[] = new Array(signals.length).fill(false);
  let zone: Zone = "cover";
  let listing = false;

  for (let i = 0; i < signals.length; i++) {
    const s = signals[i];
    if (s.isTable) { zones[i] = zone; inList[i] = listing; continue; }

    const t = s.text.trim().replace(/[:\s]+$/, "");
    const raw = s.text.trim();
    const entryLike = !CHAPTER_LABEL_ONLY_RE.test(t) &&
      (LIST_ENTRY_RE.test(raw) || (listing && PAGE_NUM_TAIL.test(raw)));

    // Blok daftar berakhir pada baris pertama yang bukan entri.
    if (listing && t && !entryLike) listing = false;

    // Perpindahan zona hanya dipicu judul sungguhan, bukan barisan daftar isi.
    if (!listing && !entryLike && t.length <= 120) {
      if (s.chapterHint || CHAPTER_RE.test(t)) zone = "main";
      else if (REFERENCE_HEADING_RE.test(t) && headingish(t)) zone = "back";
      else if (APPENDIX_HEADING_RE.test(t) && headingish(t) && t.length <= 60) zone = "back";
      else if (zone === "cover" && (FRONT_HEADING_RE.test(t) || LIST_HEADING_RE.test(t)) && headingish(t)) zone = "front";
    }

    zones[i] = zone;
    inList[i] = listing;

    if (!listing && t.length <= 60 && LIST_HEADING_RE.test(t) && headingish(t)) listing = true;
  }
  return { zones, inList };
}

/** Judul-judul yang memulai halaman baru pada skripsi Indonesia pada umumnya. */
export function startsNewPage(role: Role): boolean {
  return (
    role === "chapter_label" ||
    role === "chapter" ||
    role === "front_heading" ||
    role === "list_heading" ||
    role === "reference_heading" ||
    role === "appendix_heading"
  );
}

export const isHeadingRole = (r: Role) =>
  r === "chapter" || r === "chapter_label" || r === "section" || r === "subsection" ||
  r === "front_heading" || r === "list_heading" || r === "reference_heading" ||
  r === "appendix_heading";
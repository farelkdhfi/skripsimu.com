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
  "^daftar\\s+(?:isi|tabel|gambar|ilustrasi|lampiran|grafik|bagan)" + SUFFIX + "$", "i",
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
/** Subbab = TEPAT tiga tingkat (2.1.4). Empat tingkat (2.1.4.1) = judul tingkat 4 (minor-headings), bukan subbab. */
const SUBSECTION_RE = /^(\d+)\.(\d+)\.(\d+)[.)]?\s+\S/;
const ALPHA_SECTION_RE = /^[A-Z]\.\s+\S/;
const CAPTION_RE = /^(gambar|gbr\.?|figure|fig\.?|grafik|bagan|diagram|ilustrasi|tabel|table|tbl\.?)\s*\.?\s*(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*\s*[.:]?\s*/i;

/**
 * Keterangan gambar/tabel memakai Title Case setelah nomornya
 * ("Tabel 3.1 Variabel Penelitian"). Kalimat isi yang kebetulan diawali
 * rujukan ("Tabel 3.1 menunjukkan bahwa ...") dilanjutkan huruf kecil, dan
 * itulah pembedanya. Pemeriksaan dilakukan di luar regex supaya backtracking
 * tidak bisa meloloskan bentuk kedua.
 */
/** Keterangan tanpa nomor: "Tabel Sebaran Populasi …" (judul pendek, tidak diakhiri titik). */
const UNNUMBERED_CAPTION_RE = /^(tabel|table|gambar|figure|ilustrasi)\s+[A-Z][^.]{3,140}$/;

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
  /** Paragraf hanya memuat gambar MELAYANG (tanda tangan, stempel, foto biodata). */
  floatingOnly?: boolean;
  /** Paragraf ini menutup halaman (section break / pindah halaman manual). */
  pageEnd?: boolean;
  /** Paragraf ini membuka halaman baru (page break before). */
  pageStart?: boolean;
};

/**
 * Penanda halaman pengesahan/persetujuan. Halaman seperti ini sering TANPA judul (hanya judul
 * skripsi, "OLEH :", lalu blok tanda tangan), sehingga tidak boleh ikut dianggap sampul.
 */
const APPROVAL_RE = /^(?:menyetujui|mengesahkan|disetujui|disahkan|mengetahui|pembimbing\b|ketua\s+(?:panitia|program|jurusan|departemen|prodi|sidang|komisi)|dekan\b|wakil\s+dekan)/i;

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

  // Gambar melayang di bagian awal/akhir (tanda tangan, stempel, foto) bukan gambar bernomor:
  // paragrafnya diperlakukan menurut teksnya. Begitu pula paragraf panjang yang membungkus foto.
  const floatingText = Boolean(s.floatingOnly) && (zone !== "main" || t.length >= 120);
  if (s.hasDrawing && !floatingText) return "figure";
  if (!t) return "empty";
  // Di wilayah halaman DAFTAR ISI/TABEL/GAMBAR/LAMPIRAN semua baris berisi adalah entri daftar —
  // termasuk kepala kolom ("Bab … Halaman"), baris bab tanpa nomor halaman, dan baris lanjutan.
  if (inList) return "list_entry";
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
  /** Awal halaman yang sedang dibaca & jumlah batas halaman yang sudah dilewati. */
  let pageFrom = 0;
  let pages = 0;
  /** Setelah baris tahun ("2026") — penutup halaman sampul — bila halaman tidak dipisah break. */
  let afterYear = 0;

  const bareOf = (s: Signals) => s.text.trim().replace(/[:\s]+$/, "");
  const entryAt = (k: number) => {
    const s = signals[k];
    const raw = s.text.trim();
    return !s.isTable && !CHAPTER_LABEL_ONLY_RE.test(bareOf(s)) && (LIST_ENTRY_RE.test(raw) || PAGE_NUM_TAIL.test(raw));
  };
  /** Dua paragraf berisi berikutnya. */
  const ahead = (i: number) => {
    const out: number[] = [];
    for (let k = i + 1; k < signals.length && out.length < 2; k++) {
      if (signals[k].isTable) break;
      if (signals[k].text.trim()) out.push(k);
    }
    return out;
  };

  for (let i = 0; i < signals.length; i++) {
    const s = signals[i];
    if (s.pageStart && i > 0 && i !== pageFrom) { pageFrom = i; pages++; }
    if (s.isTable) { zones[i] = zone; inList[i] = listing; continue; }

    const t = bareOf(s);
    const raw = s.text.trim();
    const entryLike = !CHAPTER_LABEL_ONLY_RE.test(t) &&
      (LIST_ENTRY_RE.test(raw) || (listing && PAGE_NUM_TAIL.test(raw)));

    // Wilayah halaman daftar berakhir pada JUDUL SUNGGUHAN, bukan pada baris pertama tanpa
    // nomor halaman: kepala kolom ("Bab … Halaman"), baris bab ("I  PENDAHULUAN"), dan baris
    // lanjutan entri panjang tetap bagian daftar. Baris bab dianggap judul sungguhan bila
    // yang menyusul bukan entri daftar (di daftar isi, baris bab selalu diikuti entri bernomor halaman).
    if (listing && t && !entryLike) {
      const named = t.length <= 120 && headingish(t) && (
        LIST_HEADING_RE.test(t) || FRONT_HEADING_RE.test(t) || REFERENCE_HEADING_RE.test(t) ||
        (APPENDIX_HEADING_RE.test(t) && t.length <= 60));
      const chapterish = t.length <= 120 && (Boolean(s.chapterHint) || CHAPTER_RE.test(t));
      const prose = raw.length >= 150 && /[.!?]$/.test(raw) && !raw.includes("\t");
      if (named || prose || (chapterish && !ahead(i).some(entryAt))) listing = false;
    }

    // Perpindahan zona hanya dipicu judul sungguhan, bukan barisan daftar isi.
    if (!listing && !entryLike && t.length <= 120) {
      if (s.chapterHint || CHAPTER_RE.test(t)) zone = "main";
      else if (REFERENCE_HEADING_RE.test(t) && headingish(t)) zone = "back";
      else if (APPENDIX_HEADING_RE.test(t) && headingish(t) && t.length <= 60) zone = "back";
      else if (zone === "cover" && (FRONT_HEADING_RE.test(t) || LIST_HEADING_RE.test(t)) && headingish(t)) zone = "front";
    }

    // Halaman setelah sampul tanpa judul (mis. pengesahan: judul skripsi, "OLEH :", tanda tangan)
    // dikenali dari isinya: begitu muncul penanda pengesahan, seluruh halaman itu bagian awal.
    // Halaman sampul pertama tidak pernah terkena aturan ini.
    // Banyak penulis berpindah halaman cukup dengan baris kosong (tanpa break): awal halamannya
    // lalu diambil dari baris tahun penutup sampul ("Jatinangor ↵ 2026") yang terakhir.
    if (zone === "cover" && !listing && (pages > 0 || afterYear > 0) && raw.length <= 80 && APPROVAL_RE.test(raw)) {
      for (let k = Math.max(pageFrom, afterYear); k < i; k++) zones[k] = "front";
      zone = "front";
    }
    if (zone === "cover" && /^(?:19|20)\d{2}$/.test(raw)) afterYear = i + 1;

    zones[i] = zone;
    inList[i] = listing;

    if (!listing && t.length <= 60 && LIST_HEADING_RE.test(t) && headingish(t)) listing = true;
    if (s.pageEnd) { pageFrom = i + 1; pages++; }
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
import { z } from "zod";
import { presetById } from "../presets";

export const formatConfigSchema = z.object({
  /** Preset pedoman kampus yang dipakai mengisi pengaturan ("umum" = bawaan). */
  presetId: z.string().min(1),
  docType: z.enum(["skripsi", "tesis", "proposal"]),
  paper: z.enum(["A4", "Letter"]),
  marginTop: z.number().min(1).max(6),
  marginBottom: z.number().min(1).max(6),
  marginLeft: z.number().min(1).max(6),
  marginRight: z.number().min(1).max(6),
  fontFamily: z.enum(["Times New Roman", "Arial", "Calibri"]),
  fontSize: z.number().min(10).max(14),
  lineSpacing: z.number().min(1).max(2.5),
  /**
   * auto  = kelipatan (spasi 2 pada huruf 12 ≈ 27,6 pt)
   * exact = nilai pasti (spasi 2 = exactly 24 pt, spasi 1 = exactly 12 pt)
   */
  lineRule: z.enum(["auto", "exact"]),
  indent: z.number().min(0).max(3),
  chapterUppercase: z.boolean(),
  /** "bab" = "BAB I ↵ PENDAHULUAN", "roman" = "I ↵ PENDAHULUAN". */
  chapterLabel: z.enum(["bab", "roman"]),
  /** Spasi baris judul bab/subbab: tunggal, atau sama dengan isi. */
  headingLineSpacing: z.enum(["single", "body"]),
  /** Jarak setelah judul bab, dalam satuan baris isi. */
  chapterSpaceAfter: z.number().min(0).max(4),

  /** Ukuran font judul topik di sampul (pt). */
  coverTitleSize: z.number().min(10).max(24),
  /** Ukuran font identitas kampus, kota, dan tahun di sampul (pt). */
  coverInstitutionSize: z.number().min(10).max(24),
  /** Identitas kampus ditampilkan huruf kapital semua. */
  coverInstitutionUppercase: z.boolean(),
  /** Identitas kampus dicetak tebal. */
  coverInstitutionBold: z.boolean(),
  /** Judul sampul ditampilkan huruf kapital semua. */
  coverTitleUppercase: z.boolean(),
  /** Nama penulis di sampul ditampilkan huruf kapital semua. */
  coverNameUppercase: z.boolean(),
  /** Baris NIM/NPM di sampul dicetak tebal. */
  coverIdentityBold: z.boolean(),
  /** Tambahkan "Oleh:" di atas nama bila sampul asli tidak memilikinya. */
  coverAuthorLabel: z.boolean(),
  /** Judul terjemahan (Inggris) sama besar dengan judul utama. */
  coverTranslationSameSize: z.boolean(),

  /**
   * Bagaimana memperlakukan daftar isi / tabel / gambar yang sudah ada.
   * - keep    : biarkan apa adanya (default — paling aman, isi dokumen tidak diubah)
   * - rebuild : ganti isinya dengan field otomatis Word
   */
  tocMode: z.enum(["keep", "rebuild"]),
  /**
   * Bentuk daftar isi / tabel / gambar / lampiran:
   * - word    : field TOC bawaan Word
   * - columns : berkolom (kepala "Bab … Halaman" / "Nomor … Halaman", baris bab tanpa nomor halaman)
   */
  listStyle: z.enum(["word", "columns"]),

  /** Rapikan halaman kosong akibat paragraf kosong beruntun / break ganda. */
  cleanupBlankPages: z.boolean(),

  /** Rapikan tata letak tabel (lebar, border, header berulang, perataan). */
  normalizeTables: z.boolean(),
  tableBorders: z.enum(["all", "minimal", "keep"]),
  /** Ukuran font isi tabel (pt). */
  tableFontSize: z.number().min(8).max(14),
  /** Perataan kolom angka di tabel. */
  tableNumericAlign: z.enum(["center", "right"]),

  /** Ukuran font judul tabel & gambar (pt). */
  captionFontSize: z.number().min(8).max(14),
  tableCaptionAlign: z.enum(["left", "center"]),
  figureCaptionAlign: z.enum(["center", "left"]),
  /** Penomoran tabel/gambar: per bab ("Tabel 2.1") atau berurutan ("Tabel 3"). */
  captionNumbering: z.enum(["chapter", "global"]),
  /** Pemisah setelah nomor: "Tabel 2.1 Judul" (space) atau "Tabel 3. Judul" (period). */
  captionSeparator: z.enum(["space", "period"]),
  /** Label keterangan gambar: ikuti dokumen, atau paksa "Gambar"/"Ilustrasi". */
  figureLabel: z.enum(["auto", "Gambar", "Ilustrasi"]),

  /** Penanda daftar: diseragamkan ("1.", "a.") atau bentuk asli dipertahankan ("(1)", "1)"). */
  listMarkers: z.enum(["uniform", "keep"]),
  /**
   * Tata letak daftar & judul berpenanda:
   * - nested : tiap tingkat makin menjorok (umum)
   * - flat   : semua penanda di margin kiri, teks sejajar indentasi alinea (mis. UNPAD FAPET)
   */
  listLayout: z.enum(["nested", "flat"]),

  /** Alinea abstrak memakai indentasi baris pertama. */
  abstractIndent: z.boolean(),
  /** Jarak antara teks abstrak dan baris "Kata Kunci"/"Keywords", dalam baris (null = bawaan formatter). */
  keywordGapLines: z.number().min(0).max(5).nullable(),

  /** Pertahankan penomoran halaman yang sudah ada di dokumen. */
  pageNumbers: z.enum(["none", "bottom-center", "bottom-right", "keep"]),
  /** Romawi untuk bagian awal, arab mulai BAB I. */
  romanFrontMatter: z.boolean(),
  /**
   * Halaman awal yang mulai diberi nomor:
   * - first   : sejak halaman setelah sampul
   * - preface : sejak Kata Pengantar (pengesahan, pernyataan, dsb. tanpa nomor)
   */
  frontNumberingFrom: z.enum(["first", "preface"]),

  citationStyle: z.enum(["APA", "IEEE", "Harvard", "Vancouver"]),
  /** Indentasi gantung entri daftar pustaka (cm). */
  referenceHanging: z.number().min(0.5).max(2),
  /** Satu baris jarak antarentri daftar pustaka. */
  referenceGap: z.boolean(),

  /** Hanya dipakai pada jalur PDF / rebuild. */
  listFigures: z.boolean(),
  listTables: z.boolean(),
});

export type FormatConfig = z.infer<typeof formatConfigSchema>;

export const defaultFormatConfig: FormatConfig = {
  presetId: "umum",
  docType: "skripsi",
  paper: "A4",
  marginTop: 3, marginBottom: 3, marginLeft: 4, marginRight: 3,
  fontFamily: "Times New Roman", fontSize: 12, lineSpacing: 1.5, lineRule: "auto", indent: 1.25,
  chapterUppercase: true,
  chapterLabel: "bab",
  headingLineSpacing: "single",
  chapterSpaceAfter: 1,
  coverTitleSize: 14,
  coverInstitutionSize: 16,
  coverInstitutionUppercase: true,
  coverInstitutionBold: true,
  coverTitleUppercase: true,
  coverNameUppercase: true,
  coverIdentityBold: false,
  coverAuthorLabel: true,
  coverTranslationSameSize: false,
  tocMode: "keep",
  listStyle: "word",
  cleanupBlankPages: true,
  normalizeTables: true,
  tableBorders: "minimal",
  tableFontSize: 11,
  tableNumericAlign: "center",
  captionFontSize: 11,
  tableCaptionAlign: "left",
  figureCaptionAlign: "center",
  captionNumbering: "chapter",
  captionSeparator: "space",
  figureLabel: "auto",
  listMarkers: "uniform",
  listLayout: "nested",
  abstractIndent: true,
  keywordGapLines: null,
  pageNumbers: "bottom-center",
  romanFrontMatter: true,
  frontNumberingFrom: "first",
  citationStyle: "APA",
  referenceHanging: 1,
  referenceGap: true,
  listFigures: false,
  listTables: false,
};

/** Nilai bawaan + nilai preset. */
export function configForPreset(id: string): FormatConfig {
  const preset = presetById(id);
  return { ...defaultFormatConfig, ...preset.config, presetId: preset.id };
}

/**
 * Pengurai sisi server. Semua field sengaja WAJIB di skema agar tipe input dan
 * output zod identik — `.default()` membuat keduanya berbeda dan `zodResolver`
 * tidak lagi cocok dengan `useForm<FormatConfig>`. Kelonggaran untuk payload
 * lama ditangani di sini: default ← preset ← isian pengguna.
 */
export function parseFormatConfig(input: unknown): FormatConfig {
  const raw = typeof input === "string" ? JSON.parse(input) : input;
  const obj = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
  const preset = presetById(typeof obj.presetId === "string" ? obj.presetId : "umum");
  return formatConfigSchema.parse({ ...defaultFormatConfig, ...preset.config, ...obj, presetId: preset.id });
}
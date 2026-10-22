import { z } from "zod";

export const formatConfigSchema = z.object({
  docType: z.enum(["skripsi", "tesis", "proposal"]),
  paper: z.enum(["A4", "Letter"]),
  marginTop: z.number().min(1).max(6),
  marginBottom: z.number().min(1).max(6),
  marginLeft: z.number().min(1).max(6),
  marginRight: z.number().min(1).max(6),
  fontFamily: z.enum(["Times New Roman", "Arial", "Calibri"]),
  fontSize: z.number().min(10).max(14),
  lineSpacing: z.number().min(1).max(2.5),
  indent: z.number().min(0).max(3),
  chapterUppercase: z.boolean(),

    /** Ukuran font judul topik di sampul (pt). */
  coverTitleSize: z.number().min(10).max(24),
  /** Ukuran font identitas kampus, kota, dan tahun di sampul (pt). */
  coverInstitutionSize: z.number().min(10).max(24),
  /** Identitas kampus ditampilkan huruf kapital semua. */
  coverInstitutionUppercase: z.boolean(),

  /**
   * Bagaimana memperlakukan daftar isi / tabel / gambar yang sudah ada.
   * - keep    : biarkan apa adanya (default — paling aman, isi dokumen tidak diubah)
   * - rebuild : ganti isinya dengan field otomatis Word
   */
  tocMode: z.enum(["keep", "rebuild"]),

  /** Rapikan halaman kosong akibat paragraf kosong beruntun / break ganda. */
  cleanupBlankPages: z.boolean(),

  /** Rapikan tata letak tabel (lebar, border, header berulang, perataan). */
  normalizeTables: z.boolean(),
  tableBorders: z.enum(["all", "minimal", "keep"]),

  /** Pertahankan penomoran halaman yang sudah ada di dokumen. */
  pageNumbers: z.enum(["none", "bottom-center", "bottom-right", "keep"]),
  /** Romawi untuk bagian awal, arab mulai BAB I. */
  romanFrontMatter: z.boolean(),

  citationStyle: z.enum(["APA", "IEEE", "Harvard", "Vancouver"]),

  /** Hanya dipakai pada jalur PDF / rebuild. */
  listFigures: z.boolean(),
  listTables: z.boolean(),
});

export type FormatConfig = z.infer<typeof formatConfigSchema>;

export const defaultFormatConfig: FormatConfig = {
  docType: "skripsi",
  paper: "A4",
  marginTop: 3, marginBottom: 3, marginLeft: 4, marginRight: 3,
  fontFamily: "Times New Roman", fontSize: 12, lineSpacing: 1.5, indent: 1.25,
  chapterUppercase: true,
    coverTitleSize: 14,
  coverInstitutionSize: 16,
  coverInstitutionUppercase: true,
  tocMode: "keep",
  cleanupBlankPages: true,
  normalizeTables: true,
  tableBorders: "minimal",
  pageNumbers: "bottom-center",
  romanFrontMatter: true,
  citationStyle: "APA",
  listFigures: false,
  listTables: false,
};

/**
 * Pengurai sisi server. Semua field sengaja WAJIB di skema agar tipe input dan
 * output zod identik — `.default()` membuat keduanya berbeda dan `zodResolver`
 * tidak lagi cocok dengan `useForm<FormatConfig>`. Kelonggaran untuk payload
 * lama ditangani di sini, bukan di skema.
 */
export function parseFormatConfig(input: unknown): FormatConfig {
  const raw = typeof input === "string" ? JSON.parse(input) : input;
  const obj = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
  return formatConfigSchema.parse({ ...defaultFormatConfig, ...obj });
}
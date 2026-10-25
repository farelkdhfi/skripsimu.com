/**
 * LAPISAN ATURAN FORMAT.
 *
 * Modul ini tidak tahu apa pun tentang isi dokumen — ia hanya memetakan
 * (peran, zona, petunjuk tata letak) menjadi properti format. Preset pedoman
 * kampus cukup mengisi `FormatConfig`; lapisan struktur dan lapisan OOXML tidak
 * perlu diubah.
 *
 * Prinsip utama: bedakan yang BERANTAKAN dari yang DISENGAJA.
 *  - Paragraf panjang rata kiri tanpa indentasi → berantakan, dirapikan.
 *  - Tanggal rata kanan, nama rata tengah, kolom tanda tangan bertab
 *    → disengaja, dipertahankan.
 */

import type { FormatConfig } from "../schemas/format-config";
import type { LayoutHint, Role, Zone } from "../structure/roles";

export const CM = 567;        // twips per cm
export const LINE = 240;      // twips per baris tunggal

export type Align = "left" | "center" | "right" | "both";

export type ParaFormat = {
    styleId?: string | null;
    /** `undefined` = pertahankan perataan asli. */
    align?: Align;
    firstLine?: number;
    left?: number;
    right?: number;
    hanging?: number;
    /** `undefined` = pertahankan jarak asli. */
    before?: number;
    after?: number;
    line?: number;
    keepNext?: boolean;
    keepLines?: boolean;
    pageBreakBefore?: boolean;
    /** Jangan sentuh indentasi sama sekali (paragraf bertab / berdaftar). */
    keepIndent?: boolean;
    /** Buang tab di awal paragraf (indentasi manual yang diganti indentasi baris pertama). */
    stripLeadingTabs?: boolean;
    /**
     * Entri daftar isi: posisi tab kanan (twips) tempat nomor halaman diletakkan,
     * dengan leader titik. Menyeragamkan leader campuran (titik, strip, garis).
     */
    tocTab?: number;
};

export type RunFormat = {
    font: string;
    /** half-points */
    size: number;
    /** `undefined` berarti biarkan apa adanya — penting agar penekanan asli tidak hilang. */
    bold?: boolean;
    /** `true` = tampilkan kapital, `false` = buang kapital paksa, `undefined` = biarkan. */
    caps?: boolean;
    /** Warna teks. `null` = buang warna langsung (kembali ke hitam bawaan). */
    color?: string | null;
    /** `false` = buang garis bawah (judul diseragamkan tanpa garis bawah). */
    underline?: false;
    /** `false` = buang cetak miring (judul & keterangan diseragamkan tegak). */
    italic?: false;
    /**
     * Ukuran mengikuti hierarki asli: run yang aslinya lebih besar dari teks isi
     * diberi ukuran `large`, sisanya `size`. Dipakai di sampul agar judul tetap
     * menonjol tanpa membiarkan campuran 11/14/20 pt yang berantakan.
     */
    hierarchy?: { large: number; threshold: number };
};

export type PageFormat = {
    widthTwips: number;
    heightTwips: number;
    margin: { top: number; bottom: number; left: number; right: number };
};

export type NumberingFormat = "decimal" | "lowerRoman";

export type TableFormat = {
    widthPct: number | null;
    borders: "all" | "minimal" | "keep";
    borderSize: number;
    cellMargin: number;
    repeatHeaderRow: boolean;
    boldHeaderRow: boolean;
    headerAlign: Align;
    /** Perataan kolom yang isinya angka. */
    numericAlign: Align;
    fontSize: number;
    lineSpacing: number;
    /** Buang warna latar sel (biru, kuning, dst) — tabel akademik hitam-putih. */
    clearShading: boolean;
    /** Ukuran font catatan di bawah tabel (half-points). */
    noteSize: number;
};

export type CoverStyle = {
    /** half-points */
    titleHalf: number;
    textHalf: number;
    institutionHalf: number;
    /** Sisi terpanjang logo sampul (twips). */
    logoMaxTwips: number;
    /** Identitas kampus ditampilkan kapital (w:caps, teks asli tidak diubah). */
    institutionCaps: boolean;
    institutionBold: boolean;
    titleCaps: boolean;
    nameCaps: boolean;
    /** Baris NIM/NPM tebal. */
    idBold: boolean;
    /** Tambah "Oleh:" bila sampul asli tidak memilikinya. */
    authorLabel: boolean;
    /** Judul terjemahan sama besar dengan judul utama. */
    translationSame: boolean;
};

export interface FormatRules {
    id: string;
    name: string;
    page(zone: Zone): PageFormat;
    /** Format nomor halaman untuk zona; `null` = jangan ubah. */
    numbering(zone: Zone): NumberingFormat | null;
    paragraph(role: Role, zone: Zone, hint: LayoutHint): ParaFormat;
    run(role: Role, zone: Zone, hint: LayoutHint): RunFormat;
    table(): TableFormat;
    cover(): CoverStyle;
}

const PAPER: Record<FormatConfig["paper"], { widthTwips: number; heightTwips: number }> = {
    A4: { widthTwips: 11906, heightTwips: 16838 },
    Letter: { widthTwips: 12240, heightTwips: 15840 },
};

/** Lebar gantungan butir daftar ketikan (0,63 cm, bawaan Word). */
const LIST_HANG = 357;

/** Paragraf sepanjang ini dianggap paragraf isi, bukan baris tata letak. */
const PROSE_MIN_LENGTH = 150;

/**
 * Aturan skripsi Indonesia, dikendalikan sepenuhnya oleh `FormatConfig`
 * (termasuk nilai dari preset kampus).
 */
export function createThesisRules(c: FormatConfig): FormatRules {
    const size = Math.round(c.fontSize * 2);
    const tableSize = Math.round(c.tableFontSize * 2);
    const captionSize = Math.round(c.captionFontSize * 2);
    const bodyLine = Math.round(c.lineSpacing * LINE);
    const firstLine = Math.round(c.indent * CM);
    const usableWidth = PAPER[c.paper].widthTwips - Math.round(c.marginLeft * CM) - Math.round(c.marginRight * CM);
    const single = LINE;

    // Judul ikut spasi isi (mis. UNPAD: semua baris exactly 24 pt) atau spasi tunggal.
    const headLine = c.headingLineSpacing === "body" ? bodyLine : single;
    const bodyGapHeadings = c.headingLineSpacing === "body";
    const chapterAfter = Math.round(c.chapterSpaceAfter * bodyLine);

    /** Paragraf isi standar: rata kiri-kanan, indentasi baris pertama. */
    const prose = (line: number): ParaFormat => ({
        align: "both", line, firstLine, left: 0, right: 0, before: 0, after: 0,
    });

    /** Abstrak: spasi tunggal; indentasi baris pertama sesuai pedoman. */
    const abstract = (): ParaFormat => ({ ...prose(single), ...(c.abstractIndent ? {} : { firstLine: 0 }) });

    /**
     * Baris tata letak (tanggal, nama penanda tangan, "Mengetahui,"): perataan
     * dan jarak asli dipertahankan, hanya indentasi baris pertama yang dibuang
     * karena itu yang membuat baris rata tengah/kanan bergeser.
     */
    const layoutLine = (): ParaFormat => ({ firstLine: 0 });

    /** Paragraf bertab: jangan sentuh perataan maupun indentasi sama sekali. */
    const tabbed = (): ParaFormat => ({ keepIndent: true });

    return {
        id: c.presetId || "default-id",
        name: "Skripsi (Indonesia)",

        page() {
            return {
                ...PAPER[c.paper],
                margin: {
                    top: Math.round(c.marginTop * CM),
                    bottom: Math.round(c.marginBottom * CM),
                    left: Math.round(c.marginLeft * CM),
                    right: Math.round(c.marginRight * CM),
                },
            };
        },

        numbering(zone) {
            if (c.pageNumbers === "keep") return null;
            return c.romanFrontMatter && (zone === "cover" || zone === "front") ? "lowerRoman" : "decimal";
        },

        paragraph(role, zone, hint) {
            const intentional = hint.align === "center" || hint.align === "right";

            switch (role) {
                case "cover":
                    // Sampul: semua rata tengah, indentasi dinolkan (sumber utama baris
                    // yang "menggantung" ke samping), tetapi JARAK VERTIKAL ASLI DIPERTAHANKAN.
                    if (hint.tabCount > 0) return tabbed();
                    return { align: "center", left: 0, right: 0, firstLine: 0 };

                case "chapter_label":
                    return {
                        styleId: hint.pairedChapter ? "JudulBab" : "Heading1", align: "center", line: headLine,
                        left: 0, right: 0, firstLine: 0, before: 0, after: 0,
                        keepNext: true, keepLines: true, pageBreakBefore: true,
                    };

                case "chapter":
                case "front_heading":
                case "list_heading":
                case "reference_heading":
                case "appendix_heading":
                    return {
                        styleId: role === "chapter" && hint.pairedChapter ? "JudulBab" : "Heading1", align: "center", line: headLine,
                        left: 0, right: 0, firstLine: 0, before: 0, after: chapterAfter,
                        keepNext: true, keepLines: true, pageBreakBefore: true,
                    };

                case "section":
                    return {
                        styleId: "Heading2", align: "left", line: headLine, left: 0, right: 0, firstLine: 0,
                        before: bodyGapHeadings ? bodyLine : LINE,
                        after: bodyGapHeadings ? 0 : Math.round(LINE * 0.5),
                        keepNext: true, keepLines: true,
                    };

                case "subsection":
                    return {
                        styleId: "Heading3", align: "left", line: headLine, left: 0, right: 0, firstLine: 0,
                        before: bodyGapHeadings ? bodyLine : Math.round(LINE * 0.75),
                        after: bodyGapHeadings ? 0 : Math.round(LINE * 0.5),
                        keepNext: true, keepLines: true,
                    };

                case "list_entry":
                    // Entri daftar isi: tab stop + dot leader bawaan dokumen dipertahankan.
                    return {
                        align: "left", line: single, firstLine: 0, after: Math.round(LINE * 0.25),
                        keepIndent: true, tocTab: usableWidth,
                    };

                case "figure":
                    return {
                        align: "center", line: single, left: 0, right: 0, firstLine: 0,
                        before: Math.round(LINE * 0.5), after: 0, keepNext: true,
                    };

                case "figure_caption":
                    return {
                        align: c.figureCaptionAlign, line: single, left: 0, right: 0, firstLine: 0,
                        before: Math.round(LINE * 0.25), after: LINE,
                    };

                case "table_caption":
                    return {
                        align: c.tableCaptionAlign, line: single, left: 0, right: 0, firstLine: 0,
                        before: LINE, after: Math.round(LINE * 0.25), keepNext: true,
                    };

                case "reference": {
                    const hang = Math.round(c.referenceHanging * CM);
                    return {
                        align: "both", line: single, firstLine: 0, left: hang, hanging: hang,
                        before: 0, after: c.referenceGap ? LINE : 0,
                    };
                }

                case "empty":
                    return { keepIndent: true };

                default: {
                    // Butir daftar ketikan: penanda di posisi indentasi paragraf, baris
                    // lanjutan sejajar teks (hanging indent) — bukan kembali ke margin.
                    if (hint.listMarker && zone !== "cover") {
                        return {
                            align: "both", line: hint.frontKind === "abstract" ? single : bodyLine,
                            left: firstLine + LIST_HANG, hanging: LIST_HANG, right: 0, before: 0, after: 0,
                            stripLeadingTabs: true,
                        };
                    }
                    if (hint.tabCount > 0) {
                        // Tab hanya di awal = indentasi ketikan pada paragraf isi.
                        if (hint.leadingTabsOnly && hint.length >= PROSE_MIN_LENGTH) {
                            return { ...(hint.frontKind === "abstract" ? abstract() : prose(bodyLine)), stripLeadingTabs: true };
                        }
                        return tabbed();
                    }

                    // Bagian awal & lampiran: hanya paragraf panjang yang dirapikan sebagai prosa.
                    // Baris pendek (tanggal, nama, NIP, "Mengetahui,", alamat surat "Yth. …"/"di …",
                    // butir isian kuesioner) dipertahankan perataan & jaraknya, tanpa indentasi.
                    if (zone === "front" || zone === "cover" || zone === "back") {
                        if (hint.length < PROSE_MIN_LENGTH) return layoutLine();
                        return hint.frontKind === "abstract" ? abstract() : prose(bodyLine);
                    }

                    // Isi: baris pendek yang sengaja rata tengah/kanan dipertahankan
                    // (mis. "Keterangan:", tanda tangan di lampiran, persamaan).
                    if (intentional && hint.length < PROSE_MIN_LENGTH) return layoutLine();
                    return prose(bodyLine);
                }
            }
        },

        run(role, _zone, hint) {
            switch (role) {
                case "cover":
                    return {
                        font: c.fontFamily, size, color: null,
                        hierarchy: { large: Math.max(size, Math.round(c.coverTitleSize * 2)), threshold: size },
                    };
                case "chapter":
                case "chapter_label":
                case "front_heading":
                case "list_heading":
                case "reference_heading":
                case "appendix_heading":
                    return { font: c.fontFamily, size, bold: true, caps: c.chapterUppercase, color: null, underline: false, italic: false };
                case "section":
                case "subsection":
                    // `caps: false` membuang kapital paksa ("1.4. BATASAN MASALAH" via w:caps).
                    return { font: c.fontFamily, size, bold: true, caps: false, color: null, underline: false, italic: false };
                case "figure_caption":
                    return { font: c.fontFamily, size: captionSize, caps: false, color: null, italic: false };
                case "table_caption":
                    return { font: c.fontFamily, size: captionSize, bold: false, caps: false, color: null, italic: false };
                case "list_entry":
                    return { font: c.fontFamily, size, color: null };
                default:
                    // Isi skripsi tidak memakai huruf kapital paksa maupun teks berwarna.
                    // Paragraf isi yang tebal SELURUHNYA dinormalkan; penekanan pada
                    // sebagian kata tetap dipertahankan.
                    return {
                        font: c.fontFamily, size, caps: false, color: null,
                        ...(hint.allBold && hint.length >= PROSE_MIN_LENGTH ? { bold: false } : {}),
                    };
            }
        },

        cover: () => ({
            titleHalf: Math.round(c.coverTitleSize * 2),
            textHalf: size,
            institutionHalf: Math.round(c.coverInstitutionSize * 2),
            institutionCaps: c.coverInstitutionUppercase,
            institutionBold: c.coverInstitutionBold,
            titleCaps: c.coverTitleUppercase,
            nameCaps: c.coverNameUppercase,
            idBold: c.coverIdentityBold,
            authorLabel: c.coverAuthorLabel,
            translationSame: c.coverTranslationSameSize,
            logoMaxTwips: Math.round(3.5 * CM),
        }),

        table: () => ({
            widthPct: c.normalizeTables ? 100 : null,
            borders: c.tableBorders,
            borderSize: 4,
            cellMargin: 85,
            repeatHeaderRow: c.normalizeTables,
            boldHeaderRow: c.normalizeTables,
            headerAlign: "center",
            numericAlign: c.tableNumericAlign,
            fontSize: tableSize,
            lineSpacing: LINE,
            clearShading: c.normalizeTables,
            noteSize: 20,
        }),
    };
}
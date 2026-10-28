/**
 * Bagian depan dari TEMPLATE DOCX preset (sampul sampai abstract).
 *
 * Template = berkas DOCX bagian depan yang sudah benar (disiapkan di belakang layar). Tata letak
 * dan format SELALU dari template; dari naskah mahasiswa hanya diambil TEKS (judul, nama, NPM,
 * pembimbing, isi kata pengantar, isi abstrak, …) dan tanda tangan/materai.
 */
import type { ParaNode } from "../docx/document-model";

export type FrontPageKind =
    | "sampulMuka"
    | "sampulKedua"
    | "pengesahan"
    | "pernyataan"
    | "kataPengantar"
    | "abstrak"
    | "abstract";

export const FRONT_PAGE_LABEL: Record<FrontPageKind, string> = {
    sampulMuka: "Sampul muka",
    sampulKedua: "Sampul kedua (bagian dalam)",
    pengesahan: "Lembar pengesahan",
    pernyataan: "Lembar pernyataan tidak melakukan plagiasi",
    kataPengantar: "Kata pengantar",
    abstrak: "Abstrak",
    abstract: "Abstract",
};

/** Potongan teks berformat: miring/subskrip/superskrip dipertahankan dari naskah. */
export type Seg = { t: string; it: boolean; sub: boolean; sup: boolean; b: boolean };
export type Rich = Seg[];

/** Gambar di naskah. `picture` = foto/scan (logo, tanda tangan, materai); selain itu bentuk/kotak. */
export type Pic = {
    /** Elemen gambar utuh (<w:drawing> atau <w:pict>). */
    xml: string;
    picture: boolean;
    cx: number;
    cy: number;
    embeds: string[];
    /** Objek yang tidak bisa dipakai ulang (gambar format lama, grafik, …): alasannya. */
    unsupported?: string;
};

export type Item =
    | { k: "line"; segs: Rich; text: string; raw: string; para: ParaNode; inTable: boolean }
    | { k: "blank"; para: ParaNode }
    | { k: "pic"; pic: Pic; para: ParaNode }
    | { k: "break" };

export type LineItem = Extract<Item, { k: "line" }>;
export type PicItem = Extract<Item, { k: "pic" }>;

export type FrontPage = {
    kind: FrontPageKind;
    items: Item[];
};

/** Data yang diisikan ke template. Nilai null = tidak ditemukan di naskah. */
export type FrontFields = {
    judul: Rich | null;
    judulInggris: Rich | null;
    nama: string | null;
    npm: string | null;
    tahun: string | null;
    pembimbingUtama: string | null;
    pembimbingAnggota: string | null;
    ketuaSidang: string | null;
    dekan: string | null;
    tanggalSidang: string | null;
    tempatTanggalPernyataan: string | null;
    tempatTanggalKataPengantar: string | null;
};

export type FrontFieldKey = keyof FrontFields;

export const FIELD_KEYS: FrontFieldKey[] = [
    "judul", "judulInggris", "nama", "npm", "tahun",
    "pembimbingUtama", "pembimbingAnggota", "ketuaSidang", "dekan", "tanggalSidang",
    "tempatTanggalPernyataan", "tempatTanggalKataPengantar",
];

/** Isian dari pengguna (form konfirmasi) — menimpa hasil pembacaan naskah. Judul berupa teks biasa. */
export type FrontFieldInput = Partial<Record<FrontFieldKey, string>>;

export const FIELD_LABEL: Record<FrontFieldKey, string> = {
    judul: "Judul",
    judulInggris: "Judul bahasa Inggris",
    nama: "Nama penulis",
    npm: "NPM",
    tahun: "Tahun",
    pembimbingUtama: "Pembimbing utama",
    pembimbingAnggota: "Pembimbing anggota",
    ketuaSidang: "Ketua panitia sidang",
    dekan: "Dekan",
    tanggalSidang: "Tanggal sidang",
    tempatTanggalPernyataan: "Tempat & tanggal lembar pernyataan",
    tempatTanggalKataPengantar: "Tempat & tanggal kata pengantar",
};

/**
 * Spesifikasi template satu kampus. Kampus baru cukup menyiapkan berkas DOCX bagian depan yang
 * benar + spesifikasi ini; kode penyusun tidak berubah.
 */
export type FrontSpec = {
    /** Berkas template, relatif terhadap folder `src/features/thesis-formatter/presets/templates`. */
    file: string;
    /** Jenis dokumen yang memakai template ini (FormatConfig.docType). */
    docTypes: ("skripsi" | "tesis" | "proposal")[];
    /** Isi section-section pertama berkas template, berurutan (satu section = satu halaman). */
    pages: FrontPageKind[];
    /**
     * Teks contoh di template yang diganti data mahasiswa. Dicocokkan tanpa memedulikan huruf
     * besar/kecil, spasi, dan tanda baca, jadi "CAIRAN RUMEN : IN VITRO" tetap cocok dengan
     * "Cairan Rumen: In vitro". Teks template lain (judul halaman, kalimat baku) tidak diubah.
     */
    sample: Partial<Record<FrontFieldKey, string>>;
    /** Label tetap paragraf kata kunci di template. */
    keywordLabels: { abstrak: string; abstract: string };
    /** Halaman yang gambarnya bagian template (logo). Gambar di halaman lain = tanda tangan/materai contoh. */
    fixedImagePages: FrontPageKind[];
    /** Batas ukuran tanda tangan yang disisipkan di atas nama penanda tangan (cm). */
    signature: { maxHeightCm: number; maxWidthCm: number };
    /** Pengisi data yang tidak ditemukan. */
    placeholder: string;
    /** Tempat untuk "tempat, tanggal" yang tidak ada di naskah (diisi otomatis dengan bulan & tahun). */
    defaultPlace: string;
    /** Judul halaman daftar yang disusun formatter sesudah template. */
    listHeadings: { contents: string; tables: string; figures: string; appendices: string };
};

export type FrontPreviewLine = { role: string; zone: string; text: string };
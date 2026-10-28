import type { FrontSpec } from "../front/types";

/**
 * Template bagian depan skripsi Fakultas Peternakan UNPAD.
 *
 * Berkas `templates/unpad-fapet-depan.docx` = bagian depan yang sudah benar (sampul sampai
 * abstract, satu section per halaman). Tata letak & format seluruhnya dari berkas itu; yang
 * diganti hanya teks contoh di bawah ini (data mahasiswa contoh) serta isi kata pengantar,
 * abstrak, dan abstract. Untuk memperbarui template: ganti berkas DOCX-nya, lalu sesuaikan
 * teks contoh bila datanya berubah.
 */
export const UNPAD_FAPET_FRONT: FrontSpec = {
  file: "unpad-fapet-depan.docx",
  docTypes: ["skripsi"],
  pages: ["sampulMuka", "sampulKedua", "pernyataan", "pengesahan", "kataPengantar", "abstrak", "abstract"],
  sample: {
    judul: "Pengaruh Kulit Buah Nanas (Ananas comosus) Dalam Ransum Terhadap Laktil-CoA, Akril-CoA dan Propionil-CoA Cairan Rumen: In vitro",
    judulInggris: "The Effect of Pineapple (Ananas comosus) Peels in Rations on Lactyl-CoA, Acryl-CoA, and Propionyl-CoA of Rumen Fluid: In vitro",
    nama: "Riki Mulyana",
    npm: "200110220087",
    tahun: "2026",
    pembimbingUtama: "Dr. Ir. Andi Mushawwir, S.Pt., MP., IPM.",
    pembimbingAnggota: "Novi Mayasari, S.Pt., M.Sc., Ph.D.",
    ketuaSidang: "Dr. Eka Wulandari, S.Si., M.Si",
    dekan: "Prof. Dr. Ir. Rahmat Hidayat, S.Pt., M.Si., IPM",
    tanggalSidang: "30 Juni 2026",
    tempatTanggalPernyataan: "Jatinangor, Juni 2026",
    tempatTanggalKataPengantar: "Sumedang, 8 Juni 2026",
  },
  keywordLabels: { abstrak: "Kata Kunci:", abstract: "Keywords:" },
  fixedImagePages: ["sampulMuka", "sampulKedua"],
  signature: { maxHeightCm: 1.6, maxWidthCm: 4.5 },
  placeholder: "……………………………",
  defaultPlace: "Jatinangor",
  listHeadings: { contents: "DAFTAR ISI", tables: "DAFTAR TABEL", figures: "DAFTAR ILUSTRASI", appendices: "DAFTAR LAMPIRAN" },
};
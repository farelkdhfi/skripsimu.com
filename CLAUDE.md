# AI Thesis Formatter — panduan kerja untuk Claude Code

## Cara kerja dengan pemilik project
- Bahasa: Indonesia santai ("gua/lu"). Jawaban ringkas, langsung ke inti.
- Kalau diminta "jangan dulu ngoding" / "konfirmasi dulu": jelaskan pemahaman & rencana saja, tanpa mengubah file.
- Jangan pernah bilang "sudah robust/aman" sebelum lolos `npm run test:docs` DAN halaman penting sudah dicek lewat render (lihat "Definisi selesai").
- Perbaikan harus menyasar akar masalah dan berlaku untuk dokumen siapa pun, bukan menambal satu dokumen.
- Hemat token: baca file seperlunya (Grep dulu, lalu Read bagian yang relevan). Jangan membaca `node_modules/`, `.next/`, `test-docs/__out__/`, atau file DOCX mentah kecuali perlu.

## Peta kode (src/features/thesis-formatter/)
- `docx/transform.ts` — jalur utama: parse → outline → modul-modul → tulis ulang paket → verifikasi.
- `docx/document-model.ts` — model dokumen (ParaNode/TableNode/CellNode), `flatten`, `cellParagraphs`, `parseBlocks`.
- `docx/xml.ts` — helper OOXML (`withProps`, `mergeProps`, urutan PPR/RPR/TBLPR/TCPR).
- `structure/roles.ts`, `structure/outline.ts` — peran (judul, daftar, keterangan, …) dan zona (cover/front/main/back).
- `rules/format-rules.ts` — peta peran+zona → format. Preset di `presets/` (mis. `unpad-fapet.ts`), skema di `schemas/format-config.ts`.
- Modul penting: `apply-format.ts`, `tables.ts` (normalisasi & catatan tabel), `table-lines.ts` (garis tabel), `captions.ts`, `column-lists.ts` (Daftar Isi/Tabel/Ilustrasi/Lampiran berkolom), `signature.ts` (blok tanda tangan), `statement.ts` (lembar pernyataan), `float-inline.ts` (gambar melayang → inline), `cover-fit.ts` (pengukuran halaman: sampul, ruang tanda tangan, pernyataan), `link-color.ts`, `label-lines.ts`, `verify.ts`.
- Uji: `scripts/test-docs.ts`, kasus di `test-docs/cases.json`, snapshot di `test-docs/__snapshots__/`, hasil di `test-docs/__out__/`.

## Data pribadi
`test-docs/` berisi skripsi asli (nama, NPM, tanda tangan, foto). Jangan di-commit, jangan disalin ke tempat lain, jangan dikutip panjang di jawaban.

## Definisi selesai (wajib sebelum melapor)
1. `npx tsc --noEmit` bersih.
2. `npm run test:docs` — semua dokumen OK; perubahan snapshot dijelaskan satu per satu (disengaja atau regresi).
3. Render hasil ke PDF (LibreOffice: `soffice --headless --convert-to pdf <file>`) dan periksa halaman: sampul, pengesahan, pernyataan, Daftar Isi/Tabel/Ilustrasi, satu tabel bertingkat, satu halaman bergambar, Daftar Pustaka, lampiran.
4. Laporkan apa yang SUDAH dicek dan apa yang BELUM.

## Rencana kerja berikutnya (urut)
1. Pengecekan berbasis render di `test:docs`: sampul/pengesahan/pernyataan tepat 1 halaman; tidak ada halaman kosong; teks/gambar tidak keluar margin atau kotak; Daftar Isi hanya sekali & tiap entri lengkap; keterangan di Daftar Tabel/Ilustrasi tidak terpotong; gambar tidak menimpa teks.
2. Halaman awal disusun dari DATA: kenali jenis halaman → ambil data (judul, nama, NPM, kelompok tanda tangan Menyetujui/Mengesahkan beserta peran, nama, NIP, gambar) → susun ulang dari template preset, teks persis milik penulis. Bila ekstraksi tidak yakin → halaman dibiarkan asli + catatan.
3. Kunci isi lampiran: isi tiap lampiran disisihkan sebelum diproses dan dikembalikan utuh; hanya judul lampiran, margin/ukuran section, gambar & tabel yang melebihi area teks, dan indentasi negatif yang dibenahi. Biodata tetap diformat.
4. Perbaikan yang sudah dianalisis:
   - Verifikasi: `<w:tab/>` dan ganti baris dibaca sebagai spasi; objek yang sengaja dibuang (kotak bingkai) dilaporkan ke verifikasi.
   - Daftar Isi: baris mirip judul (mis. "LEMBAR PENGESAHAN" tanpa nomor halaman) hanya menutup blok daftar bila baris sesudahnya bukan entri; DAFTAR TABEL/ILUSTRASI/LAMPIRAN tetap memulai daftar sendiri; Daftar Isi lama harus terganti seluruhnya.
   - Keterangan tabel/gambar yang dipotong Enter ("… terhadap ↵ (spasi) Maltosa") digabung kembali.
   - Pengesahan: kolom menurut judul kelompok (Menyetujui → kanan, Mengesahkan/Mengetahui → kiri), peran hanya bila tanpa judul kelompok; gambar tanda tangan masuk blok dan ditempatkan di ruang tanda tangan nama di bawahnya; baris yang hanya berisi blok kiri selebar penuh; halaman diukur & ruang tanda tangan dikecilkan bertahap (min ±1,5 cm) agar 1 halaman.
   - Pernyataan: bingkai buatan penulis dalam 4 bentuk (tabel 1 sel, shape persegi, border paragraf, page border) dibuang sebelum kotak seragam; blok satu penanda tangan tetap paragraf; gambar melayang di dalam kotak dijaga tetap di dalam; paragraf penutup kotak setinggi 1 pt.
   - Indentasi gantung asli dipertahankan (kata kunci mulai di margin); Tab ketikan di awal paragraf isi dibuang.
   - Catatan tabel: butir bernomor yang menempel pada catatan lain tetap catatan, kecuali setelah "Sumber:" atau bila meneruskan daftar dari atas tabel.
   - Gambar yang dijadikan inline diletakkan rata tengah.
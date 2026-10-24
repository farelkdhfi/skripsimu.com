import { z } from "zod";
import { askJson } from "@/features/ai/services/openrouter";
import type { CoverFields } from "../docx/cover-template";

const S = z.string().nullable().optional().catch(null);
const schema = z.object({
  docType: S, title: S, purpose: S, authorLabel: S, name: S, idLabel: S, studentId: S,
  program: S, faculty: S, university: S, city: S, year: S,
}).catch({});

const system = `Kamu membaca teks halaman sampul skripsi/proposal/tesis (Indonesia atau dwibahasa) dan memetakan isinya ke field.
Aturan WAJIB:
- Setiap nilai harus DISALIN PERSIS dari teks sampul. Boleh menggabungkan beberapa baris berurutan dengan satu spasi. Jangan menerjemahkan, meringkas, mengubah huruf besar/kecil, memperbaiki ejaan, atau mengarang.
- Field yang tidak ada di teks diisi null.
- docType: SEMUA baris jenis dokumen digabung dengan spasi, mis. "SKRIPSI", "SKRIPSI / THESIS", "USULAN PENELITIAN TUGAS AKHIR".
- program, faculty, university: masing-masing bagian utuh termasuk katanya (mis. "PROGRAM STUDI INFORMATIKA", "FAKULTAS TEKNIK", "UNIVERSITAS SILIWANGI"). Jika ketiganya tertulis dalam satu baris, salin MASING-MASING bagiannya secara terpisah.
- title: judul karya tulis lengkap termasuk subjudul/terjemahannya, TANPA label "Judul :" / "Title :".
- purpose: kalimat keterangan "Diajukan …"/"Untuk memenuhi …" beserta terjemahan Inggrisnya bila ada.
- authorLabel: HANYA label seperti "Oleh:" atau "Disusun oleh:". Label "Nama" BUKAN authorLabel.
- name: nama mahasiswa saja, tanpa label "Nama :" dan tanpa NIM.
- idLabel: "NIM"/"NPM"/"NRP". studentId: nomornya saja.
- city: nama kota. year: tahun (boleh dengan bulan).
- Pembimbing, konsentrasi, dan keterangan lain BUKAN field mana pun.
Balas hanya JSON dengan kunci: docType, title, purpose, authorLabel, name, idLabel, studentId, program, faculty, university, city, year.`;

export async function extractCoverAI(lines: string[]): Promise<CoverFields | null> {
  if (!lines.length) return null;
  try {
    return await askJson({
      system,
      user: lines.join("\n"),
      schema,
      maxTokens: 800,
      timeoutMs: 20_000,
      maxAttempts: 2,
    });
  } catch (e) {
    console.error("[extractCoverAI]", e instanceof Error ? e.message : e);
    return null;
  }
}
import type { FormatConfig } from "../schemas/format-config";
import { UNPAD_FAPET } from "./unpad-fapet";

export type CampusPreset = {
  id: string;
  /** Kode di nama file hasil: "skripsi-FORMATTED-UNPAD-FAPET.docx" ("" = tanpa kode). */
  code: string;
  name: string;
  /** Pedoman yang menjadi sumber preset. */
  source: string;
  /** Penjelasan singkat untuk pengguna: apa yang diatur & apa yang perlu diperiksa manual. */
  notes: string[];
  config: Partial<Omit<FormatConfig, "presetId" | "docType">>;
};

export const UMUM: CampusPreset = {
  id: "umum",
  code: "",
  name: "Umum (pedoman skripsi Indonesia)",
  source: "Aturan umum: A4, margin 4-3-3-3, Times New Roman 12, spasi 1,5.",
  notes: [],
  config: {},
};

export const PRESETS: CampusPreset[] = [UMUM, UNPAD_FAPET];

export function presetById(id: string | null | undefined): CampusPreset {
  return PRESETS.find((p) => p.id === id) ?? UMUM;
}
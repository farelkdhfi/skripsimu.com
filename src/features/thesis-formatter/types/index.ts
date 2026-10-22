import type { Role, Zone } from "../structure/roles";

export type { Role, Zone };

export type Block = { role: Role; text: string };

export type DocStats = {
  chapters: number;
  sections: number;
  subsections: number;
  figures: number;
  figureCaptions: number;
  tableCaptions: number;
  tables: number;
  references: number;
  frontMatter: number;
  appendices: number;
  referenceHeadings: number;
};

export type IntegrityCheck = {
  label: string;
  before: number | string;
  after: number | string;
  ok: boolean;
};

export type FormatResult = {
  fileName: string;
  docxBase64: string;
  source: "docx" | "pdf";
  /** `preserve` = dokumen asli disunting di tempat; `rebuild` = disusun ulang (PDF). */
  mode: "preserve" | "rebuild";
  preview: { role: Role; zone: Zone; text: string }[];
  stats: DocStats;
  /** Hasil perbandingan dokumen asli vs hasil, ditampilkan apa adanya. */
  integrity: IntegrityCheck[];
  changes: string[];
  validation: { level: "ok" | "warn" | "error"; message: string }[];
  recommendations: string[];
};
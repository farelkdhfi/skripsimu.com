export type Role =
  | "cover" | "chapter" | "section" | "subsection"
  | "figure_caption" | "table_caption" | "body" | "reference" | "skip";
export type Block = { role: Role; text: string };

export type FormatResult = {
  fileName: string;
  docxBase64: string;
  source: "docx" | "pdf";
  preview: Block[];
  stats: { chapters: number; sections: number; subsections: number; figures: number; tables: number; references: number; removed: number; wordsBefore: number; wordsAfter: number };
  changes: string[];
  validation: { level: "ok" | "warn" | "error"; message: string }[];
  recommendations: string[];
};
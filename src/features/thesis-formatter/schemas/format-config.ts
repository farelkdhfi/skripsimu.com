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
  toc: z.boolean(),
  listFigures: z.boolean(),
  listTables: z.boolean(),
  pageNumbers: z.enum(["none", "bottom-center", "bottom-right"]),
  citationStyle: z.enum(["APA", "IEEE", "Harvard", "Vancouver"]),
});
export type FormatConfig = z.infer<typeof formatConfigSchema>;

export const defaultFormatConfig: FormatConfig = {
  docType: "skripsi", paper: "A4",
  marginTop: 3, marginBottom: 3, marginLeft: 4, marginRight: 3,
  fontFamily: "Times New Roman", fontSize: 12, lineSpacing: 1.5, indent: 1.25,
  chapterUppercase: true, toc: true, listFigures: true, listTables: true,
  pageNumbers: "bottom-center", citationStyle: "APA",
};